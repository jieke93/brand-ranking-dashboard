#!/usr/bin/env python3
"""개인용 설문 대시보드 (Google Forms 스타일 최소 구현)

- 폼 정의(JSON)로 질문을 구성
- Streamlit에서 설문 작성/제출
- 로컬(JSONL)로 응답 저장
- 응답 목록 조회 및 CSV 다운로드

실행:
  streamlit run personal_form_dashboard.py
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import plotly.express as px
from collections import Counter
import re


WORK_DIR = os.path.dirname(os.path.abspath(__file__))
FORMS_DIR = os.path.join(WORK_DIR, "forms")
RESPONSES_DIR = os.path.join(WORK_DIR, "form_responses")


SUPPORTED_TYPES = {
    "short_text",
    "paragraph",
    "multiple_choice",
    "checkboxes",
    "dropdown",
    "linear_scale",
    "number",
    "date",
    "time",
    "file_upload",
    "section",
}

CHOICE_TYPES = {"multiple_choice", "checkboxes", "dropdown"}


@dataclass(frozen=True)
class Question:
    id: str
    type: str
    title: str
    description: str = ""
    required: bool = False
    options: Tuple[str, ...] = ()
    min: Optional[int] = None
    max: Optional[int] = None
    min_label: str = ""
    max_label: str = ""
    section: Optional[str] = None
    file_types: Tuple[str, ...] = ()
    max_file_size_mb: Optional[int] = None


@dataclass(frozen=True)
class FormConfig:
    form_id: str
    title: str
    description: str
    questions: Tuple[Question, ...]


def _ensure_dirs() -> None:
    os.makedirs(FORMS_DIR, exist_ok=True)
    os.makedirs(RESPONSES_DIR, exist_ok=True)


def list_form_files() -> List[str]:
    _ensure_dirs()
    files = []
    for name in sorted(os.listdir(FORMS_DIR)):
        if name.lower().endswith(".json"):
            files.append(name)
    return files


def load_form_raw(file_name: str) -> Dict[str, Any]:
    path = os.path.join(FORMS_DIR, file_name)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def default_form_raw() -> Dict[str, Any]:
    return {
        "form_id": "new_form",
        "title": "새 설문",
        "description": "",
        "questions": [
            {
                "id": "q1",
                "type": "short_text",
                "title": "질문 1",
                "description": "",
                "required": False,
            }
        ],
    }


def _sanitize_filename(name: str) -> str:
    base = "".join(ch for ch in (name or "") if ch.isalnum() or ch in ("-", "_"))
    return base or "form"


def form_raw_to_config(raw: Dict[str, Any], fallback_file_name: str = "") -> FormConfig:
    form_id = str(raw.get("form_id") or os.path.splitext(fallback_file_name)[0] or "form").strip()
    title = str(raw.get("title") or "(제목 없음)")
    description = str(raw.get("description") or "")

    questions_raw = raw.get("questions") or []
    questions: List[Question] = []
    for q in questions_raw:
        qid = str(q.get("id") or "").strip()
        qtype = str(q.get("type") or "").strip()
        qtitle = str(q.get("title") or "").strip()

        if not qid or not qtype or not qtitle:
            raise ValueError("각 질문은 id/type/title이 필요합니다.")
        if qtype not in SUPPORTED_TYPES:
            raise ValueError(f"지원하지 않는 type: {qtype}")

        options = tuple(str(x) for x in (q.get("options") or []))
        questions.append(
            Question(
                id=qid,
                type=qtype,
                title=qtitle,
                description=str(q.get("description") or ""),
                required=bool(q.get("required", False)),
                options=options,
                min=q.get("min"),
                max=q.get("max"),
                min_label=str(q.get("min_label") or ""),
                max_label=str(q.get("max_label") or ""),
            )
        )

    seen = set()
    for q in questions:
        if q.id in seen:
            raise ValueError(f"질문 id가 중복됩니다: {q.id}")
        seen.add(q.id)

    return FormConfig(form_id=form_id, title=title, description=description, questions=tuple(questions))


def save_form_raw(raw: Dict[str, Any], file_name: str) -> str:
    _ensure_dirs()
    safe_name = file_name
    if not safe_name.lower().endswith(".json"):
        safe_name += ".json"
    safe_name = _sanitize_filename(os.path.splitext(safe_name)[0]) + ".json"
    path = os.path.join(FORMS_DIR, safe_name)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=2)
    return safe_name


def load_form_config(file_name: str) -> FormConfig:
    raw = load_form_raw(file_name)
    return form_raw_to_config(raw, fallback_file_name=file_name)


def responses_path(form_id: str) -> str:
    safe = "".join(ch for ch in form_id if ch.isalnum() or ch in ("-", "_"))
    if not safe:
        safe = "form"
    return os.path.join(RESPONSES_DIR, f"{safe}.jsonl")


def append_response(form_id: str, payload: Dict[str, Any]) -> None:
    _ensure_dirs()
    path = responses_path(form_id)
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")


def load_responses(form_id: str) -> List[Dict[str, Any]]:
    path = responses_path(form_id)
    if not os.path.exists(path):
        return []
    rows: List[Dict[str, Any]] = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError:
                # 손상된 라인은 건너뜀
                continue
    return rows


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def render_question(q: Question) -> Any:
    label = q.title + (" *" if q.required else "")
    help_text = q.description or None

    key = f"q__{q.id}"

    if q.type == "short_text":
        return st.text_input(label, key=key, help=help_text)

    if q.type == "paragraph":
        return st.text_area(label, key=key, help=help_text)

    if q.type == "multiple_choice":
        if not q.options:
            return st.radio(label, options=["(옵션 없음)"], key=key, help=help_text, disabled=True)
        return st.radio(label, options=list(q.options), key=key, help=help_text)

    if q.type == "checkboxes":
        if not q.options:
            st.caption("옵션이 없습니다.")
            return []
        return st.multiselect(label, options=list(q.options), key=key, help=help_text)

    if q.type == "dropdown":
        if not q.options:
            return st.selectbox(label, options=["(옵션 없음)"], key=key, help=help_text, disabled=True)
        return st.selectbox(label, options=list(q.options), key=key, help=help_text)

    if q.type == "linear_scale":
        min_v = int(q.min) if q.min is not None else 1
        max_v = int(q.max) if q.max is not None else 5
        if min_v >= max_v:
            min_v, max_v = 1, 5

        cols = st.columns([1, 6, 1])
        with cols[0]:
            if q.min_label:
                st.caption(q.min_label)
        with cols[2]:
            if q.max_label:
                st.caption(q.max_label)
        with cols[1]:
            return st.slider(label, min_value=min_v, max_value=max_v, value=min_v, key=key, help=help_text)

    if q.type == "number":
        return st.number_input(label, key=key, help=help_text)

    if q.type == "date":
        return st.date_input(label, key=key, help=help_text)

    if q.type == "time":
        return st.time_input(label, key=key, help=help_text)

    raise ValueError(f"지원하지 않는 질문 타입: {q.type}")


def is_missing_required(q: Question, value: Any) -> bool:
    if not q.required:
        return False

    if value is None:
        return True

    if q.type in ("short_text", "paragraph"):
        return str(value).strip() == ""

    if q.type in ("multiple_choice", "dropdown"):
        return str(value).strip() == ""

    if q.type == "checkboxes":
        return len(value or []) == 0

    # date/time/number/scale은 Streamlit 기본값이 들어오므로 필수여부는 통과로 처리
    return False


def normalize_answer(q: Question, value: Any) -> Any:
    if q.type in ("date", "time"):
        # pandas/streamlit 객체를 JSON 직렬화 가능한 문자열로
        return str(value) if value is not None else None

    if q.type == "checkboxes":
        return list(value or [])

    return value


def _unique_question_id(existing: List[Dict[str, Any]]) -> str:
    used = {str(q.get("id") or "") for q in existing}
    i = 1
    while True:
        candidate = f"q{i}"
        if candidate not in used:
            return candidate
        i += 1


def _options_text_to_list(text: str) -> List[str]:
    lines = [ln.strip() for ln in (text or "").splitlines()]
    return [ln for ln in lines if ln]


def _options_list_to_text(options: List[str]) -> str:
    return "\n".join(options or [])



def render_form_editor(selected_file: str) -> None:
    st.subheader("폼 편집 (구글폼 스타일)")
    st.caption("구글폼처럼 우측 툴바에서 요소를 추가하고, 폼 파일을 선택/편집/저장할 수 있습니다.")

    state_key = "form_editor_raw"
    file_key = "form_editor_loaded_file"

    # 폼 파일 선택 드롭다운
    form_files = list_form_files()
    selected = st.selectbox("폼 파일 선택", options=form_files, index=form_files.index(selected_file) if selected_file in form_files else 0)

    col_a, col_b = st.columns([1, 1])
    with col_a:
        if st.button("선택한 폼 불러오기", use_container_width=True):
            try:
                st.session_state[state_key] = load_form_raw(selected)
                st.session_state[file_key] = selected
                st.success("불러왔습니다.")
            except Exception as e:
                st.error(f"불러오기 실패: {e}")
    with col_b:
        if st.button("새 폼 만들기", use_container_width=True):
            st.session_state[state_key] = default_form_raw()
            st.session_state[file_key] = ""
            st.success("새 폼을 시작합니다.")

    if state_key not in st.session_state:
        # 최초 진입 시 선택된 폼을 기본으로 로드
        try:
            st.session_state[state_key] = load_form_raw(selected)
            st.session_state[file_key] = selected
        except Exception as e:
            st.error(f"폼 설정을 불러오지 못했습니다: {e}")
            return

    raw: Dict[str, Any] = st.session_state[state_key]
    raw.setdefault("questions", [])

    st.markdown("---")
    st.markdown("#### 기본 정보")
    raw["form_id"] = st.text_input("form_id", value=str(raw.get("form_id") or ""))
    raw["title"] = st.text_input("제목", value=str(raw.get("title") or ""))
    raw["description"] = st.text_area("설명", value=str(raw.get("description") or ""))

    st.markdown("---")
    st.markdown("#### 질문")

    questions: List[Dict[str, Any]] = list(raw.get("questions") or [])

    # 구글폼 스타일 우측 툴바
    st.markdown("<div style='position:fixed; top:120px; right:32px; z-index:10;'>", unsafe_allow_html=True)
    col_toolbar = st.columns([1])
    with col_toolbar[0]:
        st.markdown("<div style='display:flex; flex-direction:column; gap:12px;'>", unsafe_allow_html=True)
        if st.button("➕ 질문 추가", key="toolbar_add"):
            questions.append({
                "id": _unique_question_id(questions),
                "type": "short_text",
                "title": f"질문 {len(questions) + 1}",
                "description": "",
                "required": False,
            })
            raw["questions"] = questions
            st.rerun()
        if st.button("📄 복제", key="toolbar_dup") and questions:
            new_q = dict(questions[-1])
            new_q["id"] = _unique_question_id(questions)
            questions.append(new_q)
            raw["questions"] = questions
            st.rerun()
        if st.button("🖼️ 이미지", key="toolbar_img"):
            questions.append({
                "id": _unique_question_id(questions),
                "type": "file_upload",
                "title": "이미지 업로드",
                "description": "이미지 파일을 첨부하세요.",
                "required": False,
                "file_types": ["jpg", "png"],
                "max_file_size_mb": 10,
            })
            raw["questions"] = questions
            st.rerun()
        if st.button("📺 동영상", key="toolbar_vid"):
            questions.append({
                "id": _unique_question_id(questions),
                "type": "short_text",
                "title": "동영상 링크",
                "description": "YouTube 등 동영상 URL을 입력하세요.",
                "required": False,
            })
            raw["questions"] = questions
            st.rerun()
        if st.button("🔖 섹션", key="toolbar_sec"):
            questions.append({
                "id": f"section{len([q for q in questions if q.get('type') == 'section']) + 1}",
                "type": "section",
                "title": f"{len([q for q in questions if q.get('type') == 'section']) + 1}페이지",
                "description": "",
            })
            raw["questions"] = questions
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # 질문 편집 UI (구글폼 스타일)
    for idx, q in enumerate(questions):
        q.setdefault("description", "")
        q.setdefault("required", False)
        with st.expander(f"{idx + 1}. {q.get('title') or '(제목 없음)'}", expanded=False):
            col1, col2 = st.columns([1, 1])
            with col1:
                q["id"] = st.text_input("질문 id", value=str(q.get("id") or ""), key=f"qe_id_{idx}")
            with col2:
                q["type"] = st.selectbox(
                    "타입",
                    options=sorted(SUPPORTED_TYPES),
                    index=sorted(SUPPORTED_TYPES).index(q.get("type")) if q.get("type") in SUPPORTED_TYPES else 0,
                    key=f"qe_type_{idx}",
                )
            q["title"] = st.text_input("질문 제목", value=str(q.get("title") or ""), key=f"qe_title_{idx}")
            q["description"] = st.text_area("설명(선택)", value=str(q.get("description") or ""), key=f"qe_desc_{idx}")
            if q.get("type") != "section":
                q["required"] = st.checkbox("필수", value=bool(q.get("required", False)), key=f"qe_req_{idx}")
            if q.get("type") in CHOICE_TYPES:
                opts_text = _options_list_to_text(list(q.get("options") or []))
                new_text = st.text_area(
                    "옵션(한 줄에 하나)",
                    value=opts_text,
                    key=f"qe_opts_{idx}",
                    help="객관식/체크박스/드롭다운에서 사용됩니다.",
                )
                q["options"] = _options_text_to_list(new_text)
            else:
                q.pop("options", None)
            if q.get("type") == "linear_scale":
                colm, colx = st.columns([1, 1])
                with colm:
                    q["min"] = int(st.number_input("최소", value=int(q.get("min") or 1), step=1, key=f"qe_min_{idx}"))
                    q["min_label"] = st.text_input("최소 라벨(선택)", value=str(q.get("min_label") or ""), key=f"qe_minl_{idx}")
                with colx:
                    q["max"] = int(st.number_input("최대", value=int(q.get("max") or 5), step=1, key=f"qe_max_{idx}"))
                    q["max_label"] = st.text_input("최대 라벨(선택)", value=str(q.get("max_label") or ""), key=f"qe_maxl_{idx}")
            else:
                q.pop("min", None)
                q.pop("max", None)
                q.pop("min_label", None)
                q.pop("max_label", None)
            if q.get("type") == "file_upload":
                ft_text = ", ".join(list(q.get("file_types") or []))
                ft_new = st.text_input("허용 파일 확장자(쉼표로 구분)", value=ft_text, key=f"qe_ft_{idx}")
                q["file_types"] = [x.strip() for x in ft_new.split(",") if x.strip()]
                q["max_file_size_mb"] = int(st.number_input("최대 파일 크기(MB)", value=int(q.get("max_file_size_mb") or 10), step=1, key=f"qe_fsz_{idx}"))
            if q.get("type") != "section":
                sec_opts = [q2["id"] for q2 in questions if q2.get("type") == "section"]
                q["section"] = st.selectbox("섹션(페이지) 연결", options=[""] + sec_opts, index=sec_opts.index(q.get("section")) if q.get("section") in sec_opts else 0, key=f"qe_sec_{idx}")
            st.markdown("---")
            col_up, col_down, col_del, col_dup = st.columns([1, 1, 1, 1])
            with col_up:
                if st.button("위로", key=f"qe_up_{idx}", use_container_width=True, disabled=(idx == 0)):
                    questions[idx - 1], questions[idx] = questions[idx], questions[idx - 1]
                    raw["questions"] = questions
                    st.rerun()
            with col_down:
                if st.button(
                    "아래로",
                    key=f"qe_down_{idx}",
                    use_container_width=True,
                    disabled=(idx >= len(questions) - 1),
                ):
                    questions[idx + 1], questions[idx] = questions[idx], questions[idx + 1]
                    raw["questions"] = questions
                    st.rerun()
            with col_del:
                if st.button("삭제", key=f"qe_del_{idx}", use_container_width=True):
                    questions.pop(idx)
                    raw["questions"] = questions
                    st.rerun()
            with col_dup:
                if st.button("복제", key=f"qe_dup_{idx}", use_container_width=True):
                    new_q = dict(q)
                    new_q["id"] = _unique_question_id(questions)
                    questions.insert(idx + 1, new_q)
                    raw["questions"] = questions
                    st.rerun()

    st.markdown("---")
    st.markdown("#### 저장")
    suggested_name = _sanitize_filename(str(raw.get("form_id") or "form")) + ".json"
    target_file = st.text_input("저장 파일명", value=suggested_name, help="forms 폴더에 저장됩니다.")

    col_save, col_hint = st.columns([1, 2])
    with col_save:
        if st.button("저장", use_container_width=True):
            try:
                # 저장 전 최소 검증(동일 로더로 체크)
                _ = form_raw_to_config(raw, fallback_file_name=target_file)
                saved = save_form_raw(raw, target_file)
                st.session_state[file_key] = saved
                st.success(f"저장 완료: forms/{saved}")
            except Exception as e:
                st.error(f"저장 실패: {e}")
    with col_hint:
        loaded = st.session_state.get(file_key) or selected_file
        st.caption(f"현재 로드된 폼: {loaded}")

    raw: Dict[str, Any] = st.session_state[state_key]
    raw.setdefault("questions", [])

    st.markdown("---")
    st.markdown("#### 기본 정보")
    raw["form_id"] = st.text_input("form_id", value=str(raw.get("form_id") or ""))
    raw["title"] = st.text_input("제목", value=str(raw.get("title") or ""))
    raw["description"] = st.text_area("설명", value=str(raw.get("description") or ""))

    st.markdown("---")
    st.markdown("#### 질문")

    questions: List[Dict[str, Any]] = list(raw.get("questions") or [])

    if st.button("+ 질문 추가"):
        questions.append(
            {
                "id": _unique_question_id(questions),
                "type": "short_text",
                "title": f"질문 {len(questions) + 1}",
                "description": "",
                "required": False,
            }
        )
        raw["questions"] = questions
        st.rerun()

    # 질문 편집 UI
    for idx, q in enumerate(questions):
        q.setdefault("description", "")
        q.setdefault("required", False)

        with st.expander(f"{idx + 1}. {q.get('title') or '(제목 없음)'}", expanded=False):
            col1, col2 = st.columns([1, 1])
            with col1:
                q["id"] = st.text_input("질문 id", value=str(q.get("id") or ""), key=f"qe_id_{idx}")
            with col2:
                q["type"] = st.selectbox(
                    "타입",
                    options=sorted(SUPPORTED_TYPES),
                    index=sorted(SUPPORTED_TYPES).index(q.get("type")) if q.get("type") in SUPPORTED_TYPES else 0,
                    key=f"qe_type_{idx}",
                )

            q["title"] = st.text_input("질문 제목", value=str(q.get("title") or ""), key=f"qe_title_{idx}")
            q["description"] = st.text_area("설명(선택)", value=str(q.get("description") or ""), key=f"qe_desc_{idx}")
            q["required"] = st.checkbox("필수", value=bool(q.get("required", False)), key=f"qe_req_{idx}")

            if q.get("type") in CHOICE_TYPES:
                opts_text = _options_list_to_text(list(q.get("options") or []))
                new_text = st.text_area(
                    "옵션(한 줄에 하나)",
                    value=opts_text,
                    key=f"qe_opts_{idx}",
                    help="객관식/체크박스/드롭다운에서 사용됩니다.",
                )
                q["options"] = _options_text_to_list(new_text)
            else:
                q.pop("options", None)

            if q.get("type") == "linear_scale":
                colm, colx = st.columns([1, 1])
                with colm:
                    q["min"] = int(st.number_input("최소", value=int(q.get("min") or 1), step=1, key=f"qe_min_{idx}"))
                    q["min_label"] = st.text_input("최소 라벨(선택)", value=str(q.get("min_label") or ""), key=f"qe_minl_{idx}")
                with colx:
                    q["max"] = int(st.number_input("최대", value=int(q.get("max") or 5), step=1, key=f"qe_max_{idx}"))
                    q["max_label"] = st.text_input("최대 라벨(선택)", value=str(q.get("max_label") or ""), key=f"qe_maxl_{idx}")
            else:
                q.pop("min", None)
                q.pop("max", None)
                q.pop("min_label", None)
                q.pop("max_label", None)

            st.markdown("---")
            col_up, col_down, col_del = st.columns([1, 1, 1])
            with col_up:
                if st.button("위로", key=f"qe_up_{idx}", use_container_width=True, disabled=(idx == 0)):
                    questions[idx - 1], questions[idx] = questions[idx], questions[idx - 1]
                    raw["questions"] = questions
                    st.rerun()
            with col_down:
                if st.button(
                    "아래로",
                    key=f"qe_down_{idx}",
                    use_container_width=True,
                    disabled=(idx >= len(questions) - 1),
                ):
                    questions[idx + 1], questions[idx] = questions[idx], questions[idx + 1]
                    raw["questions"] = questions
                    st.rerun()
            with col_del:
                if st.button("삭제", key=f"qe_del_{idx}", use_container_width=True):
                    questions.pop(idx)
                    raw["questions"] = questions
                    st.rerun()

    st.markdown("---")
    st.markdown("#### 저장")
    suggested_name = _sanitize_filename(str(raw.get("form_id") or "form")) + ".json"
    target_file = st.text_input("저장 파일명", value=suggested_name, help="forms 폴더에 저장됩니다.")

    col_save, col_hint = st.columns([1, 2])
    with col_save:
        if st.button("저장", use_container_width=True):
            try:
                # 저장 전 최소 검증(동일 로더로 체크)
                _ = form_raw_to_config(raw, fallback_file_name=target_file)
                saved = save_form_raw(raw, target_file)
                st.session_state[file_key] = saved
                st.success(f"저장 완료: forms/{saved}")
            except Exception as e:
                st.error(f"저장 실패: {e}")
    with col_hint:
        loaded = st.session_state.get(file_key) or selected_file
        st.caption(f"현재 로드된 폼: {loaded}")


def responses_to_dataframe(form: FormConfig, rows: List[Dict[str, Any]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()

    records: List[Dict[str, Any]] = []
    question_titles = {q.id: q.title for q in form.questions}

    for r in rows:
        base = {
            "submitted_at": r.get("submitted_at"),
        }
        answers = r.get("answers") or {}
        for qid, val in answers.items():
            col = question_titles.get(qid, qid)
            base[col] = val
        records.append(base)

    df = pd.DataFrame.from_records(records)
    # 제출시간 최신순
    if "submitted_at" in df.columns:
        df = df.sort_values("submitted_at", ascending=False)
    return df



def main() -> None:
    st.set_page_config(page_title="개인 설문 대시보드", layout="centered")

    st.sidebar.title("설문 선택")
    form_files = list_form_files()
    if not form_files:
        st.sidebar.warning("forms 폴더에 JSON 설문이 없습니다.")
        st.stop()

    selected = st.sidebar.selectbox("폼 파일", options=form_files, index=0)

    mode = st.sidebar.radio("모드", options=["설문 작성", "응답 보기", "분석", "폼 편집"], index=0)

    if mode == "폼 편집":
        st.title("개인 설문 대시보드")
        render_form_editor(selected)
        return

    try:
        form = load_form_config(selected)
    except Exception as e:
        st.error(f"폼 설정을 불러오지 못했습니다: {e}")
        st.stop()

    st.title(form.title)
    if form.description:
        st.caption(form.description)


    if mode == "설문 작성":
        st.markdown(f"<div style='background:#fff; border-radius:12px; box-shadow:0 2px 8px #eee; padding:32px 24px 24px 24px; margin-bottom:24px;'>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='margin-bottom:8px;'>{form.title}</h2>", unsafe_allow_html=True)
        if form.description:
            st.markdown(f"<div style='color:#555; margin-bottom:16px;'>{form.description}</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        with st.form("survey_form", clear_on_submit=True):
            rendered: Dict[str, Tuple[Question, Any]] = {}
            for q in form.questions:
                if q.type == "section":
                    st.markdown(f"<div style='margin:32px 0 12px 0; font-size:1.1em; color:#333; font-weight:600;'>{q.title}</div>", unsafe_allow_html=True)
                    if q.description:
                        st.markdown(f"<div style='color:#888; margin-bottom:8px;'>{q.description}</div>", unsafe_allow_html=True)
                    continue
                st.markdown("<div style='background:#fafafa; border-radius:8px; box-shadow:0 1px 4px #eee; padding:18px 16px; margin-bottom:18px;'>", unsafe_allow_html=True)
                label = q.title + (" <span style='color:#d00;'>*</span>" if q.required else "")
                st.markdown(f"<div style='font-weight:500; font-size:1.05em;'>{label}</div>", unsafe_allow_html=True)
                if q.description:
                    st.markdown(f"<div style='color:#888; margin-bottom:6px;'>{q.description}</div>", unsafe_allow_html=True)
                value = render_question(q)
                rendered[q.id] = (q, value)
                st.markdown("</div>", unsafe_allow_html=True)

            submitted = st.form_submit_button("<span style='font-size:1.1em;'>제출</span>", use_container_width=True)

        if submitted:
            missing: List[str] = []
            answers: Dict[str, Any] = {}
            for qid, (q, value) in rendered.items():
                if is_missing_required(q, value):
                    missing.append(q.title)
                answers[qid] = normalize_answer(q, value)

            if missing:
                st.error("필수 질문을 입력해 주세요: " + ", ".join(missing))
            else:
                payload = {
                    "form_id": form.form_id,
                    "submitted_at": utc_now_iso(),
                    "answers": answers,
                }
                append_response(form.form_id, payload)
                st.success("제출되었습니다. 감사합니다!")

    elif mode == "응답 보기":
        rows = load_responses(form.form_id)
        st.subheader("응답")
        st.caption(f"총 {len(rows)}개")

        df = responses_to_dataframe(form, rows)
        if df.empty:
            st.info("아직 응답이 없습니다.")
            return

        st.dataframe(df, use_container_width=True, hide_index=True)

        csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="CSV 다운로드",
            data=csv_bytes,
            file_name=f"{form.form_id}_responses.csv",
            mime="text/csv",
        )

    elif mode == "분석":
        rows = load_responses(form.form_id)
        st.subheader("자동 분석 결과")
        st.caption(f"총 {len(rows)}개 응답 분석")

        df = responses_to_dataframe(form, rows)
        if df.empty:
            st.info("아직 응답이 없습니다.")
            return

        for q in form.questions:
            if q.type == "section":
                continue
            st.markdown(f"### {q.title}")
            col = q.title
            vals = df[col].dropna().tolist() if col in df.columns else []
            if not vals:
                st.caption("응답 없음")
                continue

            # 객관식/체크박스/드롭다운
            if q.type in ("multiple_choice", "dropdown"):
                cnt = Counter(vals)
                st.bar_chart(cnt)
                st.caption("응답 분포")
            elif q.type == "checkboxes":
                flat = []
                for v in vals:
                    if isinstance(v, list):
                        flat.extend(v)
                    elif isinstance(v, str):
                        flat.extend([x.strip() for x in v.split(",") if x.strip()])
                cnt = Counter(flat)
                st.bar_chart(cnt)
                st.caption("복수 선택 분포")
            # 숫자/스케일
            elif q.type in ("number", "linear_scale"):
                nums = [float(v) for v in vals if isinstance(v, (int, float)) or re.match(r'^-?\d+(\.\d+)?$', str(v))]
                if nums:
                    st.write(f"평균: {sum(nums)/len(nums):.2f}, 표준편차: {pd.Series(nums).std():.2f}")
                    st.plotly_chart(px.histogram(nums, nbins=10, title="분포 히스토그램"))
            # 텍스트/의견
            elif q.type in ("short_text", "paragraph"):
                text = " ".join([str(v) for v in vals])
                words = [w for w in re.findall(r'\w+', text) if len(w) > 1]
                cnt = Counter(words)
                top = cnt.most_common(20)
                st.write("워드클라우드(상위 20):")
                st.write({w: c for w, c in top})
                st.write("예시 응답:")
                for v in vals[:5]:
                    st.caption(str(v))
            # 파일 업로드
            elif q.type == "file_upload":
                st.write(f"파일 응답 개수: {len(vals)}")
                ext_cnt = Counter()
                for v in vals:
                    if hasattr(v, 'name'):
                        ext = os.path.splitext(v.name)[1].lower()
                        ext_cnt[ext] += 1
                st.write("확장자 분포:", dict(ext_cnt))
            # 날짜/시간
            elif q.type in ("date", "time"):
                st.write(f"응답 개수: {len(vals)}")
                st.write("예시:")
                for v in vals[:5]:
                    st.caption(str(v))
            else:
                st.write(f"응답 개수: {len(vals)}")
            st.markdown("---")


if __name__ == "__main__":
    main()
