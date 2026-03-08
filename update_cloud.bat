@echo off
chcp 65001 >nul
echo ========================================
echo   Streamlit Cloud 자동 업데이트
echo ========================================
echo.

cd /d "%~dp0"

echo [1/3] 변경된 파일 확인 중...
git add dashboard.py requirements.txt brands_config.json *.json *.xlsx .streamlit\config.toml

echo [2/3] 커밋 생성 중...
for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set TODAY=%%a-%%b-%%c
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set NOW=%%a:%%b
git commit -m "데이터 업데이트 %TODAY% %NOW%"

echo [3/3] Streamlit Cloud에 배포 중...
git push origin master

echo.
echo ========================================
echo   업데이트 완료! 잠시 후 자동 반영됩니다.
echo   https://brand-ranking-dashboard-fwe6wtyqmcjsddasjjaunt.streamlit.app/
echo ========================================
pause
