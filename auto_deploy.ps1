$watchDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $watchDir
$watchFiles = @("dashboard.py", "*.json", "*.py", "requirements.txt", "runtime.txt")
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  자동 배포 감시 시작" -ForegroundColor Cyan
Write-Host "  Ctrl+C 로 종료" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $watchDir
$watcher.IncludeSubdirectories = $true
$watcher.EnableRaisingEvents = $false
$watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::FileName
$lastPush = [DateTime]::MinValue
$cooldownSeconds = 30

while ($true) {
    $changeType = [System.IO.WatcherChangeTypes]::Changed -bor [System.IO.WatcherChangeTypes]::Created
    $result = $watcher.WaitForChanged($changeType, 10000)
    if ($result.TimedOut) { continue }
    $changedFile = $result.Name
    $now = Get-Date
    $isRelevant = $false
    foreach ($p in $watchFiles) {
        if ($changedFile -like $p) { $isRelevant = $true; break }
    }
    if ($changedFile -like "image_archive\*") { $isRelevant = $true }
    if (-not $isRelevant) { continue }
    $elapsed = ($now - $lastPush).TotalSeconds
    if ($elapsed -lt $cooldownSeconds) { continue }
    $ts = Get-Date -Format "HH:mm:ss"
    Write-Host ""
    Write-Host "[$ts] 변경 감지: $changedFile" -ForegroundColor Green
    Start-Sleep -Seconds 3
    try {
        Write-Host "  [1/3] 스테이징..." -ForegroundColor White
        git add dashboard.py requirements.txt runtime.txt *.json *.py 2>$null
        git add image_archive/*.jpg 2>$null
        git add product_images/*.jpg 2>$null
        git add product_images_hd/*.jpg 2>$null
        $status = git diff --cached --stat 2>&1
        if ([string]::IsNullOrWhiteSpace($status)) {
            Write-Host "  [스킵] 변경사항 없음" -ForegroundColor Yellow
            continue
        }
        $stamp = Get-Date -Format "yyyy-MM-dd HH:mm"
        $msg = "자동 업데이트 $stamp"
        Write-Host "  [2/3] 커밋: $msg" -ForegroundColor White
        git commit -m $msg 2>$null
        Write-Host "  [3/3] GitHub 푸시..." -ForegroundColor White
        git push origin master 2>&1
        $lastPush = Get-Date
        Write-Host "  [완료] Streamlit Cloud 반영!" -ForegroundColor Green
    }
    catch {
        Write-Host "  [오류] 배포 실패" -ForegroundColor Red
    }
}
