$host.UI.RawUI.WindowTitle = "Yardmap Server"

# Check if port 8765 is already in use (prior session still running)
$listening = Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue
if ($listening) {
    Write-Host "Yardmap is already running. Opening browser..."
    Start-Process "http://localhost:8765"
    exit
}

# Open the browser after a short delay, hidden so no extra window appears
Start-Process "cmd.exe" -ArgumentList "/c timeout /t 2 /nobreak >nul & start http://localhost:8765" -WindowStyle Hidden

# Run the server directly in this window -- close this window to stop
Set-Location $PSScriptRoot
Write-Host "Yardmap Server running at http://localhost:8765"
Write-Host "Close this window to stop the server."
Write-Host ""
& python -m http.server 8765
