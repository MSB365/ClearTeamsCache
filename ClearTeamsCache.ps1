IF ((Test-Path 'HKCU:Software\Microsoft\MSB365_Teams_clear_cache_Tool') -eq $false)
{
################################################
$directory0 = "C:\MDM\Logging"
Write-Host "Checking if $directory0 is available..." -ForegroundColor Magenta
Start-Sleep -s 1
If ((Test-Path -Path $directory0) -eq $false)
{
        Write-Host "The Directory $directory0 don't exists!" -ForegroundColor Cyan
        Start-Sleep -s 2
        Write-Host "Creating directory $directory0 ..." -ForegroundColor Cyan
        Start-Sleep -s 2
        New-Item -Path $directory0 -ItemType directory
        Start-Sleep -s 2
        Write-Host "New Directory $directory0 is is created" -ForegroundColor Green
        Start-Sleep -s 2
}else{
Write-Host "The Path $directory0 already exists" -ForegroundColor green
Start-Sleep -s 3
}
################################################
Start-Transcript -Path "C:\MDM\Logging\ClearTeamsCach.txt" -NoClobber
################################################
$directory1 = "C:\MDM\ClearTeamsCache"
Write-Host "Checking if $directory1 is available..." -ForegroundColor Magenta
Start-Sleep -s 1
If ((Test-Path -Path $directory1) -eq $false)
{
        Write-Host "The Directory $directory1 don't exists!" -ForegroundColor Cyan
        Start-Sleep -s 2
        Write-Host "Creating directory $directory1 ..." -ForegroundColor Cyan
        Start-Sleep -s 2
        New-Item -Path $directory1 -ItemType directory
        Start-Sleep -s 2
        Write-Host "New Directory $directory1 is is created" -ForegroundColor Green
        Start-Sleep -s 2
}else{
Write-Host "The Path $directory1 already exists" -ForegroundColor green
Start-Sleep -s 3
}
################################################
$WebClient = New-Object System.Net.WebClient
$WebClient.DownloadFile("https://raw.githubusercontent.com/MSB365/ClearTeamsCache/main/ClearTeamsCache.bat","C:\MDM\ClearTeamsCache\ClearTeamsCache.bat")
Start-Sleep -s 1
################################################
Start-Process C:\MDM\ClearTeamsCache\ClearTeamsCache.bat
Start-Sleep -s 1
################################################
New-Item -Path HKCU:Software\Microsoft\MSB365_Teams_clear_cache_Tool
Get-Item -Path "HKCU:Software\Microsoft\MSB365_Teams_clear_cache_Tool" | New-ItemProperty -Name CacheCleared -Value yes
################################################
Remove-Item –path C:\MDM\ClearTeamsCache\* -include *.bat
################################################
Stop-Transcript
################################################
Start-Sleep -s 1
}else{
Start-Transcript -Path "C:\MDM\Logging\ClearTeamsCach.txt" -NoClobber
Write-Host "No action needed. The team cache has already been cleared!" -ForegroundColor Green
Stop-Transcript
}
