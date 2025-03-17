# Check if the script is running with administrative privileges
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal($currentUser)
$isAdmin = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "This script must be run as administrator." -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# Define the path to the teams-dj.ps1 script
$scriptPath = Join-Path -Path (Get-Location) -ChildPath "teams-dj.ps1"

# Define the scheduled task action
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-File `"$scriptPath`""

# Define the scheduled task trigger (runs at logon)
$trigger = New-ScheduledTaskTrigger -AtLogOn

# Define the scheduled task principal (current user)
$principal = New-ScheduledTaskPrincipal -UserId (whoami) -LogonType Interactive

# Register the scheduled task
Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName "Teams DJ" -Description "Runs the teams-dj.ps1 script at logon"

Write-Host "Scheduled task created successfully." -ForegroundColor Green
