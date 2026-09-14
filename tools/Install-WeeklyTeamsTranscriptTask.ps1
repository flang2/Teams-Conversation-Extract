[CmdletBinding()]
param(
    [string]$ConfigPath,
    [string]$TaskName = "Weekly Teams Transcript Export",
    [ValidateSet("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
    [string]$DayOfWeek = "Friday",
    [datetime]$RunAt = "10:00",
    [switch]$Replace
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"


$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $repoRoot "config.psd1"
}
$runnerScript = Join-Path $PSScriptRoot "Invoke-WeeklyTeamsTranscript.ps1"


if (-not (Test-Path -LiteralPath $runnerScript)) {
    throw "Invoke-WeeklyTeamsTranscript.ps1 is missing from the tools folder."
}
if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "config.psd1 was not found. Run Setup-WeeklyTeamsTranscript.ps1 first."
}


$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -ne $existingTask -and -not $Replace) {
    throw "A scheduled task named '$TaskName' already exists. Rerun with -Replace only if you intend to replace that task."
}


$engine = (Get-Process -Id $PID).Path
$actionArguments = '-NoLogo -NoProfile -ExecutionPolicy Bypass -File "' + $runnerScript + '" -ConfigPath "' + $ConfigPath + '"'
$action = New-ScheduledTaskAction -Execute $engine -Argument $actionArguments
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek -At $RunAt
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited


$registrationParameters = @{
    TaskName = $TaskName
    Action = $action
    Trigger = $trigger
    Settings = $settings
    Principal = $principal
    Description = "Exports the previous seven days of the signed-in user's Microsoft Teams chats to Markdown."
}
if ($Replace) {
    $registrationParameters.Force = $true
}


Register-ScheduledTask @registrationParameters | Out-Null
Write-Host "Scheduled task created: $TaskName"
Write-Host "Schedule: $DayOfWeek at $($RunAt.ToString('HH:mm')) local Windows time"
Write-Host "Windows account: $currentUser"
Write-Host "The task runs while this user is logged on so delegated Microsoft Graph authentication is available."