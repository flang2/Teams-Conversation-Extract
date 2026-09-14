[CmdletBinding()]
param(
    [string]$ConfigPath
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"


$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $repoRoot "config.psd1"
}
$exampleConfig = Join-Path $repoRoot "config.example.psd1"
$runnerScript = Join-Path $PSScriptRoot "Invoke-WeeklyTeamsTranscript.ps1"


if (-not (Test-Path -LiteralPath $runnerScript)) {
    throw "Invoke-WeeklyTeamsTranscript.ps1 is missing from the tools folder."
}
if (-not (Test-Path -LiteralPath $ConfigPath)) {
    if (-not (Test-Path -LiteralPath $exampleConfig)) {
        throw "Neither config.psd1 nor config.example.psd1 was found in the repository root."
    }
    Copy-Item -LiteralPath $exampleConfig -Destination $ConfigPath
    Write-Host "Created config.psd1 from the example configuration."
}


$config = Import-PowerShellDataFile -LiteralPath $ConfigPath
$expectedAccount = [string]$config.ExpectedAccount


if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host "Installing Microsoft.Graph.Authentication for the current Windows user..."
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -Scope CurrentUser -Force | Out-Null
    }
    $repository = Get-PSRepository -Name PSGallery
    $restoreRepositoryPolicy = $repository.InstallationPolicy -ne "Trusted"
    if ($restoreRepositoryPolicy) {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }
    try {
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
    }
    finally {
        if ($restoreRepositoryPolicy) {
            Set-PSRepository -Name PSGallery -InstallationPolicy Untrusted
        }
    }
}
else {
    Write-Host "Microsoft Graph authentication module is already installed."
}


Import-Module Microsoft.Graph.Authentication -Force
Write-Host "A Microsoft sign-in window will open."
Write-Host "Sign in with the Microsoft 365 account whose Teams chats should be exported."
Connect-MgGraph -Scopes "Chat.Read", "User.Read" -ContextScope CurrentUser -NoWelcome


$context = Get-MgContext
if ($null -eq $context -or [string]::IsNullOrWhiteSpace([string]$context.Account)) {
    throw "Microsoft Graph authentication did not return a signed-in account."
}
if (-not [string]::IsNullOrWhiteSpace($expectedAccount) -and
    -not ([string]$context.Account).Equals($expectedAccount, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Authentication completed as '$($context.Account)' instead of '$expectedAccount'. Run Disconnect-MgGraph and repeat setup."
}
if (@($context.Scopes) -notcontains "Chat.Read" -or @($context.Scopes) -notcontains "User.Read") {
    throw "Authentication succeeded, but Chat.Read and User.Read were not both granted."
}


Write-Host "Authentication succeeded as $($context.Account). Running a one-day test export..."
$engine = (Get-Process -Id $PID).Path
$argumentString = '-NoLogo -NoProfile -ExecutionPolicy Bypass -File "' + $runnerScript + '" -ConfigPath "' + $ConfigPath + '" -DaysBack 1'
$process = Start-Process -FilePath $engine -ArgumentList $argumentString -Wait -PassThru -NoNewWindow
if ($process.ExitCode -ne 0) {
    throw "The one-day test export failed with exit code $($process.ExitCode). Review Teams-Chat-Transcript-LastRun.log in the configured output folder."
}


Write-Host "Setup and the one-day test export completed successfully."
Write-Host "Run Install-WeeklyTeamsTranscriptTask.ps1 to create the weekly schedule."