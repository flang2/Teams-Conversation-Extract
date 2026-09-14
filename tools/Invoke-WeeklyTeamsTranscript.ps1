[CmdletBinding()]
param(
    [string]$ConfigPath,
    [ValidateRange(1, 30)]
    [int]$DaysBack
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"


$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $repoRoot "config.psd1"
}
$exportScript = Join-Path $PSScriptRoot "Export-WeeklyTeamsTranscript.ps1"
if (-not (Test-Path -LiteralPath $exportScript)) {
    throw "Export-WeeklyTeamsTranscript.ps1 is missing from the tools folder."
}


$exportArguments = @{ ConfigPath = $ConfigPath }
if ($PSBoundParameters.ContainsKey("DaysBack")) {
    $exportArguments.DaysBack = $DaysBack
}


& $exportScript @exportArguments