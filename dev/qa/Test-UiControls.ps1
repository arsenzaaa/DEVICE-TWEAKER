[CmdletBinding()]
param([string]$AssemblyPath, [string]$OutputDirectory, [switch]$Worker)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Private-Desktop.ps1')
$AssemblyPath = [IO.Path]::GetFullPath($AssemblyPath)
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if (-not $Worker) {
    $code = Invoke-QaPrivateDesktop -ScriptPath $PSCommandPath -Parameters @{
        AssemblyPath = $AssemblyPath; OutputDirectory = $OutputDirectory; Worker = $true
    } -ReportDirectory (Join-Path $OutputDirectory 'isolation') -TimeoutSeconds 120
    if ($null -ne $code) { exit $code }
}
Assert-QaPrivateDesktop
# The entry point is an STA thread; all windows belong to the private QA desktop.
$ErrorActionPreference = 'Continue'
& dotnet run --project (Join-Path $PSScriptRoot 'Inspect\Inspect.csproj') -- `
    --controls $AssemblyPath $OutputDirectory 2>&1 | ForEach-Object { $_.ToString() } |
    Tee-Object -FilePath (Join-Path $OutputDirectory 'controls.log') | Out-Host
$code = $LASTEXITCODE
exit $code
