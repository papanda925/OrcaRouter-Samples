param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot "import-ready")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$sourceFiles = @(
    "OrcaRouterSample.bas",
    "OrcaRouterAdvanced.bas"
)

$utf8 = [System.Text.UTF8Encoding]::new($false, $true)
$shiftJis = [System.Text.Encoding]::GetEncoding(
    932,
    [System.Text.EncoderExceptionFallback]::new(),
    [System.Text.DecoderExceptionFallback]::new()
)

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

foreach ($fileName in $sourceFiles) {
    $sourcePath = Join-Path $PSScriptRoot $fileName
    $destinationPath = Join-Path $OutputDirectory $fileName

    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Source file not found: $sourcePath"
    }

    $sourceBytes = [System.IO.File]::ReadAllBytes($sourcePath)
    $text = $utf8.GetString($sourceBytes)
    $destinationBytes = $shiftJis.GetBytes($text)

    [System.IO.File]::WriteAllBytes($destinationPath, $destinationBytes)

    Write-Host ("Created VBE import file: {0}" -f $destinationPath)
}

Write-Host ""
Write-Host "Import the two .bas files from the import-ready folder in the VBA editor."
Write-Host "Do not edit or commit the generated files back to GitHub."
