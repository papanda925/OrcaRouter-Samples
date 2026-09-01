param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot "import-ready")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$sourceFiles = @(
    "OrcaRouterSample.bas",
    "OrcaRouterAdvanced.bas"
)

$utf8 = [System.Text.Encoding]::UTF8
$shiftJis = [System.Text.Encoding]::GetEncoding(932)

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

foreach ($fileName in $sourceFiles) {
    $sourcePath = Join-Path $PSScriptRoot $fileName
    $destinationPath = Join-Path $OutputDirectory $fileName

    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Source file not found: $sourcePath"
    }

    $text = [System.IO.File]::ReadAllText($sourcePath, $utf8)
    [System.IO.File]::WriteAllText($destinationPath, $text, $shiftJis)

    Write-Host ("Created VBE import file: {0}" -f $destinationPath)
}

Write-Host ""
Write-Host "Import the two .bas files from the import-ready folder in the VBA editor."
Write-Host "Do not edit or commit the generated files back to GitHub."
