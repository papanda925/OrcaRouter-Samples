param(
    [int]$Port = 8000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$listener = [System.Net.Sockets.TcpListener]::new(
    [System.Net.IPAddress]::Loopback,
    $Port
)

function Get-ContentType {
    param([string]$Path)

    switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        ".html" { return "text/html; charset=utf-8" }
        ".css"  { return "text/css; charset=utf-8" }
        ".js"   { return "application/javascript; charset=utf-8" }
        ".json" { return "application/json; charset=utf-8" }
        ".svg"  { return "image/svg+xml" }
        ".png"  { return "image/png" }
        ".jpg"  { return "image/jpeg" }
        ".jpeg" { return "image/jpeg" }
        ".ico"  { return "image/x-icon" }
        default { return "application/octet-stream" }
    }
}

function Write-HttpResponse {
    param(
        [System.Net.Sockets.NetworkStream]$Stream,
        [int]$StatusCode,
        [string]$StatusText,
        [byte[]]$Body,
        [string]$ContentType
    )

    $headerText =
        "HTTP/1.1 $StatusCode $StatusText`r`n" +
        "Content-Type: $ContentType`r`n" +
        "Content-Length: $($Body.Length)`r`n" +
        "Cache-Control: no-store`r`n" +
        "Connection: close`r`n`r`n"

    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($headerText)

    $Stream.Write($headerBytes, 0, $headerBytes.Length)

    if ($Body.Length -gt 0) {
        $Stream.Write($Body, 0, $Body.Length)
    }

    $Stream.Flush()
}

$listener.Start()

Write-Host ""
Write-Host "OrcaRouter Web sample server started."
Write-Host "Open: http://localhost:$Port/"
Write-Host "Root: $root"
Write-Host "Stop: Ctrl+C"
Write-Host ""

try {
    while ($true) {
        $client = $listener.AcceptTcpClient()

        try {
            $stream = $client.GetStream()
            $reader = [System.IO.StreamReader]::new(
                $stream,
                [System.Text.Encoding]::ASCII,
                $false,
                1024,
                $true
            )

            $requestLine = $reader.ReadLine()

            if ([string]::IsNullOrWhiteSpace($requestLine)) {
                $client.Close()
                continue
            }

            while ($true) {
                $line = $reader.ReadLine()
                if ([string]::IsNullOrEmpty($line)) {
                    break
                }
            }

            $parts = $requestLine.Split(" ")

            if ($parts.Count -lt 2 -or $parts[0] -ne "GET") {
                $body = [System.Text.Encoding]::UTF8.GetBytes("Method Not Allowed")
                Write-HttpResponse -Stream $stream -StatusCode 405 -StatusText "Method Not Allowed" -Body $body -ContentType "text/plain; charset=utf-8"
                continue
            }

            $requestPath = $parts[1].Split("?")[0]
            $requestPath = [System.Uri]::UnescapeDataString($requestPath)

            if ($requestPath -eq "/") {
                $requestPath = "/index.html"
            }

            $relativePath = $requestPath.TrimStart("/").Replace("/", [System.IO.Path]::DirectorySeparatorChar)
            $candidate = [System.IO.Path]::GetFullPath((Join-Path $root $relativePath))
            $rootFull = [System.IO.Path]::GetFullPath($root + [System.IO.Path]::DirectorySeparatorChar)

            if (-not $candidate.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
                $body = [System.Text.Encoding]::UTF8.GetBytes("Forbidden")
                Write-HttpResponse -Stream $stream -StatusCode 403 -StatusText "Forbidden" -Body $body -ContentType "text/plain; charset=utf-8"
                continue
            }

            if (-not [System.IO.File]::Exists($candidate)) {
                $body = [System.Text.Encoding]::UTF8.GetBytes("Not Found")
                Write-HttpResponse -Stream $stream -StatusCode 404 -StatusText "Not Found" -Body $body -ContentType "text/plain; charset=utf-8"
                continue
            }

            $body = [System.IO.File]::ReadAllBytes($candidate)
            $contentType = Get-ContentType -Path $candidate

            Write-Host ("GET {0} -> 200" -f $requestPath)
            Write-HttpResponse -Stream $stream -StatusCode 200 -StatusText "OK" -Body $body -ContentType $contentType
        }
        catch {
            Write-Warning $_.Exception.Message
        }
        finally {
            if ($null -ne $client) {
                $client.Close()
            }
        }
    }
}
finally {
    $listener.Stop()
}
