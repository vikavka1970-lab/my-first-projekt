$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:8080/")
$listener.Start()
Write-Host "Server started: http://192.168.88.178:8080/vizitka.html"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path

while ($listener.IsListening) {
    $context = $listener.GetContext()
    $path = $context.Request.Url.LocalPath.TrimStart('/')
    if ($path -eq '') { $path = 'vizitka.html' }
    $filePath = Join-Path $root $path
    $response = $context.Response
    if (Test-Path $filePath) {
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        $response.ContentType = "text/html; charset=utf-8"
        $response.ContentLength64 = $bytes.Length
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
        $response.StatusCode = 404
    }
    $response.OutputStream.Close()
}
