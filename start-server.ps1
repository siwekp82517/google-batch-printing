$Port = 8080
$Url = "http://localhost:$Port"

Write-Host "Starting Google Batch PDF Printer..." -ForegroundColor Cyan
Write-Host "Opening $Url" -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop the server.`n"

$python = $null
foreach ($cmd in @('python', 'python3', 'py')) {
    try { $python = Get-Command $cmd -ErrorAction Stop; break } catch {}
}

if ($python) {
    Start-Process $Url
    & $python.Source -m http.server $Port
} else {
    Write-Host "Python not found. Using .NET HttpListener fallback..." -ForegroundColor Yellow

    Add-Type -TypeDefinition @"
using System;
using System.Net;
using System.IO;
using System.Threading;

public class SimpleServer
{
    public static void Main(string[] args)
    {
        int port = int.Parse(args[0]);
        string prefix = "http://localhost:" + port + "/";
        HttpListener listener = new HttpListener();
        listener.Prefixes.Add(prefix);
        listener.Start();
        Console.WriteLine("Serving on " + prefix);
        Console.WriteLine("Press Ctrl+C to stop.");

        while (true)
        {
            HttpListenerContext context = listener.GetContext();
            string url = context.Request.Url.AbsolutePath;
            if (url == "/") url = "/index.html";
            string path = "." + url.Replace('/', '\\');
            if (File.Exists(path))
            {
                string ext = Path.GetExtension(path).ToLower();
                string contentType = ext switch
                {
                    ".html" => "text/html",
                    ".css" => "text/css",
                    ".js" => "application/javascript",
                    ".json" => "application/json",
                    ".png" => "image/png",
                    ".svg" => "image/svg+xml",
                    ".ico" => "image/x-icon",
                    _ => "application/octet-stream"
                };
                byte[] bytes = File.ReadAllBytes(path);
                context.Response.ContentType = contentType;
                context.Response.ContentLength64 = bytes.Length;
                context.Response.OutputStream.Write(bytes, 0, bytes.Length);
                context.Response.OutputStream.Close();
            }
            else
            {
                context.Response.StatusCode = 404;
                context.Response.Close();
            }
        }
    }
}
"@ -Language CSharp

    Start-Process $Url
    [SimpleServer]::Main($Port)
}
