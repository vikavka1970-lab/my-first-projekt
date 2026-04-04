Add-Type -AssemblyName System.Windows.Forms

Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
public class WinAPI2 {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc f, IntPtr p);
    [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, IntPtr extra);
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public static IntPtr FindWindowByTitle(string part) {
        IntPtr found = IntPtr.Zero;
        EnumWindows((hWnd, lp) => {
            if (!IsWindowVisible(hWnd)) return true;
            var sb = new StringBuilder(256);
            GetWindowText(hWnd, sb, 256);
            if (sb.ToString().Contains(part)) { found = hWnd; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static void PasteToWindow(IntPtr hWnd) {
        ShowWindow(hWnd, 9);
        SetForegroundWindow(hWnd);
        Thread.Sleep(500);
        keybd_event(0x11, 0, 0, IntPtr.Zero); // Ctrl down
        keybd_event(0x56, 0, 0, IntPtr.Zero); // V down
        keybd_event(0x56, 0, 2, IntPtr.Zero); // V up
        keybd_event(0x11, 0, 2, IntPtr.Zero); // Ctrl up
    }
}
"@

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:8765/")
$listener.Start()

$htmlPath = "$PSScriptRoot\voice-input.html"
Start-Process "C:\Users\Acer\AppData\Local\Chromium\Application\chrome.exe" "http://localhost:8765/"

while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response

    if ($request.Url.LocalPath -eq "/send" -and $request.HttpMethod -eq "POST") {
        $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
        $text = $reader.ReadToEnd()
        $reader.Close()

        [System.Windows.Forms.Clipboard]::SetText($text)

        $buf = [System.Text.Encoding]::UTF8.GetBytes("ok")
        $response.StatusCode = 200
        $response.ContentLength64 = $buf.Length
        $response.OutputStream.Write($buf, 0, $buf.Length)
        $response.OutputStream.Close()

        $vsCode = [WinAPI2]::FindWindowByTitle("Visual Studio Code")
        if ($vsCode -ne [IntPtr]::Zero) {
            [WinAPI2]::PasteToWindow($vsCode)
        }
    } else {
        $html = Get-Content -Raw -Path $htmlPath -Encoding UTF8
        $buf = [System.Text.Encoding]::UTF8.GetBytes($html)
        $response.ContentType = "text/html; charset=utf-8"
        $response.ContentLength64 = $buf.Length
        $response.OutputStream.Write($buf, 0, $buf.Length)
        $response.OutputStream.Close()
    }
}
