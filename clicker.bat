@powershell -windowstyle hidden -NoProfile -ExecutionPolicy Unrestricted "&([ScriptBlock]::Create((cat -encoding utf8 \"%~f0\" | ? {$_.ReadCount -gt 2}) -join \"`n\"))" %*
@exit /b

Add-Type -AssemblyName System.Windows.Forms
$signature=@' 
      [DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
      public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@ 
$SendMouseClick = Add-Type -memberDefinition $signature -name "Win32MouseEventNew" -namespace Win32Functions -passThru;

# capslock がかかっていたら解除
if([System.Console]::CapsLock){
    $wsh = New-Object -ComObject WScript.Shell
    $wsh.SendKeys('{CAPSLOCK}')
}

[int]$closeFlag = 0
while(1) {
    # 50/1000 秒待機
    Start-Sleep -Milliseconds  50
    if([System.Console]::CapsLock){
        # capslock がかかったら左クリック
        $SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0);
        $SendMouseClick::mouse_event(0x00000004, 0, 0, 0, 0);
        $closeFlag = 1;
    }elseif ($closeFlag -eq 1) {
        # Write-Output("end")
        # capslock 外したら終了
        exit
    }
}