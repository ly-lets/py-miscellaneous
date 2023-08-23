### sth to share

```powershell
$pInvokeSignature = @'
[DllImport("user32.dll")]
public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);
'@

Add-Type -MemberDefinition $pInvokeSignature -Namespace Win32Functions -Name KeyboardFunctions

while ($true) {
    [Win32Functions.KeyboardFunctions]::keybd_event(0x5B, 0, 0, 0)
    [Win32Functions.KeyboardFunctions]::keybd_event(0x5B, 0, 2, 0)
    Start-Sleep -Seconds 2
}

```