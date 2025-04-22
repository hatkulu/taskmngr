Set WshShell = CreateObject("WScript.Shell")
Do
    WshShell.Run "prank.vbs", 0, False
    WScript.Sleep 3000
Loop
