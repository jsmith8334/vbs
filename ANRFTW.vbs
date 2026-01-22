' File: PreventLock_RPA_Safe.vbs
' No keyboard input – Safe for UiPath, Automation Anywhere, Blue Prism, etc.
' Tiny invisible mouse wiggle every 59 seconds

Set WshShell = CreateObject("WScript.Shell")

Do
    WScript.Sleep 59000   ' 59 seconds – keeps you under the usual 1–15 minute lock policies

    ' Move mouse +1 pixel (invisible)
    WshShell.Run "powershell -command ""$p=[System.Windows.Forms.Cursor]::Position; [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($p.X + 1), $p.Y)""", 0, True
    
    WScript.Sleep 300
    
    ' Move back -1 pixel (net movement = 0)
    WshShell.Run "powershell -command ""$p=[System.Windows.Forms.Cursor]::Position; [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($p.X - 1), $p.Y)""", 0, True

Loop