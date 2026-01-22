' hollywood_hack.vbs - A fun VBScript to simulate a Hollywood-style hacking terminal

Option Explicit
Dim WshShell, command, i, delay
Set WshShell = CreateObject("WScript.Shell")

' Open a Command Prompt window with a green text color
WshShell.Run "cmd.exe /K color 0a" & Chr(13) & "cls", 1, False
WScript.Sleep 500 ' Wait a moment for the window to open

For i = 1 to 500 ' Run for 500 lines (adjust as needed)
    command = "echo [INFO] Processing sector " & i & " of 500... Status: " & GetRandomStatus()
    WshShell.Run "cmd.exe /C " & command, 0, True ' Run a silent command

    command = "echo [SUCCESS] Operation complete: " & GetRandomOperation() & " on server " & GetRandomIP()
    WshShell.Run "cmd.exe /C " & command, 0, True

    command = "echo [ALERT] Unusual activity detected from " & GetRandomIP() & "... Monitoring connection."
    WshShell.Run "cmd.exe /C " & command, 0, True

    ' Add a variable delay to make it look more organic
    delay = Int((50 * Rnd) + 10)
    WScript.Sleep delay
Next

WshShell.Run "cmd.exe /C echo [COMPLETE] Operation finished. System status nominal. Press any key to exit.", 0, True

Function GetRandomStatus()
    Dim statuses(3)
    statuses(0) = "OK"
    statuses(1) = "RUNNING"
    statuses(2) = "PENDING"
    statuses(3) = "SECURE"
    GetRandomStatus = statuses(Int(4 * Rnd))
End Function

Function GetRandomOperation()
    Dim operations(3)
    operations(0) = "Data integrity check"
    operations(1) = "Log file analysis"
    operations(2) = "Network traffic simulation"
    operations(3) = "Security protocol test"
    GetRandomOperation = operations(Int(4 * Rnd))
End Function

Function GetRandomIP()
    GetRandomIP = Int(256 * Rnd) & "." & Int(256 * Rnd) & "." & Int(256 * Rnd) & "." & Int(256 * Rnd)
End Function