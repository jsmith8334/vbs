Set WshShell = CreateObject("WScript.Shell")

Do

	strCmd = "powershell -Command " & Chr(34) & _
		"Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; " & _
		"public class Mouse { [DllImport(" & Chr(34) & "user32.dll" & Chr(34) & ")] public static extern bool GetCursorPos(out POINT pt); " & _
		"[DllImport(" & Chr(34) & "user32.dll" & Chr(34) & ")] public static extern bool SetCursorPos(int x, int y); " & _
		"[StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; } }'; " & _
		"$pt = New-Object Mouse+POINT; [Mouse]::GetCursorPos([ref]$pt) | Out-Null; " & _
		"[Mouse]::SetCursorPos($pt.X + 2, $pt.Y) | Out-Null; " & _
		"Start-Sleep -Milliseconds 500; " & _
		"[Mouse]::SetCursorPos($pt.X, $pt.Y) | Out-Null; " & _
		"Start-Sleep -Milliseconds 500; " & _
		"[Mouse]::SetCursorPos($pt.X + 2, $pt.Y) | Out-Null; " & _
		"Start-Sleep -Milliseconds 500; " & _
		"[Mouse]::SetCursorPos($pt.X, $pt.Y) | Out-Null;" & _
		Chr(34)

	WshShell.Run strCmd, 0, True

Loop
