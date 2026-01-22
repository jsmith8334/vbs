'Function Syntax
Function AddNumbers(ByVal num1, ByVal num2)
    Dim sum
    sum = num1 + num2
    AddNumbers = sum
End Function

Result = AddNumbers(5, 6)
Wscript.Echo "The sum is: " & result

'=============================================================================================================

'Sub Syntax
Sub displayMsg()
	Wscript.Echo "Message."
End Sub

'=============================================================================================================

'Display System Details
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
 
For Each objItem in colItems
    MsgBox "Operating System: " & objItem.Caption & vbCrLf & _
           "Version: " & objItem.Version & vbCrLf & _
           "Service Pack: " & objItem.ServicePackMajorVersion & "." & objItem.ServicePackMinorVersion, vbInformation, "System Information"
Next

'=============================================================================================================

'Backup Registry
strBackupDir = "C:\RegistryBackups"
Set objShell = CreateObject("WScript.Shell")
strRegFile = objShell.ExpandEnvironmentStrings("%TEMP%") & "\registry_backup.reg"
 
strCommand = "reg export HKCU C:\RegistryBackups\HKCU_backup.reg /y"
objShell.Run strCommand, 0, True
 
MsgBox "Registry backup created successfully.", vbInf

'=============================================================================================================

'Disable Windows Defender
Set objShell = CreateObject("WScript.Shell")
strCommand = "powershell.exe -Command ""Set-MpPreference -DisableRealtimeMonitoring $true"""
objShell.Run strCommand, 0, True
MsgBox "Windows Defender disabled successfully.", vbInformation, "Success"

'=============================================================================================================

'Clear Temporary Files
Set objFSO = CreateObject("Scripting.FileSystemObject")
strTempFolder = objFSO.GetSpecialFolder(2)
Set objFolder = objFSO.GetFolder(strTempFolder)
 
For Each objFile In objFolder.Files
    objFile.Delete
Next
 
MsgBox "Temporary files cleared successfully.", vbInformation, "Success"


'=============================================================================================================

'Create System Restor Point
strDesc = "My Restore Point"
strResult = CreateRestorePoint(strDesc, 100, 7)
If strResult = 0 Then
    MsgBox "System Restore Point created successfully.", vbInformation, "Success"
Else
    MsgBox "Failed to create System Restore Point.", vbExclamation, "Error"
End If
 
Function CreateRestorePoint(strDesc, intType, intEventType)
    Set objSRP = GetObject("winmgmts:\\.\root\default:Systemrestore")
    objSRP.CreateRestorePoint strDesc, intType, intEventType
    CreateRestorePoint = Err.Number
End Function

'=============================================================================================================

' MSGBox Icons and buttons.  The numbers there define what buttons and icons appear.  For the buttons:
' 0 - ok button
' 1 - ok and cancel
' 2 - abort, retry and ignore
' 3 - yes no and cancel
' 4 - yes and no
' 5 - retry and cancel
' For the icons:
' 16 - critical message icon
' 32 - warning icon
' 48 - warning message
' 64 - info message
' To use these, simply add the numbers to the code like this: If I wanted a warning icon and yes and no buttons, I would write this:
'MsgBox "Is this a warning?", 4+32 ,"Warning!" 

'=============================================================================================================

' Computer Speak
set speech = Wscript.CreateObject("SAPI.spVoice")
speech.speak "Hello, I am ready to do your bidding."

'=============================================================================================================

'=============================================================================================================

'=============================================================================================================
