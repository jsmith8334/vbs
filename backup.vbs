Option Explicit

Dim objFSO, sourcePath, destinationPath, overwriteFiles
Const OverwriteExisting = True

sourcePath = InputBox("Please enter the folder that you wish to backup: ", "User Input")
destinationPath = InputBox("Please enter the location path to where the backup will be saved: ", "User Input")
' ---------------------

Set objFSO = CreateObject("Scripting.FileSystemObject")

If Not objFSO.FolderExists(sourcePath) Then
    Wscript.Echo "Source folder not found: " & sourcePath
    Wscript.Quit
End If

If Not objFSO.FolderExists(destinationPath) Then
    objFSO.CreateFolder(destinationPath)
    Wscript.Echo "Created destination folder: " & destinationPath
End If

On Error Resume Next
'objFSO.CopyFolder sourcePath & "\*", destinationPath & "\", OverwriteExisting
objFSO.CopyFile sourcePath & "\*", destinationPath & "\", OverwriteExisting

If Err.Number <> 0 Then
    Wscript.Echo "Backup failed with error: " & Err.Description
    Err.Clear
Else
    Wscript.Echo "Backup of " & sourcePath & " to " & destinationPath & " completed successfully."
End If
On Error Goto 0

Set objFSO = Nothing