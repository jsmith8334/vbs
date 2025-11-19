Option Explicit

Dim objFSO, objFolder, objFile, strSourcePath, strExtension, strDestFolder

' --- Configuration ---
' Set the path of the folder you want to organize
strSourcePath = "C:\Users\YourUser\Downloads" 
' ---------------------

Set objFSO = CreateObject("Scripting.FileSystemObject")

If Not objFSO.FolderExists(strSourcePath) Then
    Wscript.Echo "Source folder not found: " & strSourcePath
    Wscript.Quit
End If

Set objFolder = objFSO.GetFolder(strSourcePath)

' Iterate through all files in the source folder
For Each objFile In objFolder.Files
    ' Get the file extension and use it as the destination folder name (without the dot)
    strExtension = objFSO.GetExtensionName(objFile.Name)
    
    If strExtension = "" Then
        strDestFolder = objFSO.BuildPath(strSourcePath, "No_Extension")
    Else
        ' Convert extension to uppercase for consistency
        strDestFolder = objFSO.BuildPath(strSourcePath, UCase(strExtension))
    End If
    
    ' Create the destination folder if it doesn't exist
    If Not objFSO.FolderExists(strDestFolder) Then
        objFSO.CreateFolder(strDestFolder)
    End If
    
    ' Move the file
    ' Use MoveFile; if the file name already exists in the destination, this will error.
    ' A robust script might check for existence or use CopyFile followed by DeleteFile.
    On Error Resume Next ' Simple error handling for existing files
    objFSO.MoveFile objFile.Path, objFSO.BuildPath(strDestFolder, objFile.Name)
    
    If Err.Number <> 0 Then
        Wscript.Echo "Could not move file: " & objFile.Name & " (Error: " & Err.Description & ")"
        Err.Clear
    Else
        ' Optional: log moved files
        ' Wscript.Echo "Moved: " & objFile.Name & " to " & strDestFolder
    End If
    On Error Goto 0
Next

Wscript.Echo "File organization complete in " & strSourcePath

Set objFile = Nothing
Set objFolder = Nothing
Set objFSO = Nothing