Option Explicit

Dim fso, startFolder, otherFolder
Set fso = CreateObject("Scripting.FileSystemObject")

' === CHANGE THIS PATH TO YOUR STARTING FOLDER ===
' Tip: end with \ or not — both are fine
startFolder = "W:\ImageNow Downloads\CF SQ SSC"
otherFolder = "W:\ImageNow Downloads\BPA"

If Not fso.FolderExists(startFolder) Then
    WScript.Echo "Starting folder not found: " & startFolder
    WScript.Quit
End If

CleanEmptyFolders startFolder
CleanEmptyFolders otherFolder

WScript.Echo "Done."

'============================================================
Sub CleanEmptyFolders(folderPath)
    Dim folder, subFolder
    
    If Not fso.FolderExists(folderPath) Then Exit Sub
    
    Set folder = fso.GetFolder(folderPath)
    
    ' 1. First clean all children (depth-first)
    For Each subFolder In folder.SubFolders
        CleanEmptyFolders subFolder.Path
    Next
    
    ' 2. Only now check & delete current folder if empty
    If folder.Files.Count = 0 And folder.SubFolders.Count = 0 Then
        On Error Resume Next
        folder.Delete
        If Err.Number <> 0 Then
            WScript.Echo "Could not delete: " & folder.Path & vbCrLf & _
                         "Error " & Err.Number & " - " & Err.Description
        End If
        On Error GoTo 0
    End If
End Sub