'=======================================================================
'  Folder Mirror + Move (Copy if needed → Delete source if destination exists)
'  - Skips files that already exist in destination
'  - Deletes source only after confirming file exists in destination
'  - Silent operation (no per-file messages)
'=======================================================================
Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' ────────────────────────────────────────────────
'  CONFIGURATION - change these paths
' ────────────────────────────────────────────────
Const SOURCE_FOLDER   = "W:\ImageNow Downloads\CF SQ SSC"
Const TARGET_ROOT     = "G:\Shared drives\ImageNow Contract Files 11-5-25\CF SQ SSC"
' ────────────────────────────────────────────────

Main SOURCE_FOLDER, TARGET_ROOT

WScript.Echo "Operation completed."
' WScript.Sleep 1500   ' optional - small pause before window closes

' ===============================================================
Sub Main(srcPath, tgtRoot)
' ===============================================================
    If Not fso.FolderExists(srcPath) Then Exit Sub
    
    Dim srcFolder
    Set srcFolder = fso.GetFolder(srcPath)
      
	If Not fso.FolderExists(tgtRoot) Then
		fso.CreateFolder tgtRoot
	End If
    
    ProcessFolder srcFolder, tgtRoot
End Sub

' ===============================================================
Sub ProcessFolder(srcFolder, tgtParentPath)
' ===============================================================
    Dim tgtFolderPath
    tgtFolderPath = BuildTargetPath(srcFolder.Path, SOURCE_FOLDER, tgtParentPath)
    
    ' Create destination folder if needed
    If Not fso.FolderExists(tgtFolderPath) Then
		fso.CreateFolder tgtFolderPath
    End If
    
    ' ───── Files ───────────────────────────────────────
    Dim file
    For Each file In srcFolder.Files
        Dim targetFilePath
        targetFilePath = tgtFolderPath & "\" & file.Name
        
        If fso.FileExists(targetFilePath) Then
            ' Already exists → check if identical → delete source if safe
            If FilesAreIdentical(file.Path, targetFilePath) Then
                On Error Resume Next
                file.Delete True
                On Error GoTo 0
            End If
            ' else: different content → leave source as-is
        Else
            ' Does NOT exist → copy then (if successful) delete source
            On Error Resume Next
            file.Copy targetFilePath, False
            If Err.Number = 0 Then
                If fso.FileExists(targetFilePath) Then
                    file.Delete True
                End If
            End If
            On Error GoTo 0
        End If
    Next
    
    ' ───── Subfolders ──────────────────────────────────
    Dim subFolder
    For Each subFolder In srcFolder.SubFolders
        ProcessFolder subFolder, tgtFolderPath
    Next
End Sub

' ===============================================================
Function BuildTargetPath(fullSourcePath, baseSourcePath, targetRoot)
' ===============================================================
    Dim relPath
    If Right(baseSourcePath, 1) <> "\" Then baseSourcePath = baseSourcePath & "\"
    
    relPath = Mid(fullSourcePath, Len(baseSourcePath) + 1)
    If Left(relPath, 1) = "\" Then relPath = Mid(relPath, 2)
    
    If relPath = "" Then
        BuildTargetPath = targetRoot
    Else
        BuildTargetPath = targetRoot & "\" & relPath
    End If
End Function

' ===============================================================
Function FilesAreIdentical(path1, path2)
' ===============================================================
    Dim f1, f2
    Set f1 = fso.GetFile(path1)
    Set f2 = fso.GetFile(path2)
    
    If f1.Size <> f2.Size Then
        FilesAreIdentical = False
        Exit Function
    End If
    
    If f1.DateLastModified <> f2.DateLastModified Then
        FilesAreIdentical = False
        Exit Function
    End If
    
    FilesAreIdentical = True
End Function