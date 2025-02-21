Option Explicit

Dim sourceFolder, destinationFolder, fileExtension
'sourceFolder = InputBox("Please enter the folder you wish to search: (ex. C:\Documents", "User Input")
sourceFolder = "C:\Users\jonathanrsmith\.nuget\packages"
destinationFolder = InputBox("Please enter the folder you wish to copy the found files to: (ex. C:\Temp", "User Input")
'destinationFolder = "\\ent.ds.gsa.gov\vdi_apps\Arpit\Packages\10-24-2023"
'fileExtension = InputBox("Please enter the file extension you wish to search for with no special characters: (ex. pdf or nupkg)", "User Input")
fileExtension = "nupkg"

CopyFiles sourceFolder, destinationFolder, fileExtension
WScript.Echo "File copy completed."

Sub CopyFiles(folderPath, destFolder, ext)
     Dim fso, folder, subfolder, files, file, destPath
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set folder = fso.GetFolder(folderPath)
     Set files = folder.Files
     For Each file In files
         If LCase(fso.GetExtensionName(file.Path)) = LCase(ext) Then
             destPath = fso.BuildPath(destFolder, fso.GetFileName(file.Path))
             fso.CopyFile file.Path, destPath, True
         End If
     Next

     For Each subfolder In folder.SubFolders
         CopyFiles subfolder.Path, destFolder, ext
     Next
End Sub
