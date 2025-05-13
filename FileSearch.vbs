Option Explicit

Dim strSearchTerm, strInputFolder, strExtensions, xlApp, xlBook, xlSheet, currentRow
Dim currentFolder, filePath, oShell, FSO, logFile, errorLogPath
Dim totalFilesScanned, totalMatchesFound

Set oShell = WScript.CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
currentFolder = FSO.GetParentFolderName(WScript.ScriptFullName)
errorLogPath = currentFolder & "\SearchErrors.log"
Set logFile = FSO.CreateTextFile(errorLogPath, True)
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False

strInputFolder = InputBox("Please input the folder path you wish to search:")
strSearchTerm = InputBox("Please enter the term you want to search for:")
strExtensions = LCase(InputBox("Enter file extensions to search (comma-separated, e.g., xaml,xml,txt):"))

If strInputFolder = "" OR strSearchTerm = "" OR strExtensions = "" Then
    MsgBox "Cancelled, At least one input was blank."
Else
    filePath = currentFolder & "\Search Report - " & strSearchTerm & ".xlsx"
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    xlSheet.Name = "Results"

    xlSheet.Cells(1, 1).Value = "File Path"
    xlSheet.Cells(1, 2).Value = "Search Term"
    xlSheet.Name = "Results"
    xlSheet.Cells(1, 3).Value = "Found (True/False)"
    xlSheet.Cells(1, 4).Value = "Count"
    xlSheet.Cells(1, 5).Value = "Line Numbers"

    'Format columns
    xlSheet.Columns("A").ColumnWidth = 150

    xlSheet.Columns("B").ColumnWidth = 15
    xlSheet.Columns("B").HorizontalAlignment = -4108

    xlSheet.Columns("C").ColumnWidth = 20
    xlSheet.Columns("C").HorizontalAlignment = -4108


    xlSheet.Columns("D").ColumnWidth = 15
    xlSheet.Columns("D").HorizontalAlignment = -4108

    xlSheet.Columns("E").ColumnWidth = 25
    xlSheet.Columns("E").HorizontalAlignment = -4108

    currentRow = 2
    totalFilesScanned = 0
    totalMatchesFound = 0

    LoopSubFolders FSO.GetFolder(strInputFolder), 5
    FinalizeExcelReport filePath

    logFile.Close

    MsgBox "Completed search of '" & strInputFolder & "' for term '" & strSearchTerm & "'." & vbCrLf & _
    "Total files scanned: " & totalFilesScanned & vbCrLf & _
    "Total matches found: " & totalMatchesFound
End If



'------------------ Subs & Functions ------------------

Sub LoopSubFolders(Folder, Depth)
    Dim file, Subfolder
    For Each file In Folder.Files
        totalFilesScanned = totalFilesScanned + 1
        LoopFiles file.Name, LCase(FSO.GetExtensionName(file)), file.Path, strSearchTerm
    Next

    If Depth > 0 Then
        For Each Subfolder In Folder.SubFolders
            LoopSubFolders Subfolder, Depth - 1
        Next
    End If
End Sub

Sub LoopFiles(inFileName, inFileExtension, inFilePath, inSearchTerm)
    If InStr("," & strExtensions & ",", "," & inFileExtension & ",") > 0 Then
        Dim intSearchCount, strLineNumbers
        strLineNumbers = ""
        intSearchCount = SearchFile(inFilePath, inSearchTerm, strLineNumbers)
        If intSearchCount > 0 Then
            WriteToExcelRow inFilePath, inSearchTerm, True, intSearchCount, strLineNumbers
            totalMatchesFound = totalMatchesFound + intSearchCount
        End If
    End If
End Sub

Sub WriteToExcelRow(inFilePath, inSearchTerm, inFound, inCount, inLines)
    xlSheet.Cells(currentRow, 1).Value = inFilePath
    xlSheet.Cells(currentRow, 2).Value = inSearchTerm
    xlSheet.Cells(currentRow, 3).Value = inFound
    xlSheet.Cells(currentRow, 4).Value = inCount
    xlSheet.Cells(currentRow, 5).Value = inLines
    currentRow = currentRow + 1
End Sub
  
Sub FinalizeExcelReport(filePath)
    Dim xlSummary
    Set xlSummary = xlBook.Sheets.Add
    xlSummary.Name = "Summary"
    xlSummary.Cells(1, 1).Value = "Summary"
	xlSummary.Cells(2, 1).Value = "Starting Folder"
    xlSummary.Cells(2, 2).Value = strInputFolder
    xlSummary.Cells(3, 1).Value = "Search Term"
    xlSummary.Cells(3, 2).Value = strSearchTerm
    xlSummary.Cells(4, 1).Value = "File Extensions"
    xlSummary.Cells(4, 2).Value = strExtensions
    xlSummary.Cells(5, 1).Value = "Total Files Scanned"
    xlSummary.Cells(5, 2).Value = totalFilesScanned
    xlSummary.Cells(6, 1).Value = "Total Matches Found"
    xlSummary.Cells(6, 2).Value = totalMatchesFound

    xlSummary.Columns("A").ColumnWidth = 30

    xlSummary.Columns("B").ColumnWidth = 25
    xlSummary.Columns("B").HorizontalAlignment = -4108

    xlBook.SaveAs filePath
    xlBook.Close False
    xlApp.Quit

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub

Function SearchFile(filePath, searchTerm, ByRef outLineNumbers)
    Dim text, line, pos, count, lineNum
    count = 0
    outLineNumbers = ""
    On Error Resume Next

    If FSO.FileExists(filePath) Then
        Dim currentFile
        Set currentFile = FSO.OpenTextFile(filePath, 1)
        lineNum = 0
        Do Until currentFile.AtEndOfStream
            line = currentFile.ReadLine
            lineNum = lineNum + 1
            pos = InStr(1, line, searchTerm, vbTextCompare)
            If pos > 0 Then
                count = count + 1
                outLineNumbers = outLineNumbers & lineNum & ", "
            End If
        Loop
        currentFile.Close
        If Len(outLineNumbers) > 2 Then
            If Right(outLineNumbers, 2) = ", " Then
                outLineNumbers = Left(outLineNumbers, Len(outLineNumbers) - 2)
            End If
        End If
    Else
        logFile.WriteLine "File not found: " & filePath
    End If

    If Err.Number <> 0 Then
        logFile.WriteLine "Error reading file: " & filePath & " - " & Err.Description
        Err.Clear
    End If

    On Error GoTo 0
    SearchFile = count
End Function
