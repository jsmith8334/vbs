Option Explicit

'Declaring variables
Dim objWord, objDoc, objRange, objLink, objExcel, objWorkbook, objWorksheet, objCell, strScriptDir, objFSO, strFile, oShell
Dim intRow, strTempFileDir, strCurrentFilename, objShell, strCompleted, strNotSupported, strScriptDirPath

strScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.shell")
Set oShell = WScript.CreateObject ("WScript.Shell")


'''''Start Main Code'''''

oShell.run "taskkill /f /im excel.exe", 0, True
oShell.run "taskkill /f /im word.exe", 0, True

MsgBox "The Link Checker will run once you click 'OK'. This process will take some time, so please do not open any files or folders within the '" + CStr(strScriptDir) + "' parent folder.  You will see another notification once the process has completed.", vbOKOnly, "Broken Link Checker"

Call CreateFolder("Completed")
Call CreateFolder("Final Reports")
Call CreateFolder("Not Supported")

'Loop through parent directory
For Each strFile In objFSO.GetFolder(strScriptDir).Files
	'Split filename into name var and extension var	
		Dim strFilename, strFileExtension, strExcelFilename
		strFilename = objFSO.GetBaseName(strFile)
		strFileExtension = objFSO.GetExtensionName(strFile) 
	'Validate the filetype
		If strFilename = "Link Checker" Then 
			'Do nothing
		ElseIf (strFileExtension = "xlsx") OR (strFileExtension = "xls") OR (strFileExtension = "doc") Or (strFileExtension = "docx") OR (strFileExtension = "pdf") Then
			strExcelFilename = strScriptDir + "\" + "Final Reports" + "\" + strFileName + " - Link Report.xlsx"
			If (strFileExtension = "xlsx") OR (strFileExtension = "xls") Then 
				Call findLinksExcel(strFile, strExcelFilename)
			ElseIf (strFileExtension = "doc") OR (strFileExtension = "docx") Then 
				Call findLinksWord(strFile, strExcelFilename)
			ElseIf strFileExtension = "pdf" Then 
				Call findLinksPDF(strFile, strExcelFilename)
			Else
				Exit For
			End If
			
			'Check links
			Call checkURL(strExcelFilename)
		
			'Take screenshots
			'Where to save, how to display in report?
			
			'Finalize report
			Call formatExcel(strExcelFilename)			
			
			'Move completed file to completed folder
			Call moveFile(strFile, strScriptDir + "\Completed\")
			'MsgBox "Completed processing " + strFilename + "." + strFileExtension + "."
		Else
			'WScript.Echo("The file " + strFile + " is an invalid filetype.")
			Call moveFile(strFile, strScriptDir + "\Not Supported\")
		End If
Next


'Overall clean Up


'Cleanup
Set objLink = Nothing
Set objRange = Nothing
Set objDoc = Nothing
Set objWord = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
Set objCell = Nothing

MsgBox "The Broken Link Checker process has completed!", vbOKOnly, "Broken Link Checker"

'''''End Main Code'''''




'Function to open and loop through exel
Function findLinksExcel(strFilename, strExcelFilename)
	On Error Resume Next
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(CStr(strFilename))
	Set objWorksheet = objWorkbook.Worksheets(1)
	Set objRange = objWorksheet.UsedRange
	Set objWorkbook = objExcel.Workbooks.Add
	Set objWorksheet = objWorkbook.Worksheets.Add
	intRow = 2
	objWorksheet.Cells(1, 1).Value = "LinkID"
	objWorksheet.Cells(1, 2).Value = "Link Name"
	objWorksheet.Cells(1, 3).Value = "Link URL"
	objWorksheet.Cells(1, 4).Value = "Link Check"
	objWorksheet.Cells(1, 5).Value = "Check Code"
	For Each objLink In objRange.Hyperlinks
		If InStr(objLink.Address, "http") = 1 Then
			If Len(objLink.TextToDisplay) < 3 Then
			'If (objLink.TextToDisplay = " ") OR (objLink.TextToDisplay = "	") OR (objLink.TextToDisplay = "") OR (objLink.TextToDisplay = ",") Or (objLink.TextToDisplay = ".") Or (objLink.TextToDisplay = ";") Or (objLink.TextToDisplay = ":") Then
				Continue	
			Else 
				objWorksheet.Cells(intRow, 1).Value = intRow - 1
				objWorksheet.Cells(intRow, 2).Value = objLink.TextToDisplay
				objWorksheet.Cells(intRow, 3).Value = "=HYPERLINK(""" & objLink.Address & """)"
				intRow = intRow + 1
			End If
		Else 
			Continue
		End If
	Next

	objWorkbook.SaveAs CStr(strExcelFilename)
	objWorkbook.Close
	objExcel.Quit

End Function

'Function to open and loop through docx or pdf
Function findLinksWord(strFilename, strExcelFilename)
	On Error Resume Next
	Set objWord = CreateObject("Word.Application")
	Set objExcel = CreateObject("Excel.Application")
	Set objDoc = objWord.Documents.Open(CStr(strFilename))
	Set objWorkbook = objExcel.Workbooks.Add
	Set objWorksheet = objWorkbook.Worksheets(1)
	intRow = 2
	For Each objLink In objDoc.Hyperlinks
		objWorksheet.Cells(1, 1).Value = "LinkID"
		objWorksheet.Cells(1, 2).Value = "Link Name"
		objWorksheet.Cells(1, 3).Value = "Link URL"
		objWorksheet.Cells(1, 4).Value = "Link Check"
		objWorksheet.Cells(1, 5).Value = "Check Code"
		If InStr(objLink.Address, "http") = 1 Then
			If Len(objLink.TextToDisplay) < 3 Then
			'If (objLink.TextToDisplay = " ") Or (objLink.TextToDisplay = "	") Or (objLink.TextToDisplay = "") Or (objLink.TextToDisplay = ",") Or (objLink.TextToDisplay = ".") Or (objLink.TextToDisplay = ";") Or (objLink.TextToDisplay = ":") Or (objLink.TextToDisplay = "(") Or (objLink.TextToDisplay = ")") Or (objLink.TextToDisplay = "-") Or (objLink.TextToDisplay = "_") Then
				Continue
			Else
				objWorksheet.Cells(intRow, 1).Value = intRow - 1
				objWorksheet.Cells(intRow, 2).Value = objLink.TextToDisplay
				objWorksheet.Cells(intRow, 3).Value = "=HYPERLINK(""" & objLink.Address & """)"
				intRow = intRow + 1
			End If
		Else 
			Continue
        End If
	Next
	
	objWorkbook.SaveAs CStr(strExcelFilename)
	objDoc.Close
	objWord.Quit
	objWorkbook.Close
	objExcel.Quit

End Function

'Function to open and loop through docx or pdf
Function findLinksPDF(strFilename, strExcelFilename)
	On Error Resume Next
	Set objWord = CreateObject("Word.Application")
	Set objExcel = CreateObject("Excel.Application")
	Set objDoc = objWord.Documents.Open(CStr(strFilename))
	Set objWorkbook = objExcel.Workbooks.Add
	Set objWorksheet = objWorkbook.Worksheets(1)
	objWorksheet.Cells(1, 1).Value = "LinkID"
	objWorksheet.Cells(1, 2).Value = "Link Name"
	objWorksheet.Cells(1, 3).Value = "Link URL"
	objWorksheet.Cells(1, 4).Value = "Link Check"
	objWorksheet.Cells(1, 5).Value = "Check Code"
	
	Dim writeLinkName, writeLinkURL, linkNum
	linkNum = 1
	
	For Each objLink in objDoc.Hyperlinks
		If InStr(currentLink.Address, "http") = 1 Then
			If Len(currentLink.TextToDisplay) < 3 Then
				If writeLinkName = "" Then 
					writeLinkName = objLink.TextToDisplay
					writeLinkURL = objLink.Address
				ElseIf writeLinkURL = objLink.Address Then
					writeLinkName = writeLinkName & " " & objLink.TextToDisplay
				Else 
					objWorksheet.Cells(linkNum + 1, 1).Value = linkNum
					objWorksheet.Cells(linkNum + 1, 2).Value = writeLinkName
					objWorksheet.Cells(linkNum + 1, 3).Value = "=HYPERLINK(""" & writeLinkURL & """)"
					writeLinkName = objLink.TextToDisplay
					writeLinkURL = objLink.Address
					linkNum = linkNum + 1
				End If
			Else 
				Continue
			End If
		Else 
			Continue
		End If
	Next
	
	objWorkbook.SaveAs CStr(strExcelFilename)
	objDoc.Close
	objWord.Quit
	objWorkbook.Close
	objExcel.Quit

End Function

'Function to loop through URLs
Function checkURL(strExcelFilename)
	On Error Resume Next
	DIm o
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(strExcelFilename)
	Set objWorksheet = objWorkbook.Worksheets(1)
	Set objRange = objWorksheet.UsedRange
	Set o = CreateObject("MSXML2.ServerXMLHTTP")
	For intRow = 2 To objRange.Rows.Count
		Set objCell = objWorksheet.Cells(intRow, 3)
		o.open "GET", CStr(objCell.Value), False
		xmlHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
		o.send
		If o.Status = "" Then
			objWorksheet.Cells(intRow, 4).Value = "Internal Document Link"
		ElseIf o.Status = " " Then
			objWorksheet.Cells(intRow, 4).Value = "Internal Document Link"
		ElseIf o.Status = 200 Then
			objWorksheet.Cells(intRow, 4).Value = "Valid"
		ElseIf o.Status = 301 Then
			objWorksheet.Cells(intRow, 4).Value = "Redirect URL: Update"
		ElseIf o.Status = 401 Then
			objWorksheet.Cells(intRow, 4).Value = "Unauthorized: Check Manually"
		ElseIf o.Status  = 403 Then
			objWorksheet.Cells(intRow, 4).Value = "Forbidden: Check Manually"
		ElseIf o.Status  = 405 Then
			objWorksheet.Cells(intRow, 4).Value = "Requires POST,GET,PULL"
		ElseIf o.Status  = 0 Then 
			objWorksheet.Cells(intRow, 4).Value = "No Response: Check Manually"
		Else
			objWorksheet.Cells(intRow, 4).Value = "Invalid: Fix Link"
		End If
		objWorksheet.Cells(intRow, 5).Value = CStr(o.Status)
	Next

	objWorkbook.Save
	objWorkbook.Close
	objExcel.Quit
	
End Function

Function formatExcel(excelFile)
	Dim objXLA
	Set objXLA = CreateObject("Excel.Application")
	Set objWorkbook = objXLA.Workbooks.Open(excelFile)
	objXLA.Application.DisplayAlerts = False
	
	objXLA.Range("A1:A1").Select
	objXLA.Selection.ColumnWidth = 10
		
	objXLA.Range("B1:B1").Select
	objXLA.Selection.ColumnWidth = 45
	objXLA.Selection.HorizontalAlignment = -4131
	
	objXLA.Range("C1:C1").Select
	objXLA.Selection.ColumnWidth = 65
	objXLA.Selection.HorizontalAlignment = -4131
	
	objXLA.Range("D1:D1").Select
	objXLA.Selection.ColumnWidth = 28
	
	objXLA.Range("E1:E1").Select
	objXLA.Selection.ColumnWidth = 15
	
	objXLA.Range("A1:A1000").Select
	objXLA.Selection.HorizontalAlignment = &HFFFFEFF4
	objXLA.Range("D1:D1000").Select
	objXLA.Selection.HorizontalAlignment = &HFFFFEFF4
	objXLA.Range("E1:E1000").Select
	objXLA.Selection.HorizontalAlignment = &HFFFFEFF4
	
	objXLA.Range("A1:E1").Select
	objXLA.Selection.Font.Name = "Arial"
	objXLA.Selection.Font.Size = 12
	objXLA.Selection.Font.Bold = True
	objXLA.Selection.HorizontalAlignment = &HFFFFEFF4
	objXLA.Selection.Interior.ColorIndex = 15
	'objXLA.Selection.NumberFormat = "@"
	objXLA.Selection.RowHeight = 16
	objXLA.Selection.Borders.LineStyle = 1
	
	objWorkbook.Save
	objWorkbook.Close
	objXLA.Quit
End Function

' Function to check if a URL exists
Function UrlExists(strUrl)
    Dim objHttp, intStatus
    On Error Resume Next
    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
    objHttp.Open "HEAD", strUrl, False
    objHttp.Send
	UrlExists = objHttp.Status
    Set objHttp = Nothing
    On Error Goto 0
End Function

' Function to create folder
Function CreateFolder(folderPath)
    Dim fso, folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then
        CreateFolder = False
        Exit Function
    End If
    Set folder = fso.CreateFolder(folderPath)
    If fso.FolderExists(folderPath) Then
        CreateFolder = True
    Else
        CreateFolder = False
    End If    
    Set folder = Nothing
    Set fso = Nothing
End Function

' Function move file to folder
Function moveFile(file, folder)
	objFSO.MoveFile file, folder
End Function


'Function to delete everything in a folder
Function clearFolder(folderPath)
Dim fso, folder, f
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(folderPath)
for each f in folder.Files
   On Error Resume Next
   name = f.name
   f.Delete True
   On Error GoTo 0
Next
End Function