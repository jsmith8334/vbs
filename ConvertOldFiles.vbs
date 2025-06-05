'Counts and logs files and their location of any extension type

'Authors: Michael Montenegro || Henry Schmandt

'------------------------------------------------

on error resume next

Dim objfso, objfile, objWScript, objExcel, objExcelBook, objWord, objWordDoc, objPowerPoint, objPowerPointPres
Dim rootFolderPath, outputFolderPath, myLog
Dim numFiles, numExcelFiles, numWordFiles, numPowerPointFiles
Dim ext, lastModified, newFilePath, logErrorMessage
Dim shouldConvert, rootFolderPathError, driveSpace, cumulativeFileSizes, driveSpaceError, conversionMessage
Dim dashedSeparator
Dim excelFilesArray(), wordFilesArray(), powerpointFilesArray(), newIndex, incrementor
Dim drive

dashedSeparator = "------------------------------------------------------------"

'Quit running instances of Word and Excel
Set objWScript = CreateObject("WScript.Shell")
objWScript.Run "taskkill /f /im EXCEL.EXE", 0
objWScript.Run "taskkill /f /im WINWORD.EXE", 0
objWScript.Run "taskkill /f /im POWERPNT.EXE", 0

Set objfso = CreateObject("Scripting.FileSystemObject")

rootFolderPath = objfso.GetParentFolderName(WScript.ScriptFullName)

If Right(rootFolderPath, 1) = "\" Then
	rootFolderPath = Left(rootFolderPath,Len(rootFolderPath)-1)
End If

Set drive = objfso.GetDrive(objfso.GetDriveName(rootFolderPath))
driveSpace = FormatNumber(drive.FreeSpace/1024, 0)

shouldConvert = Msgbox("Searching through root folder: " & rootFolderPath & vbNewLine & vbNewLine & "Click Cancel and move the script location to change this folder." & vbNewLine & vbNewLine & "Note: After clicking OK this script will search through all visible and hidden paths stemming from the root folder. It will run in the background and may take a while depending on the size of the directory. You will see a similar pop up window when this search is complete.", 1)

If shouldConvert = 1 Then
	
	outputFolderPath = rootFolderPath & "\Office Conversion Logs"
	
	numFiles = 0
	cumulativeFileSizes = 0

	'Will do everything ignoring element 0
	ReDim Preserve excelFilesArray(0)
	excelFilesArray(0) = "Ignore"
	ReDim Preserve wordFilesArray(0)
	wordFilesArray(0) = "Ignore"
	ReDim Preserve powerpointFilesArray(0)
	powerpointFilesArray(0) = "Ignore"
	
	If (objfso.FolderExists(outputFolderPath)) Then
		Call objfso.DeleteFolder(outputFolderPath,True)
	End If
	Call objfso.CreateFolder(outputFolderPath)

	Set myLog = objfso.OpenTextFile(outputFolderPath & "\Conversion_Logs.txt",8,True)

	myLog.WriteLine(dashedSeparator)
	myLog.WriteLine("Searching for old file formats (.xls, .doc, and .ppt)")
	myLog.WriteLine(dashedSeparator)

	ShowSubFolders objfso.GetFolder(rootFolderPath)
	rootFolderPathError = Err.Number

	Sub ShowSubFolders(Folder)

		on error resume next

		For Each objfile in Folder.Files
			If Err.Number <> 0 Then
				'Do Nothing
			Else
				ext = objfso.GetExtensionName(objfile.Name)
				If (ext = "xls" Or ext = "doc" Or ext = "ppt") And (InStr(objfile.path, "$RECYCLE.BIN") = 0 And InStr(objfile.path, "\~$") = 0) Then
					If ext = "xls" Then
						ext = "Excel"
					End If
					If ext = "doc" Then
						ext = "Word"
					End If
					If ext = "ppt" Then
						ext = "PowerPoint"
					End If
					myLog.WriteLine(Now & " || " & ext & " || " & objfile.path & " || File Size: " & FormatNumber(objfile.size/1024, 0) & " Kilobytes || Old File Type Found")
					If Err.Number <> 0 Then
						'Do nothing
					Else
						numFiles = numFiles + 1
						cumulativeFileSizes = cumulativeFileSizes + FormatNumber(objfile.size/1024, 0)

						If ext = "Excel" Then
							newIndex = UBound(excelFilesArray) + 1
							ReDim Preserve excelFilesArray(newIndex)
							excelFilesArray(newIndex) = objfile.path
						End If
						If ext = "Word" Then
							newIndex = UBound(wordFilesArray) + 1
							ReDim Preserve wordFilesArray(newIndex)
							wordFilesArray(newIndex) = objfile.path
						End If
						If ext = "PowerPoint" Then
							newIndex = UBound(powerpointFilesArray) + 1
							ReDim Preserve powerpointFilesArray(newIndex)
							powerpointFilesArray(newIndex) = objfile.path
						End If
					End If
				End If
			End If
		Next
		
		For Each Subfolder in Folder.SubFolders
			' Helper Code: Msgbox(Subfolder.Attributes & " " & Subfolder.Path)
			If Subfolder.Attributes = 16 Or Subfolder.Attributes = 17 Then
				ShowSubFolders Subfolder
			End If
		Next

	End Sub

	myLog.WriteLine("")
	myLog.WriteLine("Found " & UBound(excelFilesArray) & " excel file(s), " & UBound(wordFilesArray) & " word file(s), and " & UBound(powerpointFilesArray) & " powerpoint file(s) for a total of " & numFiles & " file(s) that need conversion" )
	myLog.WriteLine("")

	If cumulativeFileSizes > driveSpace Then
		driveSpaceError = 1
		conversionMessage = vbNewLine & vbNewLine & "Error: Converting these files will require " & cumulativeFileSizes & " Kilobytes of space but the current drive (" & drive.Path & ") only has " & driveSpace & " Kilobytes of free space available.  Please clear up some space or contact IT about allocating more storage to this drive."
		myLog.WriteLine("Error: Drive Space || Free space in drive " & drive.Path & " --> " & driveSpace & " Kilobytes || Free space required --> " & cumulativeFileSizes & " Kilobytes")
	Else
		driveSpaceError = 0
		conversionMessage = " Press OK to convert all of them." & vbNewLine & vbNewLine & "Note: The conversion process will run in the background and will take around 2 seconds per file. You will see a similar pop up window when the conversions are complete."
	End If

	If rootFolderPathError = 0 Then
		shouldConvert = Msgbox("There are " & UBound(excelFilesArray) & " excel file(s), " & UBound(wordFilesArray) & " word file(s), and " & UBound(powerpointFilesArray) & " powerpoint file(s) for a total of " & numFiles & " file(s) that need conversion." & conversionMessage, 1)
	Else
		shouldConvert = 0
		Msgbox("Error accessing input root folder: """ & rootFolderPath & """")
	End If
	
End If

If shouldConvert = 1 And driveSpaceError = 0 And rootFolderPathError = 0 Then

	'----------------------------------
	'	Converting Excel Files
	'----------------------------------

	myLog.WriteLine(dashedSeparator)
	myLog.WriteLine("Converting Excel files")
	myLog.WriteLine(dashedSeparator)
	
	numExcelFiles = 0

	Set objExcel = CreateObject("Excel.application")

	objExcel.visible=false
	objExcel.displayalerts=false

	ShowSubFoldersExcel

	Sub ShowSubFoldersExcel()
	
		on error resume next
		incrementor = 1

		Do While incrementor <= UBound(excelFilesArray)

			Set objfile = objfso.GetFile(excelFilesArray(incrementor))
			If Err.Number <> 0 Then
				logErrorMessage = Err.Description
				logErrorMessage = Replace(logErrorMessage, vbCr, " ")
				logErrorMessage = Replace(logErrorMessage, vbLf, " ")
				If logErrorMessage = "" Then
					logErrorMessage = "Unknown"
				End If
				myLog.WriteLine(Now & " || Error: " & logErrorMessage & " || " & excelFilesArray(incrementor))
			Else
				newFilePath = ""
				ext = objfso.GetExtensionName(objfile.Name)
				If ext = "xls" And (InStr(objfile.path, "$RECYCLE.BIN") = 0 And InStr(objfile.path, "\~$") = 0) Then
					Set objExcelBook = objExcel.Workbooks.Open(objfile)
					lastModified = Replace(Replace(objExcelBook.BuiltinDocumentProperties("Last Save Time"), "/", "-"), ":", "-")
					If Err.Number <> 0 Then
						lastModified = Replace(Replace(objfile.DateLastModified, "/", "-"), ":", "-")
					End If
					newFilePath = Replace(objfile.path, ".xls", " (" & lastModified & ").xlsx")
					If (Not objfso.FileExists(newFilePath)) Then
						objExcelBook.SaveAs newFilePath, 51
						If Err.Number <> 0 Then
							logErrorMessage = Err.Description
							logErrorMessage = Replace(logErrorMessage, vbCr, " ")
							logErrorMessage = Replace(logErrorMessage, vbLf, " ")
							If logErrorMessage = "" Then
								logErrorMessage = "Unknown"
							End If
							myLog.WriteLine(Now & " || Error: " & logErrorMessage & " || " & objfile.path & " --> " & newFilePath)
						Else
							myLog.WriteLine(Now & " || Success: Excel file converted || " & objfile.path & " --> " & newFilePath)
							If (Err.Number <> 0) Then
								'Do Nothing
							Else
								numExcelFiles = numExcelFiles + 1
							End If
						End If
					Else
						myLog.WriteLine(Now & " || Success: This Excel file had already been converted || " & objfile.path & " --> " & newFilePath)
						If (Err.Number <> 0) Then
							'Do Nothing
						Else
							numExcelFiles = numExcelFiles + 1
						End If
					End If
					objExcelBook.Close
				End If
			End If
			incrementor = incrementor + 1
		Loop
	End Sub

	myLog.WriteLine("")
	myLog.WriteLine("Successfully converted " & numExcelFiles & " excel files")
	myLog.WriteLine("")
	
	'----------------------------------
	'	Converting Word Files
	'----------------------------------

	myLog.WriteLine(dashedSeparator)
	myLog.WriteLine("Converting Word files")
	myLog.WriteLine(dashedSeparator)
	
	numWordFiles = 0
	
	Set objWord = CreateObject("Word.application")

	objWord.visible=false
	objWord.displayalerts=false

	ShowSubFoldersWord

	Sub ShowSubFoldersWord()
	
		on error resume next
		incrementor = 1

		Do While incrementor <= UBound(wordFilesArray)

			Set objfile = objfso.GetFile(wordFilesArray(incrementor))
			If Err.Number <> 0 Then
				logErrorMessage = Err.Description
				logErrorMessage = Replace(logErrorMessage, vbCr, " ")
				logErrorMessage = Replace(logErrorMessage, vbLf, " ")
				If logErrorMessage = "" Then
					logErrorMessage = "Unknown"
				End If
				myLog.WriteLine(Now & " || Error: " & logErrorMessage & " || " & wordFilesArray(incrementor))
			Else
				newFilePath = ""
				ext = objfso.GetExtensionName(objfile.Name)
				If ext = "doc" And (InStr(objfile.path, "$RECYCLE.BIN") = 0 And InStr(objfile.path, "\~$") = 0) Then
					Set objWordDoc = objWord.Documents.Open(objfile.path)
					lastModified = Replace(Replace(objfile.DateLastModified, "/", "-"), ":", "-")
					newFilePath = Replace(objfile.path, ".doc", " (" & lastModified & ").docx")
					If (Not objfso.FileExists(newFilePath)) Then
						objWordDoc.SaveAs2 newFilePath, 16
						If Err.Number <> 0 Then
							logErrorMessage = Err.Description
							logErrorMessage = Replace(logErrorMessage, vbCr, " ")
							logErrorMessage = Replace(logErrorMessage, vbLf, " ")
							If logErrorMessage = "" Then
								logErrorMessage = "Unknown"
							End If
							myLog.WriteLine(Now & " || Error: " & logErrorMessage & " || " & objfile.path & " --> " & newFilePath)
						Else
							myLog.WriteLine(Now & " || Success: Word file converted || " & objfile.path & " --> " & newFilePath)
							If (Err.Number <> 0) Then
								'Do Nothing
							Else
								numWordFiles = numWordFiles + 1
							End If
						End If
					Else
						myLog.WriteLine(Now & " || Success: This Word file had already been converted || " & objfile.path & " --> " & newFilePath)
						If (Err.Number <> 0) Then
							'Do Nothing
						Else
							numWordFiles = numWordFiles + 1
						End If
					End If
					objWordDoc.Close
				End If
			End If
			incrementor = incrementor + 1
		Loop
	End Sub

	myLog.WriteLine("")
	myLog.WriteLine("Successfully converted " & numWordFiles & " word files")
	myLog.WriteLine("")

	'---------------------------------------
	'	Converting PowerPoint Files
	'---------------------------------------

	myLog.WriteLine(dashedSeparator)
	myLog.WriteLine("Converting PowerPoint files")
	myLog.WriteLine(dashedSeparator)
	
	numPowerPointFiles = 0
	
	Set objPowerPoint = CreateObject("PowerPoint.application")

	ShowSubFoldersPowerPoint

	Sub ShowSubFoldersPowerPoint()
	
		on error resume next
		incrementor = 1

		Do While incrementor <= UBound(powerpointFilesArray)
			
			Set objfile = objfso.GetFile(powerpointFilesArray(incrementor))
			If Err.Number <> 0 Then
				logErrorMessage = Err.Description
				logErrorMessage = Replace(logErrorMessage, vbCr, " ")
				logErrorMessage = Replace(logErrorMessage, vbLf, " ")
				If logErrorMessage = "" Then
					logErrorMessage = "Unknown"
				End If
				myLog.WriteLine(Now & " || Error: " & logErrorMessage & " || " & powerpointFilesArray(incrementor))
			Else
				newFilePath = ""
				ext = objfso.GetExtensionName(objfile.Name)
				If ext = "ppt" And (InStr(objfile.path, "$RECYCLE.BIN") = 0 And InStr(objfile.path, "\~$") = 0) Then
					Set objPowerPointPres = objPowerPoint.Presentations.Open(objfile.path, 0, 0, 0)
					lastModified = Replace(Replace(objfile.DateLastModified, "/", "-"), ":", "-")
					newFilePath = Replace(objfile.path, ".ppt", " (" & lastModified & ").pptx")
					If (Not objfso.FileExists(newFilePath)) Then
						objPowerPointPres.SaveAs newFilePath
						If Err.Number <> 0 Then
							logErrorMessage = Err.Description
							logErrorMessage = Replace(logErrorMessage, vbCr, " ")
							logErrorMessage = Replace(logErrorMessage, vbLf, " ")
							If logErrorMessage = "" Then
								logErrorMessage = "Unknown"
							End If
							myLog.WriteLine(Testing & " || Error: " & logErrorMessage)
							myLog.WriteLine(Now & " || Error: " & logErrorMessage & " || " & objfile.path & " --> " & newFilePath)
						Else
							myLog.WriteLine(Now & " || Success: PowerPoint file converted || " & objfile.path & " --> " & newFilePath)
							If (Err.Number <> 0) Then
								'Do Nothing
							Else
								numPowerPointFiles = numPowerPointFiles + 1
							End If
						End If
					Else
						myLog.WriteLine(Now & " || Success: This PowerPoint file had already been converted || " & objfile.path & " --> " & newFilePath)
						If (Err.Number <> 0) Then
							'Do Nothing
						Else
							numPowerPointFiles = numPowerPointFiles + 1
						End If
					End If
					objPowerPointPres.Close
				End If
			End If
			incrementor = incrementor + 1
		Loop
	End Sub

	myLog.WriteLine("")
	myLog.WriteLine("Successfully converted " & numPowerPointFiles & " powerpoint files")
	myLog.Close
	
	dim numErrors, errorMessage
	numErrors = numFiles - numExcelFiles - numWordFiles - numPowerPointFiles
	
	If numErrors <> 0 Then
		errorMessage = "  Errors were encountered converting " & numErrors & " file(s)."
	Else
		errorMessage = ""
	End If

	Msgbox("Script complete. " & numExcelFiles & " excel file(s), " & numWordFiles & " word file(s), and " & numPowerPointFiles & " powerpoint file(s) were converted." & errorMessage & "  Logs of all conversions can be found at:" & vbNewLine & vbNewLine & outputFolderPath)
Else
	myLog.Close
End If