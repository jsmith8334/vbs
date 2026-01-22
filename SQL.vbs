'====================================================================
' VBScript: Connect to SQL Server, Run Query, Export to Text File
' Save as: ExportQueryToText.vbs
'====================================================================

Option Explicit

Dim conn, rs, cmd
Dim strConnection, strQuery, strOutputFile
Dim fso, outputFile
Dim i, fieldCount

' ==================== CONFIGURE THESE VALUES ======================
strConnection = "Provider=SQLOLEDB;Data Source=YOUR_SERVER_NAME;Initial Catalog=YOUR_DATABASE;Integrated Security=SSPI;" 
' For SQL Authentication, use this instead:
' strConnection = "Provider=SQLOLEDB;Data Source=YOUR_SERVER_NAME;Initial Catalog=YOUR_DATABASE;User ID=your_user;Password=your_password;"

strQuery = "SELECT TOP 100 CustomerID, CompanyName, ContactName, City, Country FROM Customers"  ' <-- Change your query here

strOutputFile = "C:\Temp\QueryResults.txt"   ' <-- Change path as needed
' ==================================================================

' Create objects
Set conn = CreateObject("ADODB.Connection")
Set rs   = CreateObject("ADODB.Recordset")
Set fso  = CreateObject("Scripting.FileSystemObject")
Set outputFile = fso.CreateTextFile(strOutputFile, True, True)  ' True = overwrite, True = Unicode

On Error Resume Next

' Open connection
conn.Open strConnection

If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot connect to database!" & vbCrLf & Err.Description
    WScript.Quit
End If
On Error Goto 0

' Execute query
Set rs = conn.Execute(strQuery)

If rs.State = 0 Then
    outputFile.WriteLine "Query returned no results or failed."
    outputFile.Close
    conn.Close
    WScript.Echo "Query failed or returned no rows."
    WScript.Quit
End If

' Write column headers
For i = 0 To rs.Fields.Count - 1
    If i > 0 Then outputFile.Write vbTab
    outputFile.Write rs.Fields(i).Name
Next
outputFile.WriteLine

' Write data rows
Do While Not rs.EOF
    For i = 0 To rs.Fields.Count - 1
        If i > 0 Then outputFile.Write vbTab
        
        If IsNull(rs.Fields(i).Value) Then
            outputFile.Write ""   ' Write empty for NULL
        Else
            ' Replace tabs and newlines in data to avoid breaking format
            Dim fieldValue
            fieldValue = Replace(rs.Fields(i).Value, vbTab, " ")
            fieldValue = Replace(fieldValue, vbCrLf, " ")
            fieldValue = Replace(fieldValue, vbCr, " ")
            fieldValue = Replace(fieldValue, vbLf, " ")
            outputFile.Write fieldValue
        End If
    Next
    outputFile.WriteLine
    rs.MoveNext
Loop

' Cleanup
rs.Close
conn.Close
outputFile.Close

Set rs = Nothing
Set conn = Nothing
Set fso = Nothing
Set outputFile = Nothing

WScript.Echo "Query executed successfully!" & vbCrLf & "Results saved to: " & strOutputFile