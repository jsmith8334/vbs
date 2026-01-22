'=====================================================================
' LinkChecker.vbs - VBScript port of your Python link checker
' Save as LinkChecker.vbs and run with wscript.exe or cscript.exe
' Requires: config.xml in the same folder
'=====================================================================

Option Explicit

Const XML_FILE         = "config.xml"
Const REQUEST_TIMEOUT  = 20        ' seconds
Const SLEEP_BETWEEN_PAGES = 1000   ' ms
Const SLEEP_BETWEEN_TESTS = 500    ' ms

' ========================= CONFIGURATION =========================
Const SMTP_SERVER      = "smtp.gsa.gov"
Const SMTP_PORT        = 25
Const SENDER_EMAIL     = "bot@gsa.gov"
' No password needed for unauthenticated internal SMTP (GSA)
' =================================================================

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim xmlDoc : Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
xmlDoc.async = False
xmlDoc.setProperty "ProhibitDTD", False
If Not xmlDoc.Load(fso.GetParentFolderName(WScript.ScriptFullName) & "\" & XML_FILE) Then
    MsgBox "Cannot load " & XML_FILE, vbCritical
    WScript.Quit
End If

Dim teamNode
For Each teamNode In xmlDoc.SelectNodes("//team")
    ProcessTeam teamNode
Next

'=====================================================================
Sub ProcessTeam(teamNode)
    Dim teamName : teamName = GetText(teamNode.SelectSingleNode("name"))
    Dim owner    : owner    = GetText(teamNode.SelectSingleNode("process_owner"))
    Dim emails   : emails   = GetTexts(teamNode.SelectNodes(".//email"))
    Dim urls     : urls     = GetTexts(teamNode.SelectNodes(".//url"))
    Dim template : template = GetText(teamNode.SelectSingleNode("email_template"))
    If template = "" Then template = "<html><body><h2>Link Check Report for {team_name}</h2>{results_table}</body></html>"

    WScript.Echo vbCrLf & "=== Processing team: " & teamName & " ==="

    If UBound(urls) < 0 Then
        WScript.Echo "No URLs defined - skipping."
        Exit Sub
    End If

    Dim htmlReport
    htmlReport = GenerateReport(urls)

    Dim finalHtml : finalHtml = template
    finalHtml = Replace(finalHtml, "{team_name}", ServerHTMLEncode(teamName))
    finalHtml = Replace(finalHtml, "{process_owner}", ServerHTMLEncode(owner))
    finalHtml = Replace(finalHtml, "{results_table}", htmlReport)

    Dim subject : subject = "Link Check Report - " & teamName & " - " & FormatDateTime(Date, vbShortDate)

    If UBound(emails) >= 0 Then
        SendEmail Join(emails, ","), subject, finalHtml
    Else
        WScript.Echo "No recipient e-mail addresses found."
    End If
End Sub

'=====================================================================
Function GenerateReport(urlList)
    Dim allLinks : Set allLinks = CreateObject("Scripting.Dictionary")
    Dim i, baseUrl, linksOnPage

    For i = 0 To UBound(urlList)
        baseUrl = Trim(urlList(i))
        WScript.Echo "Scraping: " & baseUrl
        linksOnPage = GetLinksFromPage(baseUrl)
        Dim linkItem
        For Each linkItem In linksOnPage.Keys
            If Not allLinks.Exists(linkItem) Then
                allLinks(linkItem) = linksOnPage(linkItem)   ' value = Array(source, text)
            End If
        Next
        WScript.Sleep SLEEP_BETWEEN_PAGES
    Next

    WScript.Echo vbCrLf & "Testing " & allLinks.Count & " unique links..."

    Dim broken : Set broken = CreateObject("System.Collections.ArrayList")
    Dim total : total = 0
    Dim working : working = 0

    Dim url
    For Each url In allLinks.Keys
        total = total + 1
        Dim srcText : srcText = allLinks(url)
        Dim status  : status = TestUrl(url)
        If status(0) = 0 Then   ' OK
            working = working + 1
        Else
            broken.Add Array(url, srcText(0), srcText(1), status(0), status(1))
        End If
        WScript.Sleep SLEEP_BETWEEN_TESTS
    Next

    Dim summary, table
    summary = "<h2>Link Checker Report</h2>" & vbCrLf & _
              "<p style=""font-size:16px;"">" & _
              "<strong>Total unique links checked:</strong> " & total & "<br/>" & _
              "<span style=""color:green;"">Working links:</span> " & working & "<br/>" & _
              "<span style=""color:red;"">Broken or failed links:</span> " & broken.Count & "</p>"

    If broken.Count = 0 Then
        table = "<p style=""color:green;font-size:18px;font-weight:bold;"">All links are working perfectly!</p>"
    Else
        Dim rows : rows = ""
        Dim b
        For Each b In broken
            rows = rows & "<tr>" & vbCrLf & _
                   "  <td style=""max-width:250px;word-wrap:break-word;font-size:12px;""><a href=""" & b(1) & """>" & ServerHTMLEncode(b(1)) & "</a></td>" & vbCrLf & _
                   "  <td style=""max-width:200px;word-wrap:break-word;font-weight:bold;color:#d35400;"">" & ServerHTMLEncode(b(2)) & "</td>" & vbCrLf & _
                   "  <td style=""max-width:400px;word-wrap:break-word;font-size:12px;""><a href=""" & b(0) & """>" & ServerHTMLEncode(b(0)) & "</a></td>" & vbCrLf & _
                   "  <td style=""text-align:center;color:red;font-weight:bold;"">" & IIf(b(3)="N/A","Error",b(3)) & "</td>" & vbCrLf & _
                   "  <td style=""text-align:center;color:red;font-weight:bold;"">" & ServerHTMLEncode(b(4)) & "</td>" & vbCrLf & _
                   "</tr>" & vbCrLf
        Next

        table = "<h3>Broken Links Found (" & broken.Count & "):</h3>" & vbCrLf & _
                "<table border=""1"" cellpadding=""8"" cellspacing=""0"" style=""border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:13px;"">" & vbCrLf & _
                "  <thead style=""background-color:#c0392b;color:white;""><tr>" & vbCrLf & _
                "    <th width=""22%"">Found on Page</th>" & vbCrLf & _
                "    <th width=""18%"">Link Text</th>" & vbCrLf & _
                "    <th width=""40%"">Broken URL</th>" & vbCrLf & _
                "    <th width=""10%"">Status</th>" & vbCrLf & _
                "    <th width=""10%"">Result</th>" & vbCrLf & _
                "  </tr></thead><tbody>" & vbCrLf & rows & "</tbody></table>"
    End If

    GenerateReport = summary & table
End Function

'=====================================================================
Function GetLinksFromPage(pageUrl)
    Dim dict : Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim xmlhttp : Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlhttp.setTimeouts 5000, 5000, 30000, 120000
    xmlhttp.Open "GET", pageUrl, False
    xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) LinkCheckerBot/1.0"
    ' Ignore SSL errors (common in corporate environments)
    xmlhttp.setOption 2, 13056 ' SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
    xmlhttp.Send
    If Err.Number <> 0 Or xmlhttp.Status <> 200 Then
        WScript.Echo "  → Failed to load page (" & xmlhttp.Status & ")"
        Set GetLinksFromPage = dict
        Exit Function
    End If

    Dim htmlDoc : Set htmlDoc = CreateObject("HTMLFile")
    htmlDoc.write xmlhttp.responseText
    htmlDoc.close

    Dim a
    For Each a In htmlDoc.getElementsByTagName("a")
        If Len(a.href) > 0 Then
            Dim absUrl : absUrl = ResolveUrl(a.href, pageUrl)
            If Left(absUrl, 4) = "http" Then
                Dim txt : txt = Trim(a.innerText)
                If txt = "" Then txt = "(no text / image link)"
                If Len(txt) > 80 Then txt = Left(txt,77) & "..."
                If Not dict.Exists(absUrl) Then dict(absUrl) = Array(pageUrl, txt)
            End If
        End If
    Next
    WScript.Echo "  → Found " & dict.Count & " links"
    Set GetLinksFromPage = dict
    On Error Goto 0
End Function

'=====================================================================
Function TestUrl(url)
    Dim code, text
    code = 0 : text = "OK"
    On Error Resume Next
    Dim winhttp : Set winhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    winhttp.SetTimeouts 5000, 5000, 15000, 60000
    winhttp.Option(4) = 13056  ' WinHttpRequestOption_SecureProtocols - ignore cert errors
    winhttp.Open "HEAD", url, False
    winhttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) LinkCheckerBot/1.0"
    winhttp.Send
    If Err.Number = 0 And winhttp.Status < 400 Then
        code = winhttp.Status
    Else
        ' Fallback to GET
        Err.Clear
        winhttp.Open "GET", url, False
        winhttp.Send
        If Err.Number = 0 And winhttp.Status < 400 Then
            code = winhttp.Status
        Else
            code = IIf(Err.Number <> 0, "N/A", winhttp.Status)
            text = "FAILED (" & Err.Description & ")"
        End If
    End If
    On Error Goto 0
    TestUrl = Array(code, text)
End Function

'=====================================================================
Sub SendEmail(toList, subject, htmlBody)
    Dim msg : Set msg = CreateObject("CDO.Message")
    msg.From = SENDER_EMAIL
    msg.To = toList
    msg.Subject = subject
    msg.HTMLBody = htmlBody
    msg.Configuration.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    msg.Configuration.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP_SERVER
    msg.Configuration.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTP_PORT
    msg.Configuration.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
    msg.Configuration.Fields.Update
    On Error Resume Next
    msg.Send
    If Err.Number = 0 Then
        WScript.Echo "Email sent → " & toList
    Else
        WScript.Echo "SMTP ERROR: " & Err.Description
    End If
    On Error Goto 0
End Sub

'=====================================================================
Function ResolveUrl(rel, base)
    Dim obj : Set obj = CreateObject("MSXML2.XMLHTTP")
    obj.Open "GET", base, False
    obj.Send
    ResolveUrl = obj.getResponseHeader("Location")
    If ResolveUrl = "" Then ResolveUrl = base
    Set obj = CreateObject("Scripting.FileSystemObject")
    If Not obj.GetFile(ResolveUrl) Is Nothing Then  ' just a trick to get absolute URL
        ResolveUrl = CreateObject("MSXML2.XMLHTTP").getAbsoluteURI(rel, ResolveUrl)
    Else
        Dim re : Set re = New RegExp
        re.Pattern = "^https?://"
        If re.Test(rel) Then
            ResolveUrl = rel
        ElseIf Left(rel,1) = "/" Then
            ResolveUrl = Left(ResolveUrl, InStr(8, ResolveUrl, "/")) & Mid(rel,2)
        Else
            Dim p : p = InStrRev(ResolveUrl, "/")
            ResolveUrl = Left(ResolveUrl, p) & rel
        End If
    End If
End Function

'=====================================================================
Function GetText(node)
    If Not node Is Nothing Then GetText = Trim(node.text) Else GetText = ""
End Function

Function GetTexts(nodes)
    Dim a(), i
    ReDim a(-1)
    For i = 0 To nodes.length-1
        If Trim(nodes(i).text) <> "" Then
            ReDim Preserve a(UBound(a)+1)
            a(UBound(a)) = Trim(nodes(i).text)
        End If
    Next
    GetTexts = a
End Function

Function ServerHTMLEncode(s)
    ServerHTMLEncode = Replace(Replace(Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;"), vbCrLf, "<br/>")
End Function