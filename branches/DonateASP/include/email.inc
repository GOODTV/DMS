<%
Sub SendEmail (FromName,FromEmail,ToName,ToEmail,cc,Subject,MailType,Content,AttachFile)
    Set Mail = Server.CreateObject("CDONTS.NewMail")
    Mail.from = FromName & "<" & FromEmail & ">"
    Mail.To = ToName & "<" & ToEmail & ">"
    if cc <> "" then
       Mail.cc = cc
    end if
    Mail.Subject = Subject
    if MailType="html" then
	Mail.BodyFormat = 0
	Mail.MailFormat = 0
    end if
    Mail.Body = Content
    if AttachFile <> "" then
       Mail.AttachFile Server.MapPath(AttachFile)
    end if
    Mail.Send
    Set Mail=nothing
End sub

Sub SendMessage (FromName,FromEmail,ToName,ToEmail,cc,Subject,MailType,Content,AttachFile)
    Set Mail=CreateObject("CDO.Message")
    Mail.Subject=Subject
    Mail.From=FromName & "<" & FromEmail & ">"
    Mail.To=ToName & "<" & ToEmail & ">"
    if cc <> "" then
       Mail.cc = cc
    end if
    if MailType="html" then
       Mail.HTMLBody=Content
    elseif MailType="WebPage" then
       Mail.CreateMHTMLBody Content
    else
       Mail.TextBody=Content
    end if
    if AttachFile <> "" then
       Mail.AddAttachment Server.MapPath(AttachFile)
    end if
    Mail.Send
    set Mail=nothing
end sub

Sub SendJMail (FromName,FromEmail,ToName,ToEmail,cc,Subject,MailType,Content,AttachFile)
    Set JMail = Server.CreateObject("JMail.Message")
    JMail.Charset = "utf-8"
    JMail.From = FromEmail
    JMail.FromName = FromName
    JMail.MailServerUserName = "just@sino1.com.tw"
    JMail.MailServerPassWord = "Abc123"
    JMail.Subject = Subject
    JMail.AddRecipient ToEmail,ToName
    if cc <> "" then
       JMail.AddRecipientCC cc
    end if
    if MailType="html" then
       JMail.HTMLBody= Content
    else 
       JMail.Body = Content
    end if
    if AttachFile<>"" then
       contentId = JMail.AddAttachment(Server.MapPath(AttachFile))
    end if   
    JMail.Priority = 3
    JMail.Send("192.168.10.3")
    Set JMail = Nothing
End sub

Sub DoEmail (FromName,FromEmail,RecipientEmail,MailTitle,MailBody,MailType)
    Set JMail = Server.CreateObject("JMail.Message")    
    JMail.Charset = "utf-8"
    JMail.From = FromEmail
    JMail.FromName = FromName
    JMail.Subject = MailTitle
    JMail.AddRecipient RecipientEmail
    if MailType="html" then
      JMail.HTMLBody= MailBody
    else
      JMail.Body = MailBody
    end if
    JMail.Body = MailBody
    JMail.Priority = 3
    on error resume next
    JMail.Send(session("JMail"))
    Set JMail = Nothing
End Sub

Sub SendEmailtoRaw (FromEmail,ToEmail,Subject,mc,Content)
    Set FS = Server.CreateObject("Scripting.FileSystemObject")
    Set F = FS.OpenTextFile("D:\Inetpub\Mailroot\pickup\"&mc&".RAW",2,True)
    F.WriteLine("x-sender: "&FromEmail )
    F.WriteLine("x-receiver: "&ToEmail )
    F.WriteLine("From: "&FromEmail)
    F.WriteLine("To: "&ToEmail )
    F.WriteLine("Subject: "&Subject )
    F.WriteLine("Content-Type: multipart/related; boundary=""==PageRelatedBoundary"";type=""text/html""")
    F.WriteLine("")
    F.WriteLine("This is a multi-part message in MIME format.")
    F.WriteLine("")
    F.WriteLine("--==PageRelatedBoundary")
    F.WriteLine("Content-Type: text/html; charset=""utf-8""")
    F.WriteLine("")
    F.Write(Content)
    F.WriteLine("")
    F.Close
    Set F = Nothing
End Sub

Sub SendEmailtoMDaemon (FromName,FromEmail,ToName,ToEmail,Subject,mc,Content)
    Set FS = Server.CreateObject("Scripting.FileSystemObject")
    Set F = FS.OpenTextFile("D:\MDaemon\Queues\Raw\"&mc&".RAW",2,True)
    F.WriteLine("From <"&FromName&" <"&FromEmail&">>")
    F.WriteLine("To <"&ToName&" <"&ToEmail&">>")
    F.WriteLine("Subject < "&Subject&" >")
    F.WriteLine("Header <Content-Type: multipart/related; boundary=""==PageRelatedBoundary""; type=""text/html"">")
    F.WriteLine("")
    F.WriteLine("This is a multi-part message in MIME format.")
    F.WriteLine("")
    F.WriteLine("--==PageRelatedBoundary")
    F.WriteLine("Content-Type: text/html; charset=""utf-8""")
    F.WriteLine("")
    F.Write(Content)
    F.WriteLine("")
    F.Close
    Set F = Nothing
End Sub

Sub SendEmailtoMDaemon1 (FromName,FromEmail,ToName,ToEmail,Subject,mc,Content)
    Set FS = Server.CreateObject("Scripting.FileSystemObject")
    Set F = FS.OpenTextFile("D:\MDaemon\Queues\Raw1\"&mc&".RAW",2,True)
    F.WriteLine("From <"&FromName&" <"&FromEmail&">>")
    F.WriteLine("To <"&ToName&" <"&ToEmail&">>")
    F.WriteLine("Subject < "&Subject&" >")
    F.WriteLine("Header <Content-Type: multipart/related; boundary=""==PageRelatedBoundary""; type=""text/html"">")
    F.WriteLine("")
    F.WriteLine("This is a multi-part message in MIME format.")
    F.WriteLine("")
    F.WriteLine("--==PageRelatedBoundary")
    F.WriteLine("Content-Type: text/html; charset=""utf-8""")
    F.WriteLine("")
    F.Write(Content)
    F.WriteLine("")
    F.Close
    Set F = Nothing
End Sub

Sub SendEmailtoMDaemon2 (FromName,FromEmail,ToName,ToEmail,Subject,mc,Content)
    Set FS = Server.CreateObject("Scripting.FileSystemObject")
    Set F = FS.OpenTextFile("D:\MDaemon\Queues\Raw2\"&mc&".RAW",2,True)
    F.WriteLine("From <"&FromName&" <"&FromEmail&">>")
    F.WriteLine("To <"&ToName&" <"&ToEmail&">>")
    F.WriteLine("Subject < "&Subject&" >")
    F.WriteLine("Header <Content-Type: multipart/related; boundary=""==PageRelatedBoundary""; type=""text/html"">")
    F.WriteLine("")
    F.WriteLine("This is a multi-part message in MIME format.")
    F.WriteLine("")
    F.WriteLine("--==PageRelatedBoundary")
    F.WriteLine("Content-Type: text/html; charset=""utf-8""")
    F.WriteLine("")
    F.Write(Content)
    F.WriteLine("")
    F.Close
    Set F = Nothing
End Sub

Function WinHttp (WebUrl,HtmType)
  Dim whttp
  Set wHttp=Server.CreateObject("WinHttp.WinHttpRequest.5.1")
  If HtmType="POST" Then
    wHttp.Open "POST",WebUrl
  Else
    wHttp.Open "GET",WebUrl, false
  End If
  wHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  wHttp.Send
  WinHttp=wHttp.ResponseText
End Function

Function WinHttp2(MailLink,Charset)
  Dim wHttp, oStream
  Set wHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1") 
  wHttp.Open "POST", MailLink
  wHttp.SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
  wHttp.Send

  Set oStream = Server.CreateObject("adodb.stream")
  oStream.Type=1
  oStream.Mode=3
  oStream.Open()
  oStream.Write(wHttp.ResponseBody)
  oStream.Position=0
  oStream.Type=2
  oStream.Charset=Charset
  WinHttp2= oStream.ReadText
  Set oStream = Nothing
  Set WinHttpReq = Nothing
End Function
  
Function Send_EMail(SendType,FromName,FromEmail,ToName,ToEmail,cc,Subject,MailType,Content,AttachFile,mc)
    Response.Write "SendType:<br />"&SendType
    Response.End
  If SendType="Email" Then
    call SendEmail (FromName,FromEmail,ToName,ToEmail,cc,Subject,MailType,Content,AttachFile)
  ElseIf SendType="Message" Then
    call SendMessage (FromName,FromEmail,ToName,ToEmail,cc,Subject,MailType,Content,AttachFile)
  ElseIf SendType="JMail" Then
    call SendJMail (FromName,FromEmail,ToName,ToEmail,cc,Subject,MailType,Content,AttachFile)
  ElseIf SendType="DoEmail" Then
    call DoEmail (FromName,FromEmail,ToName,Subject,Content,MailType)
  ElseIf SendType="Raw" Then
    call SendEmailtoRaw (FromEmail,ToEmail,Subject,mc,Content)
  ElseIf SendType="MDaemon" Then
    If Cstr(Request.ServerVariables("SERVER_NAME"))="localhost" Or Cstr(Request.ServerVariables("SERVER_NAME"))="127.0.0.1" Then 
      MDName="local"
    ElseIf Cstr(Request.ServerVariables("SERVER_NAME"))="web.npois.com.tw" Then
      MDName=Split(Request.ServerVariables("URL"),"/")(2)
    Else
      Ary_MDName=Split(Request.ServerVariables("SERVER_NAME"),".")
      If LCase(Ary_MDName(0))="www" Then
        MDName=Ary_MDName(1)
      Else
        MDName=Ary_MDName(0)
      End If
    End If
    send_mc=MDName&Cstr(Year(now()))&Cstr(Month(now()))&Cstr(day(now()))&Cstr(Hour(now()))&Cstr(Minute(now()))&Cstr(Second(now()))&Cstr(mc)
    'Response.Write "Content:<br />"&Content
    'Response.End
    call SendEmailtoMDaemon (FromName,FromEmail,ToName,ToEmail,Subject,send_mc,Content)
  ElseIf SendType="MDaemon1" Then
    If Cstr(Request.ServerVariables("SERVER_NAME"))="localhost" Or Cstr(Request.ServerVariables("SERVER_NAME"))="127.0.0.1" Then 
      MDName="local"
    ElseIf Cstr(Request.ServerVariables("SERVER_NAME"))="web.npois.com.tw" Then
      MDName=Split(Request.ServerVariables("URL"),"/")(2) 
    Else
      Ary_MDName=Split(Request.ServerVariables("SERVER_NAME"),".")
      If LCase(Ary_MDName(0))="www" Then
        MDName=Ary_MDName(1)
      Else
        MDName=Ary_MDName(0)
      End If
    End If
    send_mc=MDName&Cstr(Year(now()))&Cstr(Month(now()))&Cstr(day(now()))&Cstr(Hour(now()))&Cstr(Minute(now()))&Cstr(Second(now()))&Cstr(mc)
    call SendEmailtoMDaemon1 (FromName,FromEmail,ToName,ToEmail,Subject,send_mc,Content)
  ElseIf SendType="MDaemon2" Then
    If Cstr(Request.ServerVariables("SERVER_NAME"))="localhost" Or Cstr(Request.ServerVariables("SERVER_NAME"))="127.0.0.1" Then 
      MDName="local"
    ElseIf Cstr(Request.ServerVariables("SERVER_NAME"))="web.npois.com.tw" Then
      MDName=Split(Request.ServerVariables("URL"),"/")(2) 
    Else
      Ary_MDName=Split(Request.ServerVariables("SERVER_NAME"),".")
      If LCase(Ary_MDName(0))="www" Then
        MDName=Ary_MDName(1)
      Else
        MDName=Ary_MDName(0)
      End If
    End If
    send_mc=MDName&Cstr(Year(now()))&Cstr(Month(now()))&Cstr(day(now()))&Cstr(Hour(now()))&Cstr(Minute(now()))&Cstr(Second(now()))&Cstr(mc)
    call SendEmailtoMDaemon2 (FromName,FromEmail,ToName,ToEmail,Subject,send_mc,Content)     
  End If
End Function
%>