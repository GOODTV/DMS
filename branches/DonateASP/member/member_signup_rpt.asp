<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL1="Select Member_Act_ShortName,Member_Act_IsFood,Member_Act_IsPrice From MEMBER_ACT Where Member_Act_Id='"&request("member_act_id")&"'"
call QuerySQL(SQL1,RS1)
Member_Act_ShortName=RS1("Member_Act_ShortName")
Member_Act_IsFood=RS1("Member_Act_IsFood")
Member_Act_IsPrice=RS1("Member_Act_IsPrice")
RS1.Close
Set RS1=Nothing

If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "big5"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename="&Member_Act_ShortName&"_活動報名清單.xls"
End If 
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Member_Act_ShortName%>_活動報名清單</title>
</head>
<body class=tool>
<%
  SQL="Select Ser_No,報名日期=Member_Signup_DateTime,編號=(Case When Member_No<>'' Then Member_No Else CONVERT(nvarchar(20),DONOR.Donor_Id) End),姓名=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),手機=Cellular_Phone,電話日=Tel_Office,電話夜=Tel_Home,EMail=Email,會員否=(Case When IsMember='Y' Then 'V' Else '' End),報名狀態=Member_Signup_Status "
  If request("Member_Act_IsFood")="Y" Then SQL=SQL & ",飲食習慣=Member_Signup_Food "
  If request("Member_Act_IsPrice")="Y" Then SQL=SQL & ",報名費用=Member_Signup_Price,繳費日期=Member_Signup_Pay "
  SQL=SQL & "From MEMBER_SIGNUP Join DONOR On MEMBER_SIGNUP.Donor_Id=DONOR.Donor_Id Where Member_Act_Id='"&request("member_Act_id")&"' "
  If request("Member_Signup_Date_B")<>"" Then SQL=SQL & "And Member_Signup_Date>='"&request("Member_Signup_Date_B")&"' "
  If request("Member_Signup_Date_E")<>"" Then SQL=SQL & "And Member_Signup_Date<='"&request("Member_Signup_Date_E")&"' "
  If request("Member_Signup_Status")<>"" Then SQL=SQL & "And Member_Signup_Status='"&request("Member_Signup_Status")&"' "
  If request("KeyWords")<>"" Then SQL=SQL & "And (Donor_Name Like '%"&request("KeyWords")&"%') "
  SQL=SQL&"Order By Member_Signup_DateTime Desc,Ser_No Desc"
  ReportName=Member_Act_ShortName&"_活動報名清單"
  call ReportList (SQL,ReportName)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->