<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function CodeFileList (SQL,HLink,LinkParam,LinkTarget,CodeLink,CodeParam)
  If Session("user_id")="npois" Then response.write "<a href style=""text-decoration: none"" onclick=""window.open('codefile_add.asp','','top=100,left=120,width=450,height=275')""><button style=""cursor:hand;"" class=""cbutton"">新增主檔</button></a>"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id=grid border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount-1
    Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  If Session("user_id")="npois" Then Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>編輯</span></font></td>"
  Response.Write "</tr>"
  While Not RS1.EOF
    If Hlink<>"" Then
      Response.Write "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>"   
      showhand="style='cursor:hand'"
    End if
    Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
    For I = 1 To FieldsCount-1
      Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體;'>" & RS1(I) & "</span></td>"
    Next
    If Session("user_id")="npois" Then Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & "<a href='#' onclick=""window.open('" & CodeLink & RS1(CodeParam) & "','','top=100,left=120,width=450,height=275')"">修改</a></span></td>" 
    RS1.MoveNext
    Response.Write "</tr>"
    Response.Write "</a>"
  Wend
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function

SQL_IS="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"' "
Set RS_IS=Server.CreateObject("ADODB.Recordset")
RS_IS.Open SQL_IS,Conn,1,1
IsEpaper=RS_IS("IsEpaper")
IsDonation=RS_IS("IsDonation")
IsContribute=RS_IS("IsContribute")
IsMember=RS_IS("IsMember")
IsShopping=RS_IS("IsShopping")
RS_IS.Close
Set RS_IS=Nothing
Category2="'sys'"
If IsEpaper="Y" Then Category2=Category2&",'epaper'"
If IsDonation="Y" Then Category2=Category2&",'donation'"
If IsContribute="Y" Then Category2=Category2&",'contribute'"
If IsMember="Y" Then Category2=Category2&",'member','signup'"
If IsShopping="Y" Then Category2=Category2&",'shopping'"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>代碼主檔列表</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=tool>
<%
  If Session("user_id")="npois" Then 
    SQL="Select CodeType,代碼=CodeType,代碼名稱=CodeName,類別=Category,排序=Menu_Seq,'1' As Seq From CODEFILE Where Category<>'類別選項' " & _
        "Union " & _
        "Select CodeType,代碼=CodeType,代碼名稱=CodeName,類別=Category,排序=Menu_Seq,'2' As Seq From CODEFILE Where Category='類別選項' And Category2 In ("&Category2&") " & _
        "Order By Seq,Menu_Seq,Category,CodeType"
  Else
    SQL="Select CodeType,代碼=CodeType,代碼名稱=CodeName,類別=Category From CODEFILE Where Category='類別選項'  And IsShow='Y' And Category2 In ("&Category2&") Order By CodeType"
  End If
  HLink="casecode_add.asp?codetype="
  LinkParam="codetype"
  LinkTarget="detail"  	
  CodeLink="codefile_edit.asp?codetype="
  CodeParam="codetype"
  call CodeFileList (SQL,HLink,LinkParam,LinkTarget,CodeLink,CodeParam)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->