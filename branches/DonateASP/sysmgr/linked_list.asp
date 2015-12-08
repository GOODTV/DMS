<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function LinkedList (SQL,Linked_Add,Linked2_Add,LinkParam,LinkTarget,Linked_Edit,Linked_PK,subject)
  response.write "<a href='#' style=""text-decoration: none"" onclick=""window.open('"&Linked_Add&"','','top=100,left=120,width=450,height=275')""><button style=""cursor:hand;"" class=""cbutton"">新增"&subject&"</button></a>"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id=grid border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
    Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  If Session("user_id")="npois" Then Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>編輯</span></font></td>"
  Response.Write "</tr>"
  While Not RS1.EOF
    If Linked2_Add<>"" Then
      Response.Write "<a href='" & Linked2_Add & RS1(LinkParam) & "' target='" & LinkTarget &"'>"   
      showhand="style='cursor:hand'"
    End if
    Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
    For I = 1 To FieldsCount
      Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體;'>" & RS1(I) & "</span></td>"
    Next
    Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & "<a href='#' onclick=""window.open('" & Linked_Edit & RS1(Linked_PK) & "','','top=100,left=120,width=450,height=275')"">修改</a></span></td>" 
    RS1.MoveNext
    Response.Write "</tr>"
    Response.Write "</a>"
  Wend
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title><%=request("subject")%>連動選單維護</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=tool>
<%
  SQL="Select Ser_No,"&request("subject")&"=Linked_Name,排序=Linked_Seq From LINKED Where Linked_Type='"&request("linked_type")&"' Order By Linked_Seq"
  Linked_Add="linked_add.asp?linked_type="&request("linked_type")&"&subject="&request("subject")
  Linked2_Add="linked2_add.asp?linked_type="&request("linked_type")&"&subject="&request("subject")&"&linked_id="
  LinkParam="ser_no"
  LinkTarget="detail"  	
  Linked_Edit="linked_edit.asp?linked_type="&request("linked_type")&"&subject="&request("subject")&"&ser_no="
  Linked_PK="ser_no"
  call LinkedList (SQL,Linked_Add,Linked2_Add,LinkParam,LinkTarget,Linked_Edit,Linked_PK,request("subject"))
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->