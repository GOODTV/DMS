<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function CaseCodeList (SQL,HLink,LinkParam,LinkTarget,DelLink,DelParam,DelTarget)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
    Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>刪除</span></font></td>"
  Response.Write "</tr>"
  While Not RS1.EOF
    Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
    For I = 1 To FieldsCount
      If I = 1 Then
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & "<a href='#' onclick=""window.open('" & HLink & RS1(LinkParam) & "','','top=100,left=120,width=450,height=260')"">" & RS1(I) & "</a></span></td>"
      Else
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I) & "</span></td>"
      End If
    Next
    Response.Write "<td><a href='JavaScript:if(confirm(""是否確定要刪除 ?"")){window.location.href=""" & DelLink & RS1(DelParam) & """;}' target='" & DelTarget &"'><img src='../images/x5.gif' border=0 width='16' height='14' alt='刪除'></a></td>"    
    RS1.MoveNext
    Response.Write "</tr>"
  Wend
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function

If request("action")="save" Then
  SQL="LINKED2"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,3
  RS1.AddNew
  RS1("Linked_Id")=request("Linked_Id")
  RS1("Linked2_Name")=request("Linked2_Name")
  If request("Linked2_Seq")<>"" Then  
    RS1("Linked2_Seq")=request("Linked2_Seq")
  Else
    RS1("Linked2_Seq")=null
  End If
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")=request("subject")&"次分類新增成功 ！"
End If

If request("action")="delete" Then
  SQL="Select * From LINKED2 Where Linked_Id='"&request("Linked_Id")&"'"
  Call QuerySQL(SQL,RS)
  If Not RS.EOF Then
    session("errnumber")=1
    session("msg")="『 "&request("Linked_Name")&" 』還有次分類，不能刪除！"
    response.redirect "linked.asp?linked_type="&request("linked_type")&"&subject="&request("subject")&"&linked_id="&request("linked_id")&""
  Else
    SQL="Delete From LINKED Where Ser_No='"&request("Linked_Id")&"'"
    Set RS=Conn.Execute(SQL)
    session("errnumber")=1
    session("msg")=request("subject")&"刪除成功 ！"
    response.redirect "linked.asp?linked_type="&request("linked_type")&"&subject="&request("subject")&""
  End If
End If

Max_Seq=1
SQL1="Select Max_Seq=Isnull(Max(Linked2_Seq),0) From LINKED2 Where Linked_Id='"&request("Linked_Id")&"' And Linked2_Seq<99"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then 
  Max_Seq=Cint(RS1("Max_Seq"))+1
Else
  SQL2="Select Max_Seq=Isnull(Max(Linked2_Seq),0) From LINKED2 Where Linked_Id='"&request("Linked_Id")&"' And Linked2_Seq>=99"
  Call QuerySQL(SQL2,RS2)
  If Not RS2.EOF Then Max_Seq=Cint(RS2("Max_Seq"))+1
  RS2.Close
  Set RS2=Nothing 
End If
RS1.Close
Set RS1=Nothing 

SQL1="Select * From LINKED Where Ser_No='"&request("linked_id")&"'"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then Linked_Name=RS1("Linked_Name")
RS1.Close
Set RS1=Nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title><%=request("subject")%>列表</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=tool>
  <form name="form" method="post" action="" target="main">
    <input type="hidden" name="action">
    <input type="hidden" name="linked_type" value="<%=request("linked_type")%>">
    <input type="hidden" name="subject" value="<%=request("subject")%>">	
    <input type="hidden" name="Linked_Id" value="<%=request("linked_id")%>">
    <input type="hidden" name="Linked_Name" value="<%=Linked_Name%>">			
    <table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%" align="center">
      <tr>
        <td width="23%" align="right"><%=request("subject")%><span lang="en-us">：</span></td>
        <td width="38%"><input type="text" name="Donate_Purpose" size="21" value="<%=Linked_Name%>" readonly class="font9t"> </td>
        <td width="39%"><input type="button" value="刪除<%=request("subject")%>" name="delete" class="delbutton" style="cursor:hand;" onClick="Delete_OnClick()"></td>
      </tr>
      <tr>
        <td width="100%" colspan="6" bgcolor="#C0C0C0" height="1"> </td>
      </tr>
      <tr>
        <td align="right">次分類名稱<span lang="en-us">：</span></td>
        <td><input type="text" name="Linked2_Name" size="21" class="font9" maxlength="50"></td>
        <td><input type="button" value="新增次分類" name="Save" class="addbutton" style="cursor:hand;" onClick="Save_OnClick()"></td>
      </tr>
      <tr>
        <td align="right">排序<span lang="en-us">：</span></td>
        <td><input type="text" name="Linked2_Seq" size="8" class="font9" maxlength="3" value="<%=Max_Seq%>"></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td width="100%" colspan="5">
        <%
          SQL="Select Ser_No,次分類=Linked2_Name,排序=Linked2_Seq From LINKED2 Where Linked_Id='"&request("linked_id")&"' Order By Linked2_Seq"
          HLink="linked2_edit.asp?linked_type="&request("linked_type")&"&subject="&request("subject")&"&linked_id="&request("linked_id")&"&ser_no="
          LinkParam="ser_no"
          LinkTarget="update"
             
          DelLink="linked2_delete.asp?linked_type="&request("linked_type")&"&subject="&request("subject")&"&linked_id="&request("linked_id")&"&ser_no="
          DelParam="ser_no"
          DelTarget="detail"

          call CaseCodeList (SQL,HLink,LinkParam,LinkTarget,DelLink,DelParam,DelTarget)
        %>
        </td>
      </tr>
  </form>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Save_OnClick(){
  <%call CheckStringJ("Linked2_Name","次分類名稱")%>
  <%call ChecklenJ("Linked2_Name",50,"次分類名稱")%>
  <%call CheckStringJ("Linked2_Seq","排序")%>
  <%call CheckNumberJ("Linked2_Seq","排序")%>
  document.form.target='detail';
  <%call SubmitJ("save")%>
}
function Delete_OnClick(){
  if(confirm('您是否確定要刪除'+document.form.subject.value+'？')){
    document.form.action.value='delete';
    document.form.submit();
  }
}
--></script>