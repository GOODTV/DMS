<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From NEXTPAGE Where page_id='"&request("page_id")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("Page_Desc")=request("Page_Desc")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="分頁內容修改成功 ！"
  response.redirect request("item")&"_edit.asp?code_id="&request("code_id")&"&ser_no="&request("ser_no")
End If

If request("action")="delete" Then
  SQL="Delete From NEXTPAGE Where page_id='"&request("page_id")&"'"
  Set RS=conn.Execute(SQL)
  SQL1="Select * From NEXTPAGE Where Ser_No='"&request("ser_no")&"' And Page_Type='"&request("item")&"'"
  Call QuerySQL(SQL1,RS1)
  If RS1.EOF Then
    SQL="Update "&request("item")&" Set "&request("item")&"_IsPage='N' Where ser_no='"&request("ser_no")&"'"
    Set RS=conn.Execute(SQL) 
  End If
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="分頁內容已刪除 ！"
  response.redirect request("item")&"_edit.asp?code_id="&request("code_id")&"&ser_no="&request("ser_no")
End If   

SQL="Select * From NEXTPAGE Where page_id='"&request("page_id")&"'"
Call QuerySQL(SQL,RS)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>分頁內容修改</title>
</head>
<body class=gray>
  <form name="form" method="post" action="" target="main">
    <input type="hidden" name="action">
    <input type="hidden" name="item" value="<%=request("item")%>">  
    <input type="hidden" name="page_id" value="<%=request("page_id")%>"> 
    <input type="hidden" name="ser_no" value="<%=RS("ser_no")%>">
    <input type="hidden" name="code_id" value="<%=request("code_id")%>"> 	 
    <table width="100%" border="0" cellspacing="1" cellpadding="2" style="border-collapse: collapse" bordercolor="#E1DFCE">
      <tr>
	      <td><font color="#000080">&nbsp;&nbsp;第&nbsp;<%=rs("page_no")%>&nbsp;頁</font></td>
      </tr>  
      <tr>
        <td>
	      <!--#include file="../fckeditor/fckeditor.asp" -->
        <%
          Dim oFCKeditor
	        Set oFCKeditor = New FCKeditor
	        oFCKeditor.BasePath	= "../fckeditor/"
	        oFCKeditor.Value	= RS("Page_Desc")
	        oFCKeditor.Create "Page_Desc"
        %> 
	      </td>
      </tr>
      <tr>
        <td><%EditButton%></td>
      </tr>
    </table>
  </form>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Update_OnClick(){
  <%call SubmitJ("update")%>
  window.close();
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
  window.close();
}
function Cancel_OnClick(){
  window.close();
}
--></script>