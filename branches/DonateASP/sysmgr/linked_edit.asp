<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From LINKED Where Ser_No='"&request("Ser_No")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("Linked_Name")=request("Linked_Name_New")
  If request("Linked_Seq")<>"" Then  
    RS("Linked_Seq")=request("Linked_Seq")
  Else
    RS("Linked_Seq")=null
  End If
  RS.Update
  RS.Close
  Set RS=Nothing
      
  If Cstr(request("Linked_Name_Old"))<>Cstr(request("Linked_Name_New")) Then
    'SQL="Update DONATE Set Linked_Name='"&request("Linked_Name_New")&"' Where Linked_Name='"&request("Linked_Name_Old")&"'"
    'Set RS=Conn.Execute(SQL)
  End if
  session("errnumber")=1
  session("msg")=request("subject")&"修改成功 ！"   
  response.redirect "linked.asp?linked_type="&request("linked_type")&"&subject="&request("subject")&""
End If

SQL="Select * From LINKED Where ser_no='"&request("ser_no")&"'"
Call QuerySQL(SQL,RS)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title><%=request("subject")%>修改</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=gray>
  <form name="form" method="post" action="" target="main">
    <input type="hidden" name="action">
    <input type="hidden" name="Linked_Type" value="<%=request("linked_type")%>">
    <input type="hidden" name="subject" value="<%=request("subject")%>">	
    <input type="hidden" name="Ser_No" value="<%=RS("Ser_No")%>">	
    <input type="hidden" name="Linked_Name_Old" value="<%=RS("Linked_Name")%>">	
    <div align="center"><center>
      <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td width="100%">
            <table width="100%" border=1 cellspacing="0" cellpadding="2">
              <tr>
                <td width="20%" height=22 align="right" noWrap><font color="#000080"><%=request("subject")%>：</font></td>
                <td width="80%" height=22><input type="text" name="Linked_Name_New" size="25" class="font9" maxlength="50" value="<%=RS("Linked_Name")%>" ></td>
              </tr>
              <tr>
                <td align="right"><font color="#000080">排序：</font></td>
                <td><input type="text" name="Linked_Seq" size="10" class="font9" maxlength="3" value="<%=RS("Linked_Seq")%>"></td>
              </tr>
              <tr>
                <td width="100%" height=15 colspan="4" align="center">
                  <button id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onclick='Update_OnClick()'> <img src='../images/update.gif' width='19' height='20' align='absmiddle'> 修改</button>&nbsp;
                  <button id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onclick='window.close();'> <img src='../images/icon6.gif' width='20' height='15' align='absmiddle'> 離開</button>	
                </td>
              </tr>   
            </table>
          </td>
        </tr>
      </table>
    </center></div>
  </form>
  <%Message()%> 
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Update_OnClick(){
  <%call CheckStringJ("Linked_Name_New",""&request("subject")&"")%>
  <%call ChecklenJ("Linked_Name_New",50,""&request("subject")&"")%>
  <%call CheckStringJ("Linked_Seq","排序")%>
  <%call CheckNumberJ("Linked_Seq","排序")%>
  <%call SubmitJ("update")%>
  window.close();
}
--></script>