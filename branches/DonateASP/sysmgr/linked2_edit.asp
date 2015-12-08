<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From LINKED2 Where ser_no='"&request("ser_no")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("Linked2_Name")=request("Linked2_Name_New")
  If request("Linked2_Seq")<>"" Then  
    RS("Linked2_Seq")=request("Linked2_Seq")
  Else
    RS("Linked2_Seq")=null
  End If
  RS.Update
  RS.Close
  Set RS=Nothing
  
  If Cstr(request("Linked2_Name_Old"))<>Cstr(request("Linked2_Name_New")) Then
    'SQL="Update DONATE Set Accounting_Title='"&request("Donate_Accounting_New")&"' Where Donate_Purpose='"&request("Donate_Purpose")&"' And Accounting_Title='"&request("Donate_Accounting_Old")&"'"
    'Set RS=Conn.Execute(SQL)
  End if
  
  session("errnumber")=1
  session("msg")="物品類別修改成功 ！"
  response.redirect "linked.asp?linked_type="&request("Linked_Type")&"&subject="&request("subject")&"&linked_id="&request("linked_id")&""
End If

SQL="Select * From LINKED2 Where Ser_No="&request("ser_no")&""
Call QuerySQL(SQL,RS)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>物品類別修改</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=gray>
  <form name="form" method="post" action="" target="main">
    <input type="hidden" name="action">
    <input type="hidden" name="Ser_No" value="<%=RS("Ser_No")%>">
    <input type="hidden" name="Linked_Id" value="<%=RS("Linked_Id")%>">		
    <input type="hidden" name="Linked_Type" value="<%=request("linked_type")%>">
    <input type="hidden" name="subject" value="<%=request("subject")%>">	
    <input type="hidden" name="Linked2_Name_Old" value="<%=RS("Linked2_Name")%>">
    <div align="center"><center>
      <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td width="100%">
            <table width="100%" border=1 cellspacing="0" cellpadding="2">
              <tr>
                <td width="25%" align="right"><font color="#000080"><%=request("subject")%>類別<span lang="en-us">：</span></font></td>
                <td width="75%"><input type="text" name="Linked2_Name_New" size="30" value="<%=RS("Linked2_Name")%>" class="font9"></td>                
              </tr>         
              <tr> 
                <td align="right"><font color="#000080">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;排序<span lang="en-us">：</span></font></td>
                <td><input type="text" name="Linked2_Seq" size="8" class="font9" maxlength="3" value="<%=RS("Linked2_Seq")%>"></td>
              </tr>              
              <tr>
                <td width="100%" height=15 align="center" colspan="2">
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
  <%call CheckStringJ("Linked2_Name_New",""&request("subject")&"類別")%>
  <%call ChecklenJ("Linked2_Name_New",50,""&request("subject")&"類別")%>
  <%call CheckStringJ("Linked2_Seq","排序")%>
  <%call CheckNumberJ("Linked2_Seq","排序")%>
  <%call SubmitJ("update")%>
  window.close();
}
--></script>