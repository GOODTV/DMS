<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  SQL="LINKED"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS.Addnew
  RS("Linked_Type")=request("Linked_Type")
  RS("Linked_Name")=request("Linked_Name")
  If request("Linked_Seq")<>"" Then  
    RS("Linked_Seq")=request("Linked_Seq")
  Else
    RS("Linked_Seq")=null
  End If
  RS.Update
  RS.Close
  Set RS=Nothing
  
  SQL1="Select @@IDENTITY As Linked_Id"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  Linked_Id=RS1("Linked_Id")
  RS1.Close
  Set RS1=Nothing
  
  session("errnumber")=1
  session("msg")=request("subject")&"新增成功，\n\n請繼續新增『 "&request("subject")&"類別 』 ！"
  response.redirect "linked.asp?linked_type="&request("linked_type")&"&subject="&request("subject")&"&linked_id="&Linked_Id&""
End If

Max_Seq=1
SQL1="Select Max_Seq=Isnull(Max(Linked_Seq),0) From LINKED Where Linked_Seq<99 "
If request("linked_type")<>"" Then SQL1=SQL1&"And linked_type='"&request("linked_type")&"' " 	
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
If Not RS1.EOF Then 
  Max_Seq=Cint(RS1("Max_Seq"))+1
Else
  SQL2="Select Max_Seq=Isnull(Max(Linked_Seq),0) From LINKED Where Linked_Seq>=99 "
  If request("linked_type")<>"" Then SQL2=SQL2&"And linked_type='"&request("linked_type")&"' " 	
  Set RS2 = Server.CreateObject("ADODB.RecordSet")
  RS2.Open SQL2,Conn,1,1
  If Not RS2.EOF Then Max_Seq=Cint(RS2("Max_Seq"))+1
  RS2.Close
  Set RS2=Nothing
End If
RS1.Close
Set RS1=Nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title><%=request("subject")%>新增</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=gray>
  <form name="form" method="post" action="" target="main">
    <input type="hidden" name="action">
    <input type="hidden" name="Linked_Type" value="<%=request("linked_type")%>">
    <input type="hidden" name="subject" value="<%=request("subject")%>">	
    <div align="center"><center>
      <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td width="100%">
            <table width="100%" border=1 cellspacing="0" cellpadding="2">
              <tr>
                <td width="20%" height=22 align="right" noWrap><font color="#000080"><%=request("subject")%>：</font></td>
                <td width="80%" height=22><input type="text" name="Linked_Name" size="25" class="font9" maxlength="50"></td>
              </tr>
              <tr>
                <td align="right"><font color="#000080">排序：</font></td>
                <td><input type="text" name="Linked_Seq" size="10" class="font9" maxlength="3" value="<%=Max_Seq%>"></td>
              </tr>                   
              <tr>
                <td width="100%" height=15 colspan="4" align="center">
                  <button id='save' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onclick='Save_OnClick()'> <img src='../images/save.gif' width='19' height='20' align='absmiddle'> 存檔</button>&nbsp;
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
function Save_OnClick(){
  <%call CheckStringJ("Linked_Name",""&request("subject")&"")%>
  <%call ChecklenJ("Linked_Name",50,""&request("subject")&"")%>
  <%call CheckStringJ("Linked_Seq","排序")%>
  <%call CheckNumberJ("Linked_Seq","排序")%>
  <%call SubmitJ("save")%>
  window.close();
}
--></script>