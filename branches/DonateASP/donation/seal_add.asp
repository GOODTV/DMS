<!--#include file="../include/dbfunction.asp"-->
<%
If request("action")="save" Then
  sql="SEAL"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,3
  RS1.AddNew
  RS1("Seal_Subject")=request("Seal_Subject")
  RS1("Seal_TitleImg")=""
  If request("Seal_Flg")<>"" Then
    RS1("Seal_Flg")="Y"
  Else
    RS1("Seal_Flg")="N"
  End If
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="新增資料成功，請上傳電子印鑑 ！"
  SQL="Select @@IDENTITY as ser_no"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF then
    response.redirect "seal_edit.asp?ser_no="&RS("ser_no")
  End If
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <title>印鑑管理管理</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
  <form name="form" method="POST" action="">
    <input type="hidden" name="action">
    <input type="hidden" name="Seal_Flg" value="Y">	
    <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
      <tr>
        <td>
          <table width="780"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
            <tr>
              <td width="5%"></td>
              <td width="95%">
  		          <table width="40%"  border="0" cellspacing="0" cellpadding="0">
		              <tr>
                    <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                    <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                    <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                  </tr>
                  <tr>
                    <td class="table62-bg">&nbsp;</td>
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">新增印鑑管理</td>
                    <td class="table63-bg">&nbsp;</td>
                  </tr>  
    	          </table>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td> 
	        <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
              <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
              <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
            </tr>
            <tr>
              <td class="table62-bg">&nbsp;</td>
              <td valign="top">
                <table width="100%" border="0" style="border-collapse: collapse" bordercolor="#111111">
                  <tr>
                  	<td width="5%" valign="top"> </td> 
                    <td width="35%" valign="top">經手人姓名<span lang="en-us">:</span>&nbsp;<input type="text" name="Seal_Subject" size="28" class="font9"></td>
                    <td width="60%" valign="top"> </td> 
                  </tr>
                  <tr>
                    <td width="99%" valign="top" colspan="4" height="5"> </td>
                  </tr>
                  <tr>
                    <td width="99%" valign="top" colspan="4" bgcolor="#C0C0C0" height="1"> </td>
                  </tr>
                  <tr>
                    <td width="99%" valign="top" colspan="4" height="5"> </td>
                  </tr>  
                  <tr>
                    <td class="td02-c" valign="top" width="100%" align="center" colspan="4">
                      <input type="button" value=" 存 檔 " name="save" class="cbutton">
                      <input type="button" value=" 取 消 " name="cancel" class="cbutton">
                    </td>
                  </tr>
                </table>
              </td>
              <td class="table63-bg">&nbsp;</td>
            </tr>
            <tr>
              <td style="background-color:#EEEEE3"><img src="../images/table06_06.gif" width="10" height="10"></td>
              <td class="table64-bg"><img src="../images/table06_07.gif" width="1" height="10"></td>
              <td style="background-color:#EEEEE3"><img src="../images/table06_08.gif" width="10" height="10"></td>
            </tr>
          </table>
        </td>
      </tr>
    </table>   
  </form>
  <%message%>
</center>
</div>
</body>

</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<Script Language="VBScript">
sub save_OnClick
  <%call CheckString("Seal_Subject","經手人姓名")%>
  <%call Checklen("Seal_Subject",20,"經手人姓名")%>
  form.action.value="save"
  form.submit
end sub

sub cancel_OnClick
  location.href="seal.asp"
end sub    
</Script>