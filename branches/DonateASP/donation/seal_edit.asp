<!--#include file="../include/dbfunction.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From SEAL Where Ser_No='"&request("Ser_No")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,3
  RS1("Seal_Subject")=request("Seal_Subject")
  If request("Seal_Flg")<>"" Then
    RS1("Seal_Flg")="Y"
  Else
    RS1("Seal_Flg")="N"
  End If
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="修改資料成功 !"
  response.redirect "seal.asp"
End If

if request("action")="delete" then
   sql="Delete From SEAL where Ser_No='"&request("Ser_No")&"'"
   Call QuerySQL(SQL,RS1)
   session("errnumber")=1
   session("msg")="刪除資料成功 !"
   response.redirect "seal.asp"   
End If

SQL="Select * From SEAL Where Ser_No='"&request("Ser_No")&"'"
Call QuerySQL(SQL,RS)
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
    <input type="hidden" name="Ser_No" value="<%=rs("Ser_No")%>">	
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
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">修改印鑑管理</td>
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
                  	<td width="8%" valign="top"> </td> 
                    <%If Cint(RS("Ser_No"))<=4 Then%>
                     <td width="12%" valign="top" align="right">印鑑標題<span lang="en-us">:</span></td>
                    <td width="15%" valign="top"><input type="text" name="Seal_Subject" size="28" class="font9t" readonly value="<%=RS("Seal_Subject")%>"></td>	
                    <%Else%>
                     <td width="12%" valign="top" align="right">經手人姓名<span lang="en-us">:</span></td>
                    <td width="15%" valign="top"><input type="text" name="Seal_Subject" size="28" class="font9" value="<%=RS("Seal_Subject")%>"></td>	
                    <%End If%>
                    <td width="55%" valign="top"> </td> 
                  </tr>
                  <tr>
                  	<td width="10%" valign="top"> </td> 
                    <td valign="top" align="right">電子印鑑圖檔<span lang="en-us">:</span></td>
                    <td colspan="3" align="left">
                      <%if rs("Seal_TitleImg")<>"" then%>
                      <input type="button" value="重新上傳" name="upload" class="cbutton"><br /><br />
                      <img border="0" src="../upload/<%=rs("Seal_TitleImg")%>" align="absmiddle" height="40">
                      <%Else%>
                      <input type="button" value="上傳" name="upload" class="cbutton"><br /><br />
                      <%end if%>
                    </td>
                  </tr>
                  <tr>
                    <td width="99%" valign="top" colspan="5" height="5"> </td>
                  </tr>
                  <tr>
                    <td width="99%" valign="top" colspan="5" bgcolor="#C0C0C0" height="1"> </td>
                  </tr>
                  <tr>
                    <td width="99%" valign="top" colspan="5" height="5"> </td>
                  </tr> 
                  <tr>
                    <td class="td02-c" valign="top" width="100%" align="center" colspan="5">
                      <%If Cint(rs("Ser_No"))>4 Then%>
                        <input type="button" value="修改存檔" name="update" class="cbutton">
                        <input type="button" value="刪除資料" name="delete" class="cbutton">
                      <%End If%>
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
<Script Language="VBScript">
sub update_OnClick
  <%call CheckString("Seal_Subject","印鑑標題")%>
  <%call Checklen("Seal_Subject",20,"印鑑標題")%>
  call ConfirmUpdate()
end sub

sub del_OnClick
   DIM answer
   answer = window.confirm ( "您是否確定要刪除 ?")
   if answer = true then
      form.action.value="delete"
      form.submit
   end if
end sub

sub cancel_OnClick
  location.href="seal.asp"
end sub

sub query_OnClick
  location.href="seal.asp"
end sub

sub upload_OnClick
  open "seal_upload.asp?ser_no=<%=request("Ser_No")%>&Seal_Subject=<%=RS("Seal_Subject")%>","upload","scrollbars=no,status=no,toolbar=no,top=100,left=120,width=550,height=200"
end sub    
</Script>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->