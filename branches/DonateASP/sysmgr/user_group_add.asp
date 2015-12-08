<!--#include file="../include/dbfunction.asp"-->
<%
if request("action")="save" then
   sql="select * from groupfile where user_group='"&request("usergroup")&"'"
   Set RS1 = Server.CreateObject("ADODB.RecordSet")
   RS1.Open SQL,Conn,1,1
   if Not RS1.EOF then
      session("errnumber")=1
      session("msg")="資料重覆 ! "&request("usergroup")&" 帳號類別資料已存在"
   else
      sql="insert into groupfile values ('"&request("usergroup")&"','"&request("security")&"')"
      call ExecSQL (SQL)
      Response.Redirect "user_group_edit.asp?usergroup="&request("usergroup")
   end if
end if   
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>帳號權限設定--資料輸入</title>
</head>
<BODY>
<p>
<table border="0" width="80%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"  height="25">
        帳號權限設定【新增資料】</td>
  </tr>
  <tr>
    <td width="100%">
                  <form name="form" action="" method="post">
                  <input type="hidden" name="action"> 
                    <div align="center">
                      <center>
                    <table width="100%" border=1 cellspacing="1" style="border-collapse: collapse" bordercolor="#6487DC" cellpadding="2">
                      <tr> 
                        <td noWrap align=right width="19%" bgcolor=#FFFFCC height=22>
                        <font color="#FF0000">*</font> <font color="#000080">
                        <span lang="zh-tw">帳號類別</span>：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="10" --><input class=font9 size=20 
                        name="usergroup" maxlength="10"></td>
                      </tr>
                      <tr> 
                        <td align=right width="19%" bgcolor=#FFFFCC height=22 class="td1">
                        安全控管：</td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1">
                          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="1" --><input type="text" name="security" size="4" class="font9" maxlength="1"></td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 height=22 class="font9" colspan="2">
                         <%SaveButton%>　
                        </td>
                        </table>
                        <%message%>           
                      </center>
                    </div>
                  </form>
      </td>
  </tr>
</table>
</BODY>
</HTML>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--

sub save_OnClick
  <%call CheckString("usergroup","帳號類別")%>
  <%call Checklen("usergroup",20,"帳號類別")%>
  form.action.value="save"
  form.submit   
end sub

sub cancel_OnClick
  history.back
end sub

sub query_OnClick
  location.href="user_group.asp"
end sub  

--></script>