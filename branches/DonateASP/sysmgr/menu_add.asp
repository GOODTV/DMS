<!--#include file="../include/dbfunction.asp"-->
<%
if request("action")="save" then
   sql="select * from menu where menu_id='"&request("menu_id")&"'"
   Set RS1 = Server.CreateObject("ADODB.RecordSet")
   RS1.Open SQL,Conn,1,1
   if Not RS1.EOF then
      session("errnumber")=1
      session("msg")="資料重覆 ! "&request("menu_id")&" 程式群組資料已存在"
   else
      sql="insert into menu values ('"&request("menu_id")&"','"&request("menu_seq")&"','"&request("menu_url")&"',0,'"&request("security")&"','"&request("menu_desc")&"')"
      call ExecSQL (SQL)
      Response.Redirect "menu_edit.asp?menu_id="&request("menu_id")
   end if
end if   
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>程式群組管理--資料輸入</title>
</head>
<BODY>
<table border="0" width="80%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"  height="25">
        程式群組管理【新增資料】</td>
  </tr>
  <tr>
    <td width="100%">
                  <form name="form" action="" method="post">
                  <input type="hidden" name="action"> 
                    <div align="center">
                      <center>
                    <table width="100%" border=1 cellspacing="1" style="border-collapse: collapse" bordercolor="#6487DC" cellpadding="2">
                      <tr> 
                        <td noWrap align=right width="20%" bgcolor=#FFFFCC height=22>
                        <font color="#FF0000">*</font> <font color="#000080">
                        <span lang="zh-tw">程式群組</span>：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="12" --><input class=font9 size=20 
                        name="menu_id" maxlength="12"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" height=22 bgcolor="#FFFFCC">
                        <font color="#000080"><span lang="zh-tw">連結URL</span>：</font></td>
                        <td width="80%" height=22 bgcolor="#FFFCE1"> 
                          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="40" --><input class=font9 
                        size=40 name=menu_url maxlength="40">
                        </td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" bgcolor=#FFFFCC height=22 class="td1">
                        <font color="#000080"><span lang="zh-tw">程式群組說明</span>：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="60" --><input class=font9 
                        size=60 name=menu_desc maxlength="60"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" height=22 bgcolor="#FFFFCC" class="td1">
                        <font color="#000080"><span lang="zh-tw">排序</span>：</font></td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1"> 
                          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="2" --><input type="text" name="menu_seq" size="4" class="font9" maxlength="2"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" bgcolor=#FFFFCC height=22 class="td1">
                        <span lang="zh-tw">安全</span><font color="#000080"><span lang="zh-tw">控管</span>：</font></td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1">
                          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="1" --><input type="text" name="security" size="4" class="font9" maxlength="1"></td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 height=22 class="font9" colspan="2">
                         <%SaveButton%>　
                        </td>
                        </table>
                      </center>
                    </div>
                  </form>
      </td>
  </tr>
</table>

<%message%>           

</BODY>
</HTML>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--

sub save_OnClick
  <%call CheckString("menu_id","程式群組")%>
  <%call Checklen("menu_id",20,"程式群組")%>
  <%call Checklen("menu_desc",60,"程式群組說明")%>
  form.action.value="save"
  form.submit   
end sub

sub cancel_OnClick
  history.back
end sub

sub query_OnClick
  location.href="menu.asp"
end sub  

--></script>