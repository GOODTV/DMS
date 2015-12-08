<!--#include file="../include/dbfunction.asp"-->
<%
if request("action") = "update" then
   sql="update menu set menu_seq='"&request("menu_seq")&"',menu_desc='"&request("menu_desc")&"',"_
+"      menu_url='"&request("menu_url")&"',security='"&request("security")&"'"_
+"      where menu_id = '"&request("menu_id")&"' "
   call UpdateSQL (SQL)
   Response.Redirect("menu.asp")
end if

if request("action") = "delete" then
   sql="delete from menu where menu_id = '"&request("menu_id")&"' "_
+"      delete from prog_menu where menu_id = '"&request("menu_id")&"'"_
+"      delete from user_menu where menu_id = '"&request("menu_id")&"'"   
   call UpdateSQL (SQL)
   Response.Redirect("menu.asp")
end if

  sql="select * from menu where menu_id='"&request("menu_id")&"'"
  call QuerySQL (SQL,RS)
  session("menu_id")=request("menu_id")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>程式群組管理--資料輸入</title>
</head>
<BODY>
<form name="form" action="" method="post">
  <input type="hidden" name="action"> 
  <input type="hidden" name="AddButton" value="menu_add.asp">
<table border="0" width="80%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="font11" colspan="2" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"  height="25">
        程式群組管理【修改資料】</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
                    <div align="center">
                      <center>
                    <table width="100%" border=1 cellspacing="1" style="border-collapse: collapse" bordercolor="#6487DC" cellpadding="2">
                      <tr> 
                        <td noWrap align=right width="19%" height=22 class="td1"><font color="#FF0000">* </font>程式群組：</td>
                        <td width="80%" height=22 class="td2"> 
                          <input class=font9 size=20 readonly name="menu_id" value="<%=RS("menu_id")%>"></td>
                      </tr>
                      <tr> 
                        <td align=right width="19%" height=22 bgcolor="#FFFFCC">
                        <font color="#000080"><span lang="zh-tw">連結URL</span>：</font></td>
                        <td width="80%" height=22 bgcolor="#FFFCE1"> 
                          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="40" --><input class=font9 
                        size=40 name=menu_url maxlength="40" value="<%=rs("menu_url")%>">
                        </td>
                      </tr>
                      <tr> 
                        <td align=right width="19%" bgcolor=#FFFFCC height=22 class="td1">
                        <font color="#000080"><span lang="zh-tw">程式群組說明</span>：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="60" --><input class=font9 
                        size=60 name=menu_desc maxlength="60" value="<%=rs("menu_desc")%>"></td>
                      </tr>
                      <tr> 
                        <td align=right width="19%" height=22 bgcolor="#FFFFCC" class="td1">
                        <font color="#000080"><span lang="zh-tw">排序</span>：</font></td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1"> 
                          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="2" --><input type="text" name="menu_seq" size="4" class="font9" maxlength="2" value="<%=rs("menu_seq")%>"></td>
                      </tr>
                      <tr> 
                        <td align=right width="19%" bgcolor=#FFFFCC height=22 class="td1">
                        <span lang="zh-tw">安全</span><font color="#000080"><span lang="zh-tw">控管</span>：</font></td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1">
                          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="1" --><input type="text" name="security" size="4" class="font9" maxlength="1" value="<%=rs("security")%>"></td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 height=22 class="font9" colspan="2">
                         <%EditButton()%>
                        </td>
                        </table>
                      </center>
                    </div>
  <tr>
    <td width="56%" align="right">
      <font color="#800000"><span style="font-size: 11pt">程式群組設定</span></font> </td>
    <td width="44%" align="center">
      <input type="button" value="程式設定" name="setup" class="cbutton"></td>
  </tr>
  <tr>
    <td width="100%" align="center" colspan="2">
      <% 
       SQL="select ser_no,序號=prog_seq,程式代號=progid,程式名稱=prog_desc from prog_menu inner join prog on prog_menu.progid=prog.prog_id"_
+"           where menu_id='"&request("menu_id")&"' order by prog.prog_seq"
      HLink="menu_setup_delete.asp?ser_no=" 
      LinkParam="ser_no"
      LinkTarget="_self"
      call DataGrid (SQL,HLink,LinkParam,LinkTarget) 
%> 
    </td>
  </tr>
      </td>
  </tr>
</table>
<%message%>           
</form>
</BODY>
</HTML>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--

sub Update_OnClick
  <%call Checklen("menu_desc",60,"程式群組說明")%>
  ConfirmUpdate()
end sub

sub setup_OnClick
  location.href="menu_setup.asp?menu_id=<%=request("menu_id")%>"
end sub                                             

sub cancel_OnClick
  history.back
end sub

sub query_OnClick
  location.href="menu.asp"
end sub  

--></script>