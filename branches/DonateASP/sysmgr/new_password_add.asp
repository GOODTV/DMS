<!--#include file="../include/dbfunction.asp"-->
<%
   log_type="更改密碼"
   log_desc="系統登入成功,更改密碼"
   call syslog (log_type,log_desc)
   sql="select password_day from dept where dept_id='"&session("dept_id")&"'"
   call QuerySQL(SQL,RS)
   if not RS.EOF then
      expire_date=DateAdd("d",RS("password_day"),date())   
   end if   
   sql="update userfile set password='"&request("password")&"',old_password1='"&request("old_password")&"',"_
+"       old_password5=old_password4,old_password4=old_password3,old_password3=old_password2,old_password2=old_password1,pwd_validDate='"&expire_date&"'"_
+"        where user_id = '"&request("user_id")&"' "
   call UpdateSQL (SQL)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>密碼登入訊息</title>
</head>
<body background="back_sub.GIF" bgcolor="#FFFFFF" text="#000000" link="#000080"
vlink="#000080">
<p></p>
<div align="center">
  <center>

<table border="1" width="50%" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td width="462" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" class="font11"><strong>
    密碼更改訊息</strong></td>
  </tr>
  <tr>
    <td bgcolor="#C0C0C0"><table
    BORDER="0" width="373">
      <tr>
        <td width="49" valign="middle" align="center" height="80">
        <img src="../images/ok.gif"></td>
        <td valign="middle" width="310" height="80">    
          <p><font color="#800000"><strong><font size="3">密碼更改成功</font> 
          </strong>  
          !</font></p>    
          <p><b><font color="#800000">下次進入系統, 請使用新密碼</font></b></p>    
        </td>    
      </tr>    
      <tr> 
        <td width="365" valign="top" align="right" colspan="2"> 
          <hr> 
        </td> 
      </tr>    
      <tr> 
        <td width="49" valign="top" align="right"></td> 
        <td valign="top" width="310">
        <input type="button" value=" 確定 " name="B1" OnClick="location.href='main.asp'" class="cbutton">    
        </td>    
      </tr>    
    </table>     
    </td>    
  </tr>    
</table>    
  </center>    
</div>    
</body>    
</html>