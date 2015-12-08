<% response.buffer=true
   response.expires=-1 
   Set Conn = Server.CreateObject("ADODB.Connection")
   conn.Provider="sqloledb"
   conn.open "server="&session("server")&";uid="&session("uid")&";pwd="&session("pwd")&";database="&session("database")
   session("errnumber")=session("errnumber")+1
   if session("errnumber")=6 then
      sql="update userfile set user_lock='Y' where user_id='"&session("user_id")&"'"
      Set RS = Conn.Execute(sql)
   end if   
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="_function/dms.css">
  <title>密碼登入訊息</title>
</head>

<body background="back_sub.GIF" bgcolor="#FFFFFF" text="#000000" link="#000080"
vlink="#000080">

<div align="center">
  <center>
<br>
<br>
<br>
<br>
<table border="1" width="50%" cellspacing="0">
  <tr>
    <td width="462" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"><font size="3"><strong>密碼登入訊息</strong></font></td>
  </tr>
  <tr>
    <td bgcolor="#C0C0C0"><table
    BORDER="0" width="100%">
      <tr>
        <td width="36" valign="top" align="right"><img src="../images/delete.gif" alt="connect.GIF (5711 bytes)"></td>
        <td valign="top" width="406" class="font11"><p><font color="#0000FF"><strong>密碼登入錯誤</strong> 
        !</font></p>    
        <p>請重新輸入正確之密碼</td>    
      </tr>    
      <tr> 
        <td width="442" valign="top" align="right" colspan="2"> 
          <hr> 
        </td> 
      </tr>    
      <tr> 
        <td width="36" valign="top" align="right"></td> 
        <td valign="top" width="406">
        <input type="button" value=" 確 定 " name="B1" OnClick="location.href='index.asp?loginerr=Y'"    
       class="cbutton">    
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