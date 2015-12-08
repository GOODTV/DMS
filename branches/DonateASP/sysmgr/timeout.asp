<%
response.buffer =true 
response.expires=-1
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>Timeout Message</title>
</head>
<body>
  <br><br><br>
  <div align="center"><center>
  <table border="0" width="300" height="160" align="center" cellspacing="0">
    <tr>
      <td width="100%" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" class="font11"><font color="#FFFFFF">系統訊息</td>
    </tr>
    <tr>
      <td width="100%" align="center">
        <br><font face="標楷體" size="4" color="#800000"><strong>系統暫停超過時間,  請重新登入</strong></font><hr> 
        <form name="form" method="GET" action="../admin/index.asp" target="_top"> 
          <button id=action style="width:45;height:40;font-size:9pt"><img src="../images/gnicok.gif"><br>確 定 </button> 
        </form>
      </td> 
    </tr> 
  </table> 
  </center></div> 
</body> 
</html> 
<script language="VBScript"><!-- 
sub action_OnClick 
  form.submit 
end sub 
 --></script>