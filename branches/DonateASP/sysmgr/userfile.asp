<!--#include file="../include/dbfunction.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>帳號管理</title>
</head>
<body>
<p>
<div align="center">
  <center>
<form name="form" method="POST" action="userfile_list.asp" target="detail">
<input type="hidden" name="action">
<div align="center">
  <center>
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
          <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">
          帳號管理</td>
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
	<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
        <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
        <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
      </tr>
      <tr>
        <td class="table62-bg">&nbsp;</td>
        <td valign="top">
        <table width="100%"  border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
            <tr>
              <td class="td02-c" width="128" align="right">部門:</td>
              <td class="td02-c" width="159">
              <%sql="select dept_id,dept_desc from dept order by dept_id"
                          FName="dept_id"                                                                                                                                                                                                                                                              
                          Listfield="dept_desc"                                                                                                                                                                                          
                          menusize="1"                                               
                          BoundColumn=""                                                                                                                                                                  
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                         %>
              　</td>
              <td class="td02-c" width="70" align="right">帳號類別:</td>
              <td class="td02-c" width="108">
              <%sql="select user_group from groupfile where '"&session("user_group")&"'='系統管理員' or"_
             +"       ('"&session("user_group")&"'<>'系統管理員' and user_group<>'系統管理員')"
                          FName="user_group"                                                                                                                                                                                                                                                              
                          Listfield="user_group"                                                                                                                                                                                          
                          menusize="1"                                               
                          BoundColumn=""                                                                                                                                                                  
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                         %>
              　</td>
              <td class="td02-c" width="407">
              <input type="submit" value=" 查 詢 " name="query1" class="cbutton"></td>
              </tr>
            <tr>
              <td width="99%" colspan="5" bgcolor="#C0C0C0" height="1"> </td>
              </tr>
            <tr>
              <td class="td02-c" colspan="5" width="100%">
              <iframe name="detail" src="userfile_list.asp" height="420" width="100%" frameborder="0" scrolling="auto"></iframe>
              </td>
              </tr>
        </table></td>
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
      </center>
</div>
      </form>
      <%message%>
  </center>
</div>
</body>

</html>
<!--#include file="../include/dbclose.asp"-->