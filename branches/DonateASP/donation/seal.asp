<!--#include file="../include/dbfunction.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <title>印鑑管理管理</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
  <form name="form" method="POST" action="seal_list.asp" target="detail">
    <input type="hidden" name="action">
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
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">印鑑管理管理</td>
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
                    <td width="150"> </td>
                    <td>關鍵字<span lang="en-us">:</span>&nbsp;<input type="text" name="Seal_Subject" size="20" class="font9">&nbsp;<input type="button" value=" 查 詢 " name="query" class="cbutton"></td>
                    <td width="100" align="right"> </td>		
                  </tr>
                  <tr>
                    <td class="td02-c" colspan="9" width="100%"><iframe name="detail" src="seal_list.asp?dept_id=<%=session("dept_id")%>&SQL=<%=request("SQL")%>&nowPage=<%=request("nowPage")%>" height="380" width="100%" frameborder="0" scrolling="auto"></iframe></td>
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
<Script Language="VBScript">
sub query_OnClick
  form.action.value="query"
  form.target="detail"
  form.submit
end sub  
</Script>