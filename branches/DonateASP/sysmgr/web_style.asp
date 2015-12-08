<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>系統參數設定</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="POST" action="web_style_list.asp" target="detail">
      <div align="center"><center>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
          <tr>
            <td>
              <table width="800"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
                <tr>
                  <td width="5%"></td>
                  <td width="95%">
  		              <table width="40%"  border="0" cellspacing="0" cellpadding="0">
		                  <br>
		                  <tr>
                        <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                        <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                        <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                      </tr>
                      <tr>
                        <td class="table62-bg"></td>
                        <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">系統參數設定</td>
                        <td class="table63-bg"></td>
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
                  <td class="table62-bg"></td>
                  <td align="center">
                    關鍵字<span lang="en-us"> :</span>
                    <input type="text" name="keyword" size="20" class="font9">&nbsp;&nbsp;&nbsp;&nbsp;
                    類別<span lang="en-us"> :</span>
                    <%
                      SQL="Select Distinct Style_Type From WEB_STYLE Where Style_Type<>'MD5'"
                      FName="Style_Type"
                      Listfield="Style_Type"
                      BoundColumn=""
                      menusize="1"
                      call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                    %>&nbsp;&nbsp;&nbsp;&nbsp;
                    <input type="button" value=" 查 詢 " name="query" class="cbutton" onClick="Query_OnClick()">
                  </td>
                  <td class="table63-bg"></td>
                </tr>                
                <tr>
                  <td class="table62-bg"></td>
                  <td valign="top">
                    <table width="100%"  border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
                      <tr>
                        <td class="td02-c" width="100%"><iframe name="detail" src="web_style_list.asp" height="450" width="100%" frameborder="0" scrolling="auto"></iframe></td>
                      </tr>
                    </table>
                  </td>
                  <td class="table63-bg"></td>
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
      </center></div>
    </form>
  </center></div>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Query_OnClick(){
  <%call SubmitJ("query")%>
}
--></script>
