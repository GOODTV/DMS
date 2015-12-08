<!--#include file="../include/dbfunctionJ.asp"-->
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>讀者統計</title>
  <link rel="stylesheet" type="text/css" href="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <form name="form" method="post" action="member_stat_list_qry.asp" target="main">
      <input type="hidden" name="action">
      <table width="700" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="100%"  border="0" cellspacing="0" cellpadding="0" align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"></td>
                <td width="95%">
  		            <table width="60%"  border="0" cellspacing="0" cellpadding="0">
		                <tr>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                      <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                    </tr>
                    <tr>
                      <td class="table62-bg">　</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">讀者統計</td>
                      <td class="table63-bg">　</td>
                    </tr>
    	            </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td>
	          <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">　</td>
                <td valign="top">
                  <table width="100%" border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111">
                    <tr>
                      <td class="td02-c" align="right">統計項目：</td>
                      <td class="td02-c" colspan="3">
                      	<!--20131003 Modify by GoodTV Tanya-->
                      	<!--<input type="radio" name="Stat_Kind" id="Stat_Kind1" value="Category" checked >類別&nbsp;-->
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind2" value="Sex">性別&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind3" value="Age">年齡&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind4" value="Education">教育程度&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind5" value="Occupation">職業別&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind6" value="Marriage">婚姻狀況&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind7" value="Religion">宗教信仰&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind8" value="City">通訊縣市&nbsp;
                      	<!--<input type="radio" name="Stat_Kind" id="Stat_Kind9" value="Invoice_City">收據縣市<br />-->
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind10" value="Member_Status">狀態&nbsp;
                      	<!--<input type="radio" name="Stat_Kind" id="Stat_Kind11" value="Member_Type">會員別-->
                      </td>
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick()'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick()'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"
    			              Response.Write "<span id='movebar' style='display:none'><img border='0' src='../images/movebar.gif'></span>"
    			            %>
              　      </td>
                    </tr>
                  </table>
                </td>
                <td class="table63-bg">　</td>
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
  </center></div></p>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Report_OnClick(){
  if(confirm('您是否確定要將查詢結果列表？')){
    document.form.target="report"
    document.form.action.value="report"
    window.open('','report','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600')
    document.form.submit();
  }
}

function Export_OnClick(){
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
}
--></script>