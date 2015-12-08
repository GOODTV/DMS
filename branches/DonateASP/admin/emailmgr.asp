<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>郵件管理</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="POST" action="emailmgr_list.asp" target="detail">
      <input type="hidden" name="action">
      <div align="center"><center>
        <table border="1" width="830" cellspacing="0">
          <tr>
            <td>   
              <table border="0" width="100%" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
                <tr>
	                <td width="100%" class="font11" height="25" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"><b><font color="#FFFFFF">&nbsp;&nbsp;郵件管理</font></b></td>
                </tr>
                <tr>
                  <td align="center">
                    關鍵字<span lang="en-us"> :</span>
                    <input type="text" name="keyword" size="20" class="font9">&nbsp;&nbsp;&nbsp;&nbsp;
                    <input type="button" value=" 查 詢 " name="query" class="cbutton" onClick="Query_OnClick()">
                  </td>
                </tr>
                <tr>
                  <td width="100%"><iframe name="detail" src="emailmgr_list.asp?action=query" height="465" width="100%" frameborder="0" scrolling="auto"></iframe></td>
                </tr>
              </table>
            </td>
          </tr>  
        </table>
      </center></div>
    </form>
  </center></div></p>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Query_OnClick(){
  <%call SubmitJ("query")%>
}
--></script>