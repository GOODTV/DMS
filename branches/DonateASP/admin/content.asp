<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>網站內容管理</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="POST" action="content_list.asp" target="detail">
      <input type="hidden" name="action">
      <div align="center"><center>
        <table border="1" width="830" cellspacing="0">
          <tr>
            <td>   
              <table border="0" width="100%" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
                <tr>
	                <td width="100%" class="font11" height="25" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"><b><font color="#FFFFFF">&nbsp;&nbsp;網頁內容管理</font></b></td>
                </tr>
                <tr>
                  <td align="center">
                    關鍵字<span lang="en-us"> :</span>
                    <input type="text" name="KeyWord" size="20" class="font9">&nbsp;&nbsp;&nbsp;&nbsp;
                    主功能<span lang="en-us"> :</span>
                    <%
                      SQL="Select Content_Type=CodeName From CodeFile Where Category='單一網頁' Order By Menu_Seq"
                      FName="Content_Type"
                      Listfield="Content_Type"
                      BoundColumn=""
                      menusize="1"
                      call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                    %>&nbsp;&nbsp;&nbsp;&nbsp;
                    <input type="button" value=" 查 詢 " name="query" class="cbutton" onClick="Query_OnClick()">
                  </td>
                </tr>
                <tr>
                  <td width="100%"><iframe name="detail" src="content_list.asp" height="465" width="100%" frameborder="0" scrolling="auto"></iframe></td>
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