<!--#include file="../include/dbfunction.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>檔案上傳管理</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="POST" action="source_list.asp" target="detail">
      <input type="hidden" name="action">
      <div align="center"><center>
        <table border="1" width="850" cellspacing="0">
          <tr>
            <td>   
              <table border="0" width="100%" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
                <tr>
	          <td width="100%" class="font11" height="25" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"><b><font color="#FFFFFF">&nbsp;&nbsp;檔案上傳管理</font></b></td>
                </tr>
                <tr>
                  <td align="center">
                    關鍵字<span lang="en-us"> :</span>
                    <input type="text" name="keyword" size="20" class="font9">&nbsp;&nbsp;&nbsp;&nbsp;                  
                    <input type="button" value=" 查 詢 " name="query" class="cbutton">
                  </td>
                </tr>
                <tr>
                  <td width="100%"><iframe name="detail" src="source_list.asp" height="465" width="100%" frameborder="0" scrolling="auto"></iframe></td>
                </tr>
              </table>
            </td>
          </tr>  
        </table>
      </center></div>
    </form>
  </center>
  <%message%>
</div>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<Script Language="VBScript"><!--
sub query_OnClick
  form.action.value="query"
  form.submit
end sub
--></Script>