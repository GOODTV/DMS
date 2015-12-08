<!--#include file="../include/dbfunction.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>檔案上傳</title>
</head>
<BODY class=gray>
<div align="center">
  <center>
<table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
	    <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25">
        <font color="#ffffff">檔案上傳</font></td>
  </tr>

  <tr>
    <td width="100%">
                  <form name="form" action="logo_upload_add.asp" method="post" enctype="multipart/form-data" target="main">
                    <div align="center">
                      <center>
                    <table width="100%" border=1 cellspacing="0" cellpadding="2">
                      <tr> 
                        <td align=right width="15%" height=22>
                        <font color="#000080">部門代號：</font></td>
                        <td valign=center width="85%" height=18> 
                          <input type="text" name="dept_id" size="20" class="font9t" readonly value="<%=request("dept_id")%>"> 
							(最佳LOGO圖型大小為寬 50 pix , 高 50pix )</td>
                      </tr>
                      <tr> 
                        <td align=right width="15%" height=22>
                        <font color="#000080">上傳檔案：</font></td>
                        <td valign=center width="85%" height=18> 
                          <input type="file" class=font9 size=60 name=upload_FileURL></td>
                      </tr>
                      <tr>
                        <td width="100%" height=15 class="font9" colspan="2" align="center">
                         <input type="submit" value=" 存 檔 " name="save" class="cbutton">
                         <input type="button" value=" 取 消 " name="cancel" class="cbutton"></td>
                        </table>
                      </center>
                    </div>
                  </form>
      </td>
  </tr>
</table>

  </center>
</div>

<%message%>           

</BODY>
</HTML>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--
sub cancel_OnClick
  window.close
end sub

sub save_OnClick
  window.close
end sub

--></script>