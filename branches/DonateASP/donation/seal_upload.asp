<!--#include file="../include/dbfunction.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</HEAD>
<BODY class=gray>
<div align="center">
  <center>
<table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
	    <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25">
        <font color="#ffffff">上傳電子圖檔</font></td>
  </tr>

  <tr>
    <td width="100%">
                  <form name="form" action="seal_upload_add.asp" method="post" enctype="multipart/form-data" target="main">
                    <div align="center">
                      <center>
                    <input type="hidden" name="ser_no" value="<%=request("Ser_No")%>">
                    <input type="hidden" name="SealSubject" value="<%=request("Seal_Subject")%>">
                    <table width="100%" border=1 cellspacing="0" cellpadding="2">
                      <tr> 
                        <%If Cint(request("Ser_No"))<=4 Then%>
                        <td align=right width="22%" height=22><font color="#000080">印鑑標題：</font></td>
                        <%Else%>
                        <td width="22%" valign="top" align="right"><font color="#000080">經手人姓名<span lang="en-us">:</span></font></td>
                        <%End If%>
                        <td valign=center width="78%" height=18> <input type="text" class=font9t size=20 name="Seal_Subject" readonly value="<%=request("Seal_Subject")%>"></td>
                      </tr>
                      <tr> 
                        <td align=right height=22><font color="#000080">上傳電子印鑑圖檔：</font></td>
                        <td valign=center height=22>
                        	<input type="file" class=font9 size=57 name=upload_FileURL>
                          <%If Cint(request("Ser_No"))>4 Then%><br /><font color="#000080">圖檔名稱須有『&nbsp;<font color="#ff0000"><%=request("Seal_Subject")%></font>&nbsp;』字樣(&nbsp;<font color="#ff0000">如:<%=request("Seal_Subject")%>印張.JPG</font>&nbsp;)否則系統將無法判斷</font><%End If%>
                        </td>
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