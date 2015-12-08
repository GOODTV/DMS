<!--#include file="../include/dbfunctionJ.asp"-->
<%
  MaxFileSize=2
  SQL="Select MaxUploadSize From DEPT Where dept_id='"&request("dept_id")&"'"
  call QuerySQL(SQL,RS)
  If Not RS.EOF Then
    If RS("MaxUploadSize")<>"" Then 
      MaxFileSize=RS("MaxUploadSize")
    End If   
  End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%If request("item")="ads" Or request("item")="links" Then%>上傳標題檔案<%ElseIf request("item")="album_photo" Then%>上傳相片檔案<%Else%>附加檔案上傳<%End If%></title>
</head>
<body class=gray>
  <div align="center"><center>
    <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td width="100%">
          <form name="form" action="image_word_upload_add.asp?maxfileSize=<%=MaxFileSize%>" method="post" enctype="multipart/form-data" target="main">
            <input type="hidden" name="item" value="<%=request("item")%>">
            <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
            <input type="hidden" name="code_id" value="<%=request("code_id")%>">	
            <div align="center"><center>
              <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" cellpadding="2">
                <tr> 
                  <td align=right width="22%" height=18><font color="#000080">標題：</font></td>
                  <td width="78%" height=18><input type="text" name="subject" size="60" class="font9t" readonly value="<%=request("subject")%>"></td>
                </tr>
                <tr> 
                  <td align=right height=18><font color="#000080"><%If request("item")="ads" Or request("item")="links" Then%>上傳檔案<%ElseIf request("item")="album_photo" Then%>上傳相片檔案<%Else%>附加檔案<%End If%>：<br><font color="#ff0000">(檔案大小限制<%=MaxFileSize%>M)</font></font></td>
                  <td valign=center height=18><input type="file" class=font9 size=50 name=upload_FileURL></td>
                </tr>
                <input type="hidden" name="type" value="doc">
                <input type="hidden" name="resize" value="N">
                <input type="hidden" name="img_width" value="">
                <input type="hidden" name="img_height" value="">
                <input type="hidden" name="albumdate" value="<%=date()%>">
                <input type="hidden" name="albumseq" value="0"> 
                <tr> 
                  <td align=right height=18><font color="#000080">檔案說明：</font></td>
                  <td height=18><input type="text" name="filename" size="60" class="font9" maxlength="200"></td>
                </tr>           
                <tr>
                  <td width="100%" height=15 class="font9" colspan="2" align="center">
                    <input type="button" value=" 上 傳 " name="save" class="cbutton" onclick="Save_OnClick()">&nbsp;
                    <input type="button" value=" 取 消 " name="cancel" class="cbutton" onclick="window.close();">
                  </td>
                </tr>
              </table>
            </center></div>
          </form>
        </td>
      </tr>
    </table>
  </center></div>
  <%Message()%>       
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Save_OnClick(){
  <%call CheckStringJ("upload_FileURL","附加檔案")%>
  <%call CheckExtJ("upload_FileURL","doc","附加檔案")%>
  <%call SubmitJ("save")%>
  window.close();
}
--></script>