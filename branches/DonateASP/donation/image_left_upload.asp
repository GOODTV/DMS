<!--#include file="../include/dbfunctionJ.asp"-->
<%
  MaxFileSize=2
  SQL="Select MaxFileSize From DEPT Where dept_id='"&request("dept_id")&"'"
  call QuerySQL(SQL,RS)
  If Not RS.EOF Then
    If RS("MaxFileSize")<>"" Then 
      MaxFileSize=RS("MaxFileSize")
    End If
  End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>上傳圖檔</title>
</head>
<body class=gray>
  <div align="center"><center>
    <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td width="100%">
          <form name="form" action=image_left_upload_add.asp?MaxFileSize=<%=MaxFileSize%>" method="post" enctype="multipart/form-data" target="main">
            <input type="hidden" name="item" value="<%=request("item")%>">
            <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
            <input type="hidden" name="code_id" value="<%=request("code_id")%>">	
            <div align="center"><center>
              <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" cellpadding="2">
                <tr> 
                  <td align=right height=18><font color="#000080">上傳圖檔：<br><font color="#ff0000">(圖檔大小限制<%=MaxFileSize%>M)</font></font></td>
                  <td valign=center height=18><input type="file" class=font9 size=50 name=upload_FileURL></td>
                </tr>
                <tr> 
                  <td align=right height=18><font color="#000080">圖檔縮圖：</font></td>
                  <td height=18>
                    <font color="#000080"> 
		      <input type="radio" name="resize" value="N" id="resizeNo" <%If request("img_width")="" Then%>checked<%End If%> >原圖大小
		      <input type="radio" name="resize" value="Y" id="resizeYes" <%If request("img_width")<>"" Then%>checked<%End If%> >縮圖：寬度 
		      <input type="text" name="img_width" size="8" class="font9" value="<%=request("img_width")%>" OnFocus="Resize_OnFocus()">
                    </font>
                  </td>
                </tr>
                <tr>
                  <td width="100%" height=15 class="font9" colspan="2" align="center">
                    <input type="button" value=" 上 傳 " name="save" class="addbutton" onclick="Save_OnClick()">&nbsp;
                    <input type="button" value=" 取 消 " name="cancel" class="delbutton" onclick="Cancel_OnClick()">
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
function Resize_OnFocus(){
  document.form.resize[1].checked=true;
}
function Save_OnClick(){
  <%call CheckStringJ("upload_FileURL","上傳圖檔")%>
  <%call CheckExtJ("upload_FileURL","img","上傳圖檔")%>
  if(document.form.resize[1].checked){
    <%call CheckStringJ("img_width","寬度")%>
    <%call CheckNumberJ("img_width","寬度")%>
  }  
  <%call SubmitJ("save")%>
  window.close();
}
function Cancel_OnClick(){
  window.close();
}
--></script>