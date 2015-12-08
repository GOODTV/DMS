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
  <title>上傳標題圖檔</title>
</head>
<body class=gray>
  <center>
    <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td width="100%">
          <form name="form" action="image_title_upload_add.asp?maxfileSize=<%=MaxFileSize%>" method="post" enctype="multipart/form-data" target="main">
            <input type="hidden" name="item" value="<%=request("item")%>">
            <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
            <input type="hidden" name="album_id" value="<%=request("album_id")%>">
            <input type="hidden" name="code_id" value="<%=request("code_id")%>">
            <input type="hidden" name="ads_item" value="<%=request("ads_item")%>">	
            <div align="center"><center>
              <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" cellpadding="2">
                <tr> 
                  <td align=right width="22%" height=18><font color="#000080">標題：</font></td>
                  <td valign=center width="78%" height=18><input type="text" name="subject" size="30" class="font9t" readonly value="<%=request("subject")%>"></td>
                </tr>
                <%If request("ads_item")="flash" Then%>
                <tr>
                  <td align=right height=18><font color="#000080">上傳flash：<br><font color="#ff0000">(flash大小限制<%=MaxFileSize%>M)</font></font></td>
                  <td valign=center height=18><input type="file" class=font9 size=50 name=upload_FileURL></td>
                </tr>
                <input type="hidden" name="resize" value="N">
                <input type="hidden" name="img_width" value="">
                <input type="hidden" name="img_height" value="">		
                <%Else%>
                <tr> 
                  <td align=right height=18><font color="#000080">上傳標題圖檔：<br><font color="#ff0000">(圖檔大小限制<%=MaxFileSize%>M)</font></font></td>
                  <td valign=center height=18><input type="file" class=font9 size=50 name=upload_FileURL></td>
                </tr>
                <tr> 
                  <td align=right height=18><font color="#000080">圖檔縮圖：</font></td>
                  <td height=18>
                    <font color="#000080"> 
		                  <input type="radio" name="resize" value="N" id="resizeNo" <%If request("img_height")="" And request("img_width")="" Then%>checked<%End If%> >原圖大小
		                  <input type="radio" name="resize" value="Y" id="resizeYes" <%If request("img_height")<>"" Or request("img_width")<>"" Then%>checked<%End If%> >縮圖： 
		                  <%If request("img_height")<>"" Then%>
		                  寬度<input type="text" name="img_width" size="8" class="font9" value="" OnFocus="Resize_OnFocus()">
		                  或高度<input type="text" name="img_height" size="8" class="font9" value="<%=request("img_height")%>" OnFocus="Resize_OnFocus()">			                  
		                  <%Else%>
		                  寬度<input type="text" name="img_width" size="8" class="font9" value="<%=request("img_width")%>" OnFocus="Resize_OnFocus()">
		                  或高度<input type="text" name="img_height" size="8" class="font9" value="" OnFocus="Resize_OnFocus()">
		                  <%End If%>
		                  pixal&nbsp;
                    </font>
                  </td>
                </tr>
                <%End If%>
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
  if(document.form.ads_item.value=='flash'){
    <%call CheckStringJ("upload_FileURL","上傳flash")%>
    <%call CheckExtJ("upload_FileURL","flash","上傳flash")%>
  }else{
    <%call CheckStringJ("upload_FileURL","上傳標題圖檔")%>
    <%call CheckExtJ("upload_FileURL","img","上傳標題圖檔")%>
    if(document.form.resize[1].checked){
      if(document.form.img_width.value==''&&document.form.img_height.value==''){
        alert('寬度或高度  欄位不可為空白！');
        document.form.img_width.focus();
        return;
      }
      if(document.form.img_width.value!=''){
        <%call CheckNumberJ("img_width","寬度")%>
      }
      if(document.form.img_width.img_height!=''){
        <%call CheckNumberJ("img_height","高度")%>
      }  
    }
  }
  <%call SubmitJ("save")%>
  window.close();
}
function Cancel_OnClick(){
  window.close();
}
--></script>