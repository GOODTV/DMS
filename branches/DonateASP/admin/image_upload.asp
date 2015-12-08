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
  
  album_id=""
  album_seq=1
  If request("item")="album" Then
    album_id=request("ser_no")
    SQL="Select Album_Seq=IsNull(Max(Upload_Seq),0) From UPLOAD Where Ap_Name='album' And Object_ID='"&request("ser_no")&"'"
    call QuerySQL(SQL,RS)
    If Not RS.EOF Then
      album_seq=Cstr(Cint(RS("Album_Seq"))+1)
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
          <form name="form" action="image_upload_add.asp?maxfileSize=<%=MaxFileSize%>&album_id=<%=album_id%>" method="post" enctype="multipart/form-data" target="main">
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
                <%If request("item")<>"signup" And request("item")<>"periodical" Then%>
                <tr> 
                  <td align=right height=18><font color="#000080">檔案類型：</font></td>
                  <td height=18> 
                    <font color="#000080">
                      <%If request("item")="signup" Then%>
                        <input type="radio" name="type" id="doc" value="doc" checked >報名簡章
                      <%ElseIf request("item")="periodical" Then%>
                        <input type="radio" name="type" id="doc" value="doc" checked >PDF檔	
                      <%Else%>
                      <input type="radio" name="type" id="pic" value="image" checked >圖檔
                      <input type="radio" name="type" id="flash" value="flash" <%If request("Ads_Item")="flash" Or request("links_Item")="flash" Then%>checked<%End If%> >Flash動畫       
		                  <%End If%>
		                </font>
		              </td>
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
                <%Else%>
                <input type="hidden" name="type" value="doc">
                <%End If%>
                <%If request("item")="album_photo" Then%> 
                <tr> 
                  <td align=right height=18><font color="#000080">相片日期：</font></td>
                  <td height=18>
                    <%call Calendar("albumdate",date())%>
                    &nbsp;&nbsp;&nbsp;&nbsp;<font color="#000080">排序：</font>
                    <input type="text" name="albumseq" size="5" class="font9" maxlength="5" value="<%=album_seq%>">
                  </td>
                </tr>
                <!--#include file="../include/calendar2.asp"-->
                <%Else%>
                <input type="hidden" name="albumdate" value="<%=date()%>">
                <input type="hidden" name="albumseq" value="0">
                <%End If%>                             
                <%If request("item")<>"ads" And request("item")<>"links" Then%>   
                <tr> 
                  <td align=right height=18><font color="#000080">檔案說明：</font></td>
                  <td height=18><input type="text" name="filename" size="60" class="font9" maxlength="200"></td>
                </tr>
                <%Else%>
                <input type="hidden" name="filename" value="">
                <%End If%>            
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
  if(document.form.item.value=='signup'||document.form.item.value=='periodical'){
    <%call CheckExtJ("upload_FileURL","doc","附加檔案")%>
  }else if(document.form.item.value=='ads'||document.form.item.value=='links'||document.form.item.value=='album_photo'){
    if(document.form.type[0].checked){
      <%call CheckExtJ("upload_FileURL","img","附加檔案")%>
    }else if(document.form.type[1].checked){
      <%call CheckExtJ("upload_FileURL","flash","附加檔案")%>
    }  
  }else{
    if(document.form.type[0].checked){
      <%call CheckExtJ("upload_FileURL","img","附加檔案")%>
    }else if(document.form.type[1].checked){
      <%call CheckExtJ("upload_FileURL","flash","附加檔案")%>
    }else if(document.form.type[2].checked){
      <%call CheckExtJ("upload_FileURL","doc","附加檔案")%>
    }  
  }
  if(document.form.item.value!='signup'&&document.form.item.value!='periodical'){
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
  <%call CheckStringJ("albumseq","排序")%>
  <%call CheckNumberJ("albumseq","排序")%>      
  <%call SubmitJ("save")%>
  window.close();
}
function Resize_OnFocus(){
  document.form.type[0].checked=true;
  document.form.resize[1].checked=true;
}
--></script>