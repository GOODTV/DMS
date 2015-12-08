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
  
  album_id=""
  album_seq=1
  If request("item")="album_photo" Then
    album_id=request("ser_no")
    SQL="Select IsNull(Max(Album_Seq),0) As Album_Seq From ALBUM_PHOTO Where Album_Id='"&request("ser_no")&"'"
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
  <title><%If request("item")="ads" Or request("item")="links" Then%>上傳影音檔案<%ElseIf request("item")="album_photo" Then%>上傳影音檔案<%Else%>附加影音上傳<%End If%></title>
</head>
<body class=gray>
  <div align="center"><center>
    <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td width="100%">
          <form name="form" action="image_av_upload_add.asp?maxfileSize=<%=MaxFileSize%>&album_id=<%=album_id%>" method="post" enctype="multipart/form-data" target="main">
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
                  <td align=right height=18><font color="#000080"><%If request("item")="ads" Or request("item")="links" Then%>上傳影音檔案<%ElseIf request("item")="album_photo" Then%>上傳影音檔案<%Else%>附加影音<%End If%>：<br><font color="#ff0000">(影音大小限制<%=MaxFileSize%>M)</font></font></td>
                  <td valign=center height=18><input type="file" class=font9 size=50 name=upload_FileURL></td>
                </tr>                
                <tr> 
                  <td align=right height=18><font color="#000080">影音類型：</font></td>
                  <td height=18> 
                    <font color="#000080"><span lang="zh-tw">
                      <input type="radio" id="av" name="type" value="av" checked >影音wmv 
                      <input type="radio" id="flv" name="type" value="flv" <%If request("Ads_Item")="flv" Or request("links_Item")="flv" Then%>checked<%End If%>>影音flv
		    </span></font>
		  </td>
                </tr>
                <%If request("item")="album_photo" Then%> 
                <tr> 
                  <td align=right height=18><font color="#000080">影音日期：</font></td>
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
                  <td align=right height=18><font color="#000080">影音說明：</font></td>
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
  <%call CheckStringJ("upload_FileURL","附加影音")%>
  if(document.form.type[0].checked){
    <%call CheckExtJ("upload_FileURL","wmv","影音類型")%>
  }else if(document.form.type[1].checked){
    <%call CheckExtJ("upload_FileURL","flv","影音類型")%>
  }
  <%call CheckStringJ("albumseq","排序")%>
  <%call CheckNumberJ("albumseq","排序")%>  
  <%call SubmitJ("save")%>
  window.close();
}
--></script>