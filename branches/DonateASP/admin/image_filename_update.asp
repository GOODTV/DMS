<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From UPLOAD Where Ser_No='"&request("ser_no")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  If request("Attach_Type")="YouTube" Or request("Attach_Type")="無名小站" Or request("Attach_Type")="WMV" Then
    RS("Upload_FileURL")=request("Upload_FileURL")
  End If
  RS("Upload_FileName")=request("Upload_FileName")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="檔案說明修改成功 !"
  If request("item")="epaper" Then
    response.redirect "epaper_content_edit.asp?ser_no="&request("object_id")
  Else  
    response.redirect request("item")&"_edit.asp?code_id="&request("code_id")&"&ser_no="&request("object_id")
  End If
End If

SQL="Select * From UPLOAD Where ser_no='"&request("ser_no")&"'"
Call QuerySQL(SQL,RS)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>檔案說明修改</title>
</head>
<body class=gray>
  <div align="center"><center>
    <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td width="100%">
          <form name="form" action="image_filename_update.asp" method="post" target="main">
            <input type="hidden" name="action">
            <input type="hidden" name="item" value="<%=request("item")%>">
            <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
            <input type="hidden" name="object_id" value="<%=RS("object_id")%>">
            <input type="hidden" name="Attach_Type" value="<%=RS("Attach_Type")%>">
            <input type="hidden" name="code_id" value="<%=request("code_id")%>">	
            <div align="center"><center>
              <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" cellpadding="2">
                <tr> 
                  <td align=right width="15%" height=18><font color="#000080">標題：</font></td>
                  <td valign=center width="85%" height=18><input type="text" name="subject" size="50" class="font9t" readonly value="<%=request("subject")%>"></td>
                </tr>
                <%If RS("Attach_Type")="YouTube" Or RS("Attach_Type")="無名小站" Or RS("Attach_Type")="WMV" Or RS("Attach_Type")="Flickr" Then%>
                <tr>
                  <td align=right height=18><font color="#000080">連結網址：</font></td>
                  <td valign=center height=18><input type="text" name="Upload_FileURL" size="50" maxlength="600" value='<%=rs("Upload_FileURL")%>' class="font9"></td>
                </tr>
                <tr>
                  <td align=right height=18><font color="#000080">連結說明：</font></td>
                  <td valign=center height=18><input type="text" name="Upload_FileName" size="50" maxlength="200" value="<%=rs("Upload_FileName")%>" class="font9"></td>
                </tr>
                <%Else%>
                <tr>
                  <td align=right height=18><font color="#000080">檔案說明：</font></td>
                  <td valign=center height=18><input type="text" name="Upload_FileName" size="50" maxlength="200" value="<%=rs("Upload_FileName")%>" class="font9"></td>
                </tr>
                <%End If%>
                <tr>
                  <td width="100%" height=15 class="font9" colspan="2" align="center">
                    <input type="button" value=" 存 檔 " name="save" class="cbutton" onclick="Save_OnClick()">&nbsp;
                    <input type="button" value=" 取 消 " name="cancel" class="cbutton" onclick="window.close();"></td>
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
  if(document.form.Attach_Type.value=='YouTube'||document.form.Attach_Type.value=='無名小站'||document.form.Attach_Type.value=='WMV'){
    <%call CheckStringJ("Upload_FileURL","連結網址")%>
    <%call ChecklenJ("Upload_FileURL",600,"連結網址")%>
    <%call ChecklenJ("Upload_FileName",200,"連結說明")%>
  }else{
    <%call ChecklenJ("Upload_FileName","200","檔案說明")%>
  }
  <%call SubmitJ("update")%>
  window.close();
}
--></script>