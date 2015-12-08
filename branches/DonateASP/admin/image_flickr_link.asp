<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  SQL="UPLOAD"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS.Addnew
  RS("Object_ID")=request("ser_no")
  RS("Ap_Name")=request("item")
  RS("Attach_Type")="Flickr"
  RS("Upload_FileName")=request("Upload_FileName")
  RS("Upload_FileURL")=request("Upload_FileURL")
  RS("Upload_FileSize")="0"
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="相簿連結成功 !"
  response.redirect request("item")&"_edit.asp?code_id="&request("ser_no")&"&ser_no="&request("ser_no")
End If   
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>Flickr相簿</title>
</head>
<body class=gray>
  <div align="center"><center>
    <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td width="100%">
          <form name="form" action="" method="post" target="main">
            <input type="hidden" name="action">
            <input type="hidden" name="item" value="<%=request("item")%>">
            <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
            <input type="hidden" name="code_id" value="<%=request("code_id")%>">	
            <div align="center"><center>
              <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" cellpadding="2">
                <tr> 
                  <td align=right width="20%" height=18><font color="#000080">標題：</font></td>
                  <td width="80%" height=18><input type="text" name="Subject" size="60" class="font9t" readonly value="<%=request("Subject")%>"></td>
                </tr>
                <tr> 
                  <td align=right width="20%" height=18><font color="#000080">相簿連結網址：</font></td>
                  <td width="80%" height=18> <input type="text" name="Upload_FileURL" size="60" class="font9" maxlength="1000"></td>
                </tr>
                <tr> 
                  <td align=right width="20%" height=18><font color="#000080">相簿說明：</font></td>
                  <td width="80%" height=18> <input type="text" name="Upload_FileName" size="60" class="font9" maxlength="200"></td>
                </tr>
                <tr>
                  <td width="100%" height=15 class="font9" colspan="2" align="center">
                    <input type="button" value=" 存 檔 " name="save" class="cbutton" onclick="Save_OnClick()">&nbsp;
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
  <%message%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Save_OnClick(){
  <%call CheckStringJ("Upload_FileURL","相簿連結網")%>
  <%call ChecklenJ("Upload_FileURL",1000,"相簿連結網")%>
  <%call ChecklenJ("Upload_FileName",200,"相簿說明")%>
  <%call SubmitJ("save")%>
  window.close();
}
--></script>