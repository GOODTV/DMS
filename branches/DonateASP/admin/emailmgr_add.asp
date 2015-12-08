<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  SQL="EMAILMGR"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS.Addnew
  RS("EmailMgr_Subject")=request("EmailMgr_Subject")
  RS("EmailMgr_Desc")=request("EmailMgr_Desc")
  RS("EmailMgr_ImgPosition")="right"
  RS("EmailMgr_ImageShow_Type")="barrier"
  RS("EmailMgr_RegDate")=Date()
  RS("EmailMgr_Default")="N"
  RS("BranchType")=request("BranchType")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料新增成功  ！"
  SQL="Select @@IDENTITY as ser_no"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF then
    response.redirect "emailmgr_edit.asp?ser_no="&RS("ser_no")
  End If   
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>郵件管理 【新增資料】</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="code_id" value="<%=request("code_id")%>">	
      <table border="0" width="830" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
	        <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25" colspan="2"><font color="#FFFFFF"><b>&nbsp;&nbsp;郵件管理 </b> <font size="2">【新增資料】</font></font></td>
        </tr>
        <tr>
	        <td width="100%" class="font11" valign="top">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">
   		          <tr>
   		            <td>
                    <table width="100%" border="1" cellspacing="0" style="border-collapse: collapse" cellpadding="2" bordercolor="#E1DFCE">
                      <tr>
                        <td noWrap align="right" width="12%"><font color="#000080">郵件標題：</font></td>
                        <td><input class=font9 size=39 name="EmailMgr_Subject" maxlength="100"></td>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2"></td>
                      </tr>
                      <tr>
                        <td align="right"><font color="#000080">郵件內容：</font></td>
                        <td valign="center" colspan="7">
                        <!--#include file="../fckeditor/fckeditor.asp" -->
                        <%
                          Dim oFCKeditor
			                    Set oFCKeditor = New FCKeditor
		                      oFCKeditor.BasePath	= "../fckeditor/"
			                    oFCKeditor.Value	= ""
		                      oFCKeditor.Create "EmailMgr_Desc"
		                    %>
                        </td>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2"></td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 class="font9" colspan="8"><%SaveButton%></td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </center></div>
          </td>
        </tr>
      </table>
    </form>
  </center></div></p>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Save_OnClick(){
  <%call CheckStringJ("EmailMgr_Subject","訊息標題")%>
  <%call ChecklenJ("EmailMgr_Subject",100,"訊息標題")%>
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  location.href='emailmgr.asp';
}
--></script>