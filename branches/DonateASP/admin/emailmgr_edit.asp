<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From EMAILMGR Where Ser_No='"&request("ser_no")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("EmailMgr_Subject")=request("EmailMgr_Subject")
  RS("EmailMgr_Desc")=request("EmailMgr_Desc")
  'RS("EmailMgr_ImgPosition")=request("EmailMgr_ImgPosition")
  'RS("EmailMgr_ImageShow_Type")=request("EmailMgr_ImageShow_Type")
  RS("BranchType")=request("BranchType")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 ！"
End If

If request("action")="delete" Then
  SQL="Delete From EMAILMGR Where Ser_No='"&request("ser_no")&"' " &_
      "Delete From NEXTPAGE Where Ser_No='"&request("ser_no")&"' And Page_Type='emailmgr'"
  Set RS=Conn.Execute(SQL)
  SQL1="Select * From UPLOAD Where object_id='"&request("ser_no")&"' And (Ap_Name='emailmgr' Or Ap_Name='emailmgr_title')"
  Call QuerySQL(SQL1,RS1)
  While Not RS1.EOF
    Attach_Type=RS1("Attach_Type")
    Upload_FileURL=RS1("Upload_FileURL")
    Upload_FileURL_Old=RS1("Upload_FileURL_Old")
    SQL2="Delete From UPLOAD Where ser_no="&RS1("ser_no")&""
    Set RS2=conn.Execute(SQL2)
    If Attach_Type<>"YouTube" And Attach_Type<>"無名小站" And Attach_Type<>"WMV" Then
      DataFile="../upload/"&Upload_FileURL&""
      Set objFS = Server.CreateObject("Scripting.FileSystemObject")
      If objFS.FileExists(Server.MapPath(DataFile)) Then objFS.DeleteFile Server.Mappath(DataFile)
      If Upload_FileURL_Old<>"" Then
        DataFile_Old="../upload/"&Upload_FileURL_Old&""
        If objFS.FileExists(Server.MapPath(DataFile_Old)) Then objFS.DeleteFile Server.Mappath(DataFile_Old)
      End If
    Set objFS = Nothing
    End If
    RS1.movenext
  Wend
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="資料刪除成功 ！"
  Response.Redirect "emailmgr.asp"
End If

SQL="Select * From EMAILMGR Where Ser_No='"&request("ser_no")&"'"
call QuerySQL(SQL,RS)
Item="emailmgr"
Subject=RS("EmailMgr_Subject")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>郵件管理 【修改資料】</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
      <table border="0" width="815" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
	        <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25" colspan="2"><font color="#FFFFFF"><b>&nbsp;郵件管理 </b> <font size="2">【修改資料】</font></font></td>
        </tr>
        <tr>
          <td width="20%" class="font11" align="center" valign="top" bgcolor="EEEEE3">
            <iframe name="left1" src="emailmgrtype_list.asp" height="240" width="100%" frameborder="1" scrolling="auto" target="_self"></iframe><br>
            <iframe name="left2" src="image_left_list.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=request("ser_no")%>" height="320" width="100%" frameborder="1" scrolling="auto" target="_self"></iframe>
          </td>
	        <td width="80%" class="font11" valign="top">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">
   	            <tr>
   		            <td>
                    <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" bordercolor="#E1DFCE" cellpadding="2">
                      <tr>
                        <td noWrap align="right" width="12%"><font color="#000080">郵件標題：</font></td>
                        <td><input class=font9 size=39 name="EmailMgr_Subject" maxlength="100" value="<%=RS("EmailMgr_Subject")%>"></td>
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
			                    oFCKeditor.Value	= RS("EmailMgr_Desc")
		                      oFCKeditor.Create "EmailMgr_Desc"
		                    %>
                        </td>
                      </tr>
                      <tr>
                        <td width="100%" colspan="8">
                        <%
                          Response.Write "<button id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Update_OnClick()'> <img src='../images/update.gIf' width='20' height='20' align='absmiddle'> 修改 </button>&nbsp;"
                          If RS("EmailMgr_Default")<>"Y" Then 
                            Response.Write "<button id='delete' style='position:relative;left:0;width:75;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Delete_OnClick()'> <img src='../images/delete.gIf' width='20' height='20' align='absmiddle'> 刪除 </button>&nbsp;"
                            Response.Write "<button id='cancel' style='position:relative;left:0;width:100;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Eail_OnClick()'> <img src='../images/email.gif' width='20' height='15' align='absmiddle'> 郵件發送 </button>&nbsp;"	
                          End If
                          Response.Write "<button id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Cancel_OnClick()'> <img src='../images/icon6.gIf' width='20' height='15' align='absmiddle'> 離開 </button>&nbsp;"
                        %>
                        </td>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2" colspan="7"></td>
                      </tr>
                      <!--#include file="image_upload_show.asp"-->
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
function Update_OnClick(){
  <%call CheckStringJ("EmailMgr_Subject","訊息標題")%>
  <%call ChecklenJ("EmailMgr_Subject",100,"訊息標題")%>
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Eail_OnClick(){
  location.href='emailmgr_send.asp?ser_no='+document.form.ser_no.value+'';
}
function Cancel_OnClick(){
  location.href='emailmgr.asp';
}
--></script>