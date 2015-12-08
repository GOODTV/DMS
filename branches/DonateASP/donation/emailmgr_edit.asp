<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From EMAILMGR Where Ser_No='"&request("ser_no")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("EmailMgr_Subject")=request("EmailMgr_Subject")
  RS("EmailMgr_Type")=request("EmailMgr_Type")
  RS("EmailMgr_Desc")=request("EmailMgr_Desc")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 ！"
  'Response.Redirect "emailmgr.asp"
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
If request("emailmgrtype")<>"" Then
  emailmgrtype=request("emailmgrtype")
Else
  emailmgrtype=RS("EmailMgr_Type")
End If
Item="emailmgr"
Subject=RS("EmailMgr_Subject")
%>
<%Prog_Id="emailmgr"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">    
      <table border="0" width="815" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
	        <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25" colspan="2"><font color="#FFFFFF">&nbsp;&nbsp;<%=Prog_Desc%>【修改】</font></td>
        </tr>
        <tr>
          <td width="20%" class="font11" align="center" valign="top" bgcolor="EEEEE3">
            <iframe name="left1" src="emailmgrtype_list.asp?emailmgrtype=<%=emailmgrtype%>" height="240" width="100%" frameborder="1" scrolling="auto" target="_self"></iframe><br>
            <iframe name="left2" src="image_left_list.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=request("ser_no")%>" height="320" width="100%" frameborder="1" scrolling="auto" target="_self"></iframe>
          </td>
	        <td width="80%" class="font11" valign="top">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">
   	            <tr>
   		            <td>
                    <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" bordercolor="#E1DFCE" cellpadding="2">
                      <tr>
                        <td noWrap align="right" width="13%"><font color="#000080">郵件類別：</font></td>
                        <td width="22%">
                        <%
                          SQL="Select EmailMgr_Type=CodeDesc From CASECODE Where CodeType='EmailMgrType' Order By Seq"
                          FName="EmailMgr_Type"
                          Listfield="EmailMgr_Type"
                          menusize="1"
                          BoundColumn=RS("EmailMgr_Type")
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        </td>
                        <td width="13%" noWrap align="right"><font color="#000080">郵件標題：</font></td>
                        <td width="52%"><input class=font9 size=52 name="EmailMgr_Subject" maxlength="100" value="<%=RS("EmailMgr_Subject")%>"></td>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2" colspan="7"></td>
                      </tr>
                      <tr> 
                        <td align="right"><font color="#000080">郵件內容：</font></td>
                        <td valign="center" colspan="3" height="200">
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
                        <td width="100%" colspan="4"><%EditButton%></td>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2" colspan="7"></td>
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
function Window_OnLoad(){
  document.form.EmailMgr_Type.focus();
}	
function Update_OnClick(){
  <%call CheckStringJ("EmailMgr_Type","郵件類別")%>
  <%call CheckStringJ("EmailMgr_Subject","郵件標題")%>
  <%call ChecklenJ("EmailMgr_Subject",100,"郵件標題")%>
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  location.href='emailmgr.asp';
}
--></script>