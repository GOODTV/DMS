<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  SQL="EMAILMGR"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS.Addnew
  RS("EmailMgr_Subject")=request("EmailMgr_Subject")
  RS("EmailMgr_Type")=request("EmailMgr_Type")
  RS("EmailMgr_Desc")=request("EmailMgr_Desc")
  RS("EmailMgr_ImgPosition")="right"
  RS("EmailMgr_ImageShow_Type")="barrier"
  RS("EmailMgr_RegDate")=date()
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
<%Prog_Id="emailmgr"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <table border="0" width="830" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
	        <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25" colspan="2"><font color="#FFFFFF">&nbsp;&nbsp;<%=Prog_Desc%>【新增】</font></td>
        </tr>
        <tr>
	        <td width="20%" class="font11" align="center" valign="top" bgcolor="EEEEE3"><iframe name="left" src="emailmgrtype_list.asp" height="220" width="100%" frameborder="1" scrolling="auto" target="_self"></iframe></td>
	        <td width="80%" class="font11" valign="top">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">
   		          <tr>
   		            <td>
                    <table width="100%" border="1" cellspacing="0" style="border-collapse: collapse" cellpadding="2" bordercolor="#E1DFCE">
                      <tr>
                        <td noWrap align="right" width="13%"><font color="#000080">郵件類別：</font></td>
                        <td width="22%">
                        <%
                          SQL="Select EmailMgr_Type=CodeDesc From CASECODE Where CodeType='EmailMgrType' Order By Seq"
                          FName="EmailMgr_Type"
                          Listfield="EmailMgr_Type"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        </td>
                        <td width="13%" noWrap align="right"><font color="#000080">郵件標題：</font></td>
                        <td width="52%"><input class=font9 size=52 name="EmailMgr_Subject" maxlength="100"></td>
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
			                    oFCKeditor.Value	= ""
		                      oFCKeditor.Create "EmailMgr_Desc"
		                    %>
                        </td>
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
function Window_OnLoad(){
  document.form.EmailMgr_Type.focus();
}	
function Redirect_OnFocus(){
  document.form.News_Redirect[0].checked=true;
}
function Save_OnClick(){
  <%call CheckStringJ("EmailMgr_Type","郵件類別")%>
  <%call CheckStringJ("EmailMgr_Subject","郵件標題")%>
  <%call ChecklenJ("EmailMgr_Subject",100,"郵件標題")%>
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  location.href='emailmgr.asp';
}
--></script>