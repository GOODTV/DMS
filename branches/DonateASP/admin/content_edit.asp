<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From CONTENT Where Ser_No='"&request("ser_no")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("Content_Subject")=request("Content_Subject")
  RS("Content_Type")=request("Content_Type")
  RS("Content_SubType")=request("Content_SubType")
  If request("Content_Seq")<>"" Then
    RS("Content_Seq")=request("Content_Seq")
  Else
    RS("Content_Seq")="1"  
  End If  
  RS("Content_Author")=request("Content_Author")
  If request("Content_Brief")<>"" Then
    RS("Content_Brief")=request("Content_Brief")
  Else
    RS("Content_Brief")=null
  End If
  RS("Content_Desc")=request("Content_Desc")
  RS("Content_Redirect")=request("Content_Redirect")
  RS("Content_Redirect_Url")=request("Content_Redirect_Url")
  'RS("Content_ImgPosition")=request("Content_ImgPosition")
  RS("Content_ImageShow_Type")=request("Content_ImageShow_Type")
  RS("BranchType")=request("BranchType")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 ！"
  'Response.Redirect "content.asp"
End If

If request("action")="delete" Then
  SQL="Delete From CONTENT Where Ser_No='"&request("ser_no")&"' " &_
      "Delete From NEXTPAGE Where Ser_No='"&request("ser_no")&"' And Page_Type='content'"
  Set RS=Conn.Execute(SQL)
  SQL1="Select * From UPLOAD Where object_id='"&request("ser_no")&"' And (Ap_Name='content' Or Ap_Name='content_title')"
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
  Response.Redirect "content.asp"
End If

If request("action")="addpage" Then
  SQL="Update CONTENT Set Content_IsPage='Y' Where ser_no='"&request("ser_no")&"'"
  Set RS=Conn.Execute(SQL)
  
  SQL="Select * From NEXTPAGE Where Page_type='content' And ser_no='"&request("ser_no")&"' Order By Page_No Desc"
  Call QuerySQL(SQL,RS)
  If Not RS.EOF then
    page_no=Cint(RS("page_no"))+1
  Else
    page_no=2
  End If
  RS.Close
  Set RS=Nothing
    
  SQL="NEXTPAGE"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS.Addnew
  RS("Ser_No")=request("Ser_No")
  RS("Page_Type")="content"
  RS("Page_No")=page_no
  RS("Page_Desc")=""
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="新增分頁成功 ！"
End If

SQL="Select * From CONTENT Where Ser_No='"&request("ser_no")&"'"
call QuerySQL(SQL,RS)
If request("contenttype")<>"" Then
  contenttype=request("contenttype")
Else
  contenttype=RS("Content_Type")
End If
Item="content"
Subject=RS("Content_Subject")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>網站內容管理 【修改資料】</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
      <table border="0" width="815" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
	        <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25" colspan="2"><font color="#FFFFFF"><b>&nbsp;網頁內容管理 </b> <font size="2">【修改資料】</font></font></td>
        </tr>
        <tr>
          <td width="20%" class="font11" align="center" valign="top" bgcolor="EEEEE3">
            <iframe name="left1" src="contenttype_list.asp?contenttype=<%=contenttype%>" height="240" width="100%" frameborder="1" scrolling="auto" target="_self"></iframe><br>
            <iframe name="left2" src="image_left_list.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=request("ser_no")%>" height="320" width="100%" frameborder="1" scrolling="auto" target="_self"></iframe>
          </td>
	        <td width="80%" class="font11" valign="top">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">
   	            <tr>
   		            <td>
                    <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" bordercolor="#E1DFCE" cellpadding="2">
                      <tr> 
                        <td noWrap align="right" width="12%"><font color="#000080">主功能：</font></td>
                        <td width="16%"> 
                        <%
                          SQL="Select Content_Type=CodeName From CodeFile Where Category='單一網頁' Order By Menu_Seq"
                          call Sel_CodeType (SQL,"Content_Type",RS("Content_Type"),"Content_SubType")
                        %>
                        </td>
                        <td width="8%" align="right"><font color="#000080">頁面：</font></td>
                        <td width="30%"><%call Sel_SubType(RS("Content_Type"),"Content_SubType",RS("Content_SubType"))%></td>
                        <td width="12%" align="right"><font color="#000080">所屬分會：</font></td>
                        <td width="22%">
                        <%SQL="Select BranchType=CodeDesc From CASECODE Where CodeType='BranchType' Order By Seq"
                          FName="BranchType"
                          Listfield="BranchType"
                          menusize="1"
                          BoundColumn=RS("BranchType")
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>	
                        </td>
                      </tr>
                      <tr> 
                        <td noWrap align="right"><font color="#000080">標題：</font></td>
                        <td valign="center" colspan="3">
                        	<input class=font9 size=41 name="Content_Subject" maxlength="80" value="<%=RS("Content_Subject")%>">&nbsp;&nbsp;&nbsp;&nbsp;
                          <font color="#000080">排序：</font>
                        	<input class=font9 size=3 name="Content_Seq" maxlength="3" value="<%=RS("Content_Seq")%>">	
                        </td>
                        <td noWrap align="right"><font color="#000080">作者：</font></td>
                        <td><input class=font9 size="19" name="Content_Author" maxlength="20" value="<%=RS("Content_Author")%>"></td>
                      </tr>
                      <!--<tr>
                        <td noWrap align="right"><font color="#000080">標題圖檔：</font></td>
                        <td valign="center" colspan="5">
                          <%If RS("Content_TitleImg")<>"" Then%>
                            <img border="0" src="../upload/<%=RS("Content_TitleImg")%>">
                            <input type="button" value=" 刪 除 " name="title_delete" class="delbutton" style="cursor:hand" OnClick='if(confirm("是否確定要刪除標題圖檔 ?")){window.location.href="image_title_delete.asp?item=<%=Item%>&ser_no=<%=RS("ser_no")%>";}'>
                          <%Else%>
                            <input type="button" value=" 上 傳 " name="title_upload" class="addbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_title_upload.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&img_width=&img_height=','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=110')">
                          <%End If%>
                        </td>
                      </tr>
                      <tr> 
                        <td align="right"><font color="#000080"><span lang="zh-tw">內容摘要</span>：</font></td>
                        <td colspan="5"><textarea rows="5" name="Content_Brief" cols="80" class="font9"><%=RS("Content_Brief")%></textarea></td>
                      </tr>-->
                      <tr>
                        <td align="right"><font color="#000080">是否轉址：</font></td>
                        <td colspan="5"> 
                          <input type="radio" name="Content_Redirect" id="redirectYes" value="Y" <%If RS("Content_Redirect")="Y" Then%>checked<%End If%>>是<font color="#000080">：
                          <input type="text" name="Content_Redirect_Url" size="60" maxlength="120" class="font9" value="<%=RS("Content_Redirect_Url")%>" OnFocus="Redirect_OnFocus()"><br>
                          <input type="radio" name="Content_Redirect" id="redirectNo" value="N" <%If RS("Content_Redirect")="N" Then%>checked<%End If%>></font>否：
                          請填寫以下之訊息內容
                        </td>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2" colspan="5"></td>
                      </tr>
                      <tr>
                        <td align="right"><font color="#000080">網頁內容：<p align=center><input type="button" value="分頁" name="AddPage" class="addbutton" style="cursor:hand" onClick="AddPage_OnClick()"></font></td>
                        <td valign="center" colspan="5">
                        <!--#include file="../fckeditor/fckeditor.asp" -->
                        <%
                          Dim oFCKeditor
			                    Set oFCKeditor = New FCKeditor
		                      oFCKeditor.BasePath	= "../fckeditor/"
			                    oFCKeditor.Value	= RS("Content_Desc")
		                      oFCKeditor.Create "Content_Desc"
		                    %>
                        </td>
                      </tr>
                      <!--#include file="nextpage_show.asp"-->
                      <tr>
                        <%If Session("user_id")="npois" Then%>
                        <td width="100%" colspan="6"><%EditButton%></td>
                        <%Else%>
                        <td width="100%" colspan="6"><button id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Update_OnClick()'> <img src='../images/update.gIf' width='20' height='20' align='absmiddle'> 修改</button></td>
                        <%End If%>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2" colspan="5"></td>
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
function Chg_SubType(){
  if(document.form.Content_Type.value!=''&&document.form.Content_SubType.value!=''){
    document.form.Content_Subject.value=document.form.Content_Type.value+' / '+document.form.Content_SubType.value; 
  }else{
    document.form.Content_Subject.value='';
  }
}
function Redirect_OnFocus(){
  document.form.Content_Redirect[0].checked=true;
}
function AddPage_OnClick(){
  <%call SubmitJ("addpage")%>
}
function Update_OnClick(){
  <%call CheckStringJ("Content_Type","主功能")%>
  <%call CheckStringJ("Content_SubType","頁面")%>
  <%call CheckStringJ("Content_Subject","網頁標題")%>
  <%call ChecklenJ("Content_Subject",100,"網頁標題")%>
  <%call ChecklenJ("Content_Author",20,"作者")%>
  if(document.form.Content_Redirect[0].checked){
    <%call CheckStringJ("Content_Redirect_Url","轉址")%>
  }else{
    document.form.Content_Redirect_Url.value='';
  }
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  location.href='content.asp';
}
--></script>