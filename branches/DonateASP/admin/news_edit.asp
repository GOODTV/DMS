<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From NEWS Where Ser_No='"&request("ser_no")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("News_Subject")=request("News_Subject")
  RS("News_Type")=request("News_Type")
  RS("News_ShowHome")=request("News_ShowHome")
  RS("News_ShowSubPage")=request("News_ShowSubPage")
  RS("News_Ad")=request("News_Ad")
  RS("News_ShowMarquee")=request("News_ShowMarquee")
  RS("News_Seq")=request("News_Seq")
  RS("News_Author")=request("News_Author")
  RS("News_Source")=request("News_Source")
  If request("News_Brief")<>"" Then
    RS("News_Brief")=request("News_Brief")
  Else
    RS("News_Brief")=null
  End If
  RS("News_Desc")=request("News_Desc")
  RS("News_Redirect")=request("News_Redirect")
  RS("News_Redirect_url")=request("News_Redirect_url")
  'RS("News_ImgPosition")=request("News_ImgPosition")
  RS("News_ImageShow_Type")=request("News_ImageShow_Type")
  RS("News_BeginDate")=request("News_BeginDate")
  RS("News_EndDate")=request("News_EndDate")
  RS("BranchType")=request("BranchType")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 ！"
  'Response.Redirect "news.asp?code_id="&request("code_id")
End If

If request("action")="delete" Then
  SQL="Delete From NEWS Where Ser_No='"&request("ser_no")&"' " &_
      "Delete From NEXTPAGE Where Ser_No='"&request("ser_no")&"' And Page_Type='news'"
  Set RS=Conn.Execute(SQL)
  SQL1="Select * From UPLOAD Where object_id='"&request("ser_no")&"' And (Ap_Name='news' Or Ap_Name='news_title')"
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
  Response.Redirect "news.asp?code_id="&request("code_id")
End If

If request("action")="addpage" Then
  SQL="Update NEWS Set News_IsPage='Y' Where ser_no='"&request("ser_no")&"'"
  Set RS=Conn.Execute(SQL)
  
  SQL="Select * From NEXTPAGE Where Page_type='news' And ser_no='"&request("ser_no")&"' Order By Page_No Desc"
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
  RS("Page_Type")="news"
  RS("Page_No")=page_no
  RS("Page_Desc")=""
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="新增分頁成功 ！"  
End If

News_Type="訊息內容"
If request("code_id")<>"" Then
  SQL="Select News_Type=CodeDesc From CASECODE Where CodeType='NewsType' And Seq='"&request("code_id")&"' " 
  call QuerySQL(SQL,RS)
  If Not RS.EOF Then News_Type=RS("News_Type")
End If 

SQL="Select * From NEWS Where Ser_No='"&request("ser_no")&"'"
call QuerySQL(SQL,RS)
If request("newstype")<>"" Then
  newstype=request("newstype")
Else
  newstype=RS("News_Type")
End If
Item="news"
Subject=RS("News_Subject")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=News_Type%>管理 【修改資料】</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
      <input type="hidden" name="code_id" value="<%=request("code_id")%>">	    
      <table border="0" width="815" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
	        <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25" colspan="2"><font color="#FFFFFF"><b>&nbsp;<%=News_Type%>管理 </b> <font size="2">【修改資料】</font></font></td>
        </tr>
        <tr>
          <td width="20%" class="font11" align="center" valign="top" bgcolor="EEEEE3">
            <iframe name="left1" src="newstype_list.asp?newstype=<%=newstype%>&code_id=<%=request("code_id")%>" height="240" width="100%" frameborder="1" scrolling="auto" target="_self"></iframe><br>
            <iframe name="left2" src="image_left_list.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=request("ser_no")%>&code_id=<%=request("code_id")%>" height="320" width="100%" frameborder="1" scrolling="auto" target="_self"></iframe>
          </td>
	        <td width="80%" class="font11" valign="top">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">
   	            <tr>
   		            <td>
                    <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" bordercolor="#E1DFCE" cellpadding="2">
                      <tr>
                        <td noWrap align="right" width="13%"><font color="#000080">訊息類別：</font></td>
                        <td width="15%">
                        <%
                          If News_Type="訊息內容" Then
                            SQL="Select News_Type=CodeDesc From CASECODE Where CodeType='NewsType' Order By Seq"
                          Else
                            SQL="Select News_Type=CodeDesc From CASECODE Where CodeType='NewsType' And CodeDesc='"&News_Type&"' Order By Seq"
                          End If
                          FName="News_Type"
                          Listfield="News_Type"
                          menusize="1"
                          BoundColumn=RS("News_Type")
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        </td>
                        <td width="10%" noWrap align="right"><font color="#000080">首頁顯示：</font></td>
                        <td width="14%">
                          <input type="radio" name="News_ShowHome" value="Y" <%If RS("News_ShowHome")="Y" Then%>checked<%End If%> >是&nbsp; 
                          <input type="radio" name="News_ShowHome" value="N" <%If RS("News_ShowHome")="N" Then%>checked<%End If%> >否&nbsp;
                        </td>
                        <td width="10%" noWrap align="right"><font color="#000080">次頁顯示：</font></td>
                        <td width="14%">
                          <input type="radio" name="News_ShowSubPage" value="Y" <%If RS("News_ShowSubPage")="Y" Then%>checked<%End If%> >是&nbsp; 
                          <input type="radio" name="News_ShowSubPage" value="N" <%If RS("News_ShowSubPage")="N" Then%>checked<%End If%> >否
                        </td>
                        <td width="10%" noWrap align="right"><font color="#000080">頭條新聞：</font></td>
                        <td width="14%">
                          <input type="radio" name="News_Ad" value="Y" <%If RS("News_Ad")="Y" Then%>checked<%End If%> >是&nbsp; 
                          <input type="radio" name="News_Ad" value="N" <%If RS("News_Ad")="N" Then%>checked<%End If%> >否
                        </td>
                      </tr>
                      <tr>
                        <td noWrap align="right"><font color="#000080">訊息標題：</font></td>
                        <td colspan="3"><input class=font9 size=39 name="News_Subject" maxlength="100" value="<%=RS("News_Subject")%>"></td>
                        <td noWrap align="right"><font color="#000080">跑馬燈：</font></td>
                        <td>
                          <input type="radio" name="News_ShowMarquee" value="Y" <%If RS("News_ShowMarquee")="Y" Then%>checked<%End If%> >是&nbsp; 
                          <input type="radio" name="News_ShowMarquee" value="N" <%If RS("News_ShowMarquee")="N" Then%>checked<%End If%> >否
                        </td>
                        <td noWrap align="right"><font color="#000080">所屬分會：</font></td>
                        <td>
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
                        <td noWrap align="right"><font color="#000080">標題圖檔：</font></td>
                        <td valign="center" colspan="7">
                          <%If RS("News_TitleImg")<>"" Then%>
                            <img border="0" src="../upload/<%=RS("News_TitleImg")%>">
                            <input type="button" value=" 刪 除 " name="title_delete" class="delbutton" style="cursor:hand" OnClick='if(confirm("是否確定要刪除標題圖檔 ?")){window.location.href="image_title_delete.asp?item=<%=Item%>&ser_no=<%=RS("ser_no")%>&code_id=<%=request("code_id")%>";}'>
                          <%Else%>
                            <input type="button" value=" 上 傳 " name="title_upload" class="addbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_title_upload.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&img_width=&img_height=&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=110')">
                          <%End If%>
                        </td>
                      </tr>
                      <tr> 
                        <td noWrap align="right"><font color="#000080">刊登日期：</font></td>
                        <td valign="center" colspan="3"><%call Calendar("News_BeginDate",RS("News_BeginDate"))%> ~ <%call Calendar("News_EndDate",RS("News_EndDate"))%></td>
                        <td noWrap align="right"><font color="#000080">作者：</font></td>
                        <td><input class=font9 size="11" name="News_Author" maxlength="5" value="<%=RS("News_Author")%>"></td>
                        <td align="right"><font color="#000080">排序：</font></td>
                        <td><input class=font9 size="11" name="News_Seq" maxlength="5" value="<%=RS("News_Seq")%>"></td>
                      </tr>
                      <!--#include file="../include/calendar2.asp"-->                        
                      <tr> 
                        <td align="right"><font color="#000080"><span lang="zh-tw">訊息摘要</span>：</font></td>
                        <td colspan="7"><textarea rows="4" name="News_Brief" cols="80" class="font9"><%=RS("News_Brief")%></textarea></td>
                      </tr>
                      <tr>
                        <td align="right"><font color="#000080">是否轉址：</font></td>
                        <td colspan="7"> 
                          <input type="radio" name="News_Redirect" id="redirectYes" value="Y" <%If RS("News_Redirect")="Y" Then%>checked<%End If%>>是<font color="#000080">：
                          <input type="text" name="News_Redirect_Url" size="60" maxlength="120" class="font9" value="<%=RS("News_Redirect_Url")%>" OnFocus="Redirect_OnFocus()"><br>
                          <input type="radio" name="News_Redirect" id="redirectNo" value="N" <%If RS("News_Redirect")="N" Then%>checked<%End If%>></font>否：
                          請填寫以下之訊息內容
                        </td>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2" colspan="7"></td>
                      </tr>
                      <tr>
                        <td align="right"><font color="#000080">訊息內容：<p align=center><input type="button" value="分頁" name="AddPage" class="addbutton" style="cursor:hand" onClick="AddPage_OnClick()"></font></td>
                        <td valign="center" colspan="7">
                        <!--#include file="../fckeditor/fckeditor.asp" -->
                        <%
                          Dim oFCKeditor
			                    Set oFCKeditor = New FCKeditor
		                      oFCKeditor.BasePath	= "../fckeditor/"
			                    oFCKeditor.Value	= RS("News_Desc")
		                      oFCKeditor.Create "News_Desc"
		                    %>                        
                        </td>
                      </tr>
                      <!--#include file="nextpage_show.asp"-->
                      <tr>
                        <td width="100%" colspan="8"><%EditButton%></td>
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
function Redirect_OnFocus(){
  document.form.News_Redirect[0].checked=true;
}
function AddPage_OnClick(){
  <%call SubmitJ("addpage")%>
}
function Update_OnClick(){
  <%call CheckStringJ("News_Type","訊息類別")%>
  <%call CheckStringJ("News_Subject","訊息標題")%>
  <%call ChecklenJ("News_Subject",100,"訊息標題")%>
  <%call ChecklenJ("News_Author",20,"作者")%>
  <%call CheckStringJ("News_Seq","排序")%>
  <%call CheckNumberJ("News_Seq","排序")%>
  <%call CheckStringJ("News_BeginDate","刊登起日")%>
  if(document.form.News_EndDate.value==''){
    document.form.News_EndDate.value='2099/12/31';
  }
  <%call DiffDateJ("News_BeginDate","News_EndDate")%>
  if(document.form.News_Redirect[0].checked){
    <%call CheckStringJ("News_Redirect_Url","轉址")%>
  }else{
    document.form.News_Redirect_Url.value='';
  }
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  location.href='news.asp?code_id='+document.form.code_id.value+'';
}
--></script>