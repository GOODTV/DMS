<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  SQL="NEWS"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS.Addnew
  RS("News_Subject")=request("News_Subject")
  RS("News_Type")=request("News_Type")
  RS("News_ShowHome")=request("News_ShowHome")
  RS("News_ShowSubPage")=request("News_ShowSubPage")
  RS("News_Ad")=request("News_Ad")
  RS("News_ShowMarquee")=request("News_ShowMarquee")
  RS("News_Seq")=request("News_Seq")
  RS("News_Author")=request("News_Author")
  RS("News_Source")=request("News_Source")
  RS("News_TitleImg")=""
  If request("News_Brief")<>"" Then
    RS("News_Brief")=request("News_Brief")
  Else
    RS("News_Brief")=null
  End If
  RS("News_Desc")=request("News_Desc")
  RS("News_Redirect")=request("News_Redirect")
  RS("News_Redirect_url")=request("News_Redirect_url")
  RS("News_ImgPosition")="right"
  RS("News_ImageShow_Type")="barrier"
  RS("News_RegDate")=date()
  RS("News_BeginDate")=request("News_BeginDate")
  RS("News_EndDate")=request("News_EndDate")
  RS("News_IsPage")="N"
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
    response.redirect "news_edit.asp?code_id="&request("code_id")&"&ser_no="&RS("ser_no")
  End If   
End If

News_Type="訊息內容"
If request("code_id")<>"" Then
  SQL="Select News_Type=CodeDesc From CASECODE Where CodeType='NewsType' And Seq='"&request("code_id")&"' " 
  call QuerySQL(SQL,RS)
  If Not RS.EOF Then News_Type=RS("News_Type")
End If   
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=News_Type%>管理 【新增資料】</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="code_id" value="<%=request("code_id")%>">	
      <table border="0" width="830" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
	        <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25" colspan="2"><font color="#FFFFFF"><b>&nbsp;&nbsp;<%=News_Type%>管理 </b> <font size="2">【新增資料】</font></font></td>
        </tr>
        <tr>
	        <td width="20%" class="font11" align="center" valign="top" bgcolor="EEEEE3"><iframe name="left" src="newstype_list.asp?code_id=<%=request("code_id")%>" height="220" width="100%" frameborder="1" scrolling="auto" target="_self"></iframe></td>
	        <td width="80%" class="font11" valign="top">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">
   		          <tr>
   		            <td>
                    <table width="100%" border="1" cellspacing="0" style="border-collapse: collapse" cellpadding="2" bordercolor="#E1DFCE">
                      <tr>
                        <td noWrap align="right" width="13%"><font color="#000080">訊息類別：</font></td>
                        <td width="15%">
                        <%
                          If News_Type="訊息內容" Then
                            SQL="Select News_Type=CodeDesc From CASECODE Where CodeType='NewsType' Order By Seq"
                            BoundColumn=""
                          Else
                            SQL="Select News_Type=CodeDesc From CASECODE Where CodeType='NewsType' And CodeDesc='"&News_Type&"' Order By Seq"
                            BoundColumn=News_Type
                          End If
                          FName="News_Type"
                          Listfield="News_Type"
                          menusize="1"
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        </td>
                        <td width="10%" noWrap align="right"><font color="#000080">首頁顯示：</font></td>
                        <td width="14%">
                          <input type="radio" name="News_ShowHome" value="Y" checked >是&nbsp; 
                          <input type="radio" name="News_ShowHome" value="N">否&nbsp;
                        </td>
                        <td width="10%" noWrap align="right"><font color="#000080">次頁顯示：</font></td>
                        <td width="14%">
                          <input type="radio" name="News_ShowSubPage" value="Y" checked >是&nbsp; 
                          <input type="radio" name="News_ShowSubPage" value="N">否
                        </td>
                        <td width="10%" noWrap align="right"><font color="#000080">頭條新聞：</font></td>
                        <td width="14%">
                          <input type="radio" name="News_Ad" value="Y">是&nbsp; 
                          <input type="radio" name="News_Ad" value="N" checked >否
                        </td>
                      </tr>
                      <tr>
                        <td noWrap align="right"><font color="#000080">訊息標題：</font></td>
                        <td colspan="3"><input class=font9 size=39 name="News_Subject" maxlength="100"></td>
                        <td noWrap align="right"><font color="#000080">跑馬燈：</font></td>
                        <td>
                          <input type="radio" name="News_ShowMarquee" value="Y">是&nbsp; 
                          <input type="radio" name="News_ShowMarquee" value="N" checked >否
                        </td>
                        <td noWrap align="right"><font color="#000080">所屬分會：</font></td>
                        <td>
                        <%SQL="Select BranchType=CodeDesc From CASECODE Where CodeType='BranchType' Order By Seq"
                          FName="BranchType"
                          Listfield="BranchType"
                          menusize="1"
                          BoundColumn=Session("BranchType")
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        </td>
                      </tr>
                      <tr> 
                        <td noWrap align="right"><font color="#000080">刊登日期：</font></td>
                        <td valign="center" colspan="3"><%call Calendar("News_BeginDate",date())%> ~ <%call Calendar("News_EndDate","")%></td>
                        <td noWrap align="right"><font color="#000080">作者：</font></td>
                        <td><input class=font9 size="11" name="News_Author" maxlength="5" value="<%=session("user_name")%>"></td>
                        <td align="right"><font color="#000080">排序：</font></td>
                        <td><input class=font9 size="11" name="News_Seq" maxlength="5" value="999"></td>
                      </tr>
                      <!--#include file="../include/calendar2.asp"-->
                      <tr> 
                        <td align="right"><font color="#000080"><span lang="zh-tw">訊息摘要</span>：</font></td>
                        <td colspan="7"><textarea rows="6" name="News_Brief" cols="91" class="font9"></textarea></td>
                      </tr>
                      <tr>
                        <td align="right"><font color="#000080">是否轉址：</font></td>
                        <td colspan="7"> 
                          <input type="radio" name="News_Redirect" id="redirectYes" value="Y">是<font color="#000080">：
                          <input type="text" name="News_Redirect_Url" size="60" maxlength="120" class="font9" OnFocus="Redirect_OnFocus()"><br>
                          <input type="radio" name="News_Redirect" id="redirectNo" value="N" checked ></font>否：
                          請填寫以下之訊息內容
                        </td>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2" colspan="7"></td>
                      </tr>
                      <tr> 
                        <td align="right"><font color="#000080">訊息內容：</font></td>
                        <td valign="center" colspan="7" height="200">
                        <!--#include file="../fckeditor/fckeditor.asp" -->
                        <%
                          Dim oFCKeditor
			                    Set oFCKeditor = New FCKeditor
		                      oFCKeditor.BasePath	= "../fckeditor/"
			                    oFCKeditor.Value	= ""
		                      oFCKeditor.Create "News_Desc"
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
function Redirect_OnFocus(){
  document.form.News_Redirect[0].checked=true;
}
function Save_OnClick(){
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
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  location.href='news.asp?code_id='+document.form.code_id.value+'';
}
--></script>