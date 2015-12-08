<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" then
  SQL1="Select * From CONTENT Where Content_Type='"&request("Content_Type")&"' And Content_SubType='"&request("Content_SubType")&"' And Content_Subject='"&request("Content_Subject")&"'"
  call QuerySQL(SQL1,RS1)
  If RS1.EOF Then
    SQL="CONTENT"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Content_Subject")=request("Content_Subject")
    RS("Content_Type")=request("Content_Type")
    RS("Content_SubType")=request("Content_SubType")
    If request("Content_Seq")<>"" Then
      RS("Content_Seq")=request("Content_Seq")
    Else
      RS("Content_Seq")="1"  
    End If
    RS("Content_Author")=request("Content_Author")
    RS("Content_TitleImg")=""
    If request("Content_Brief")<>"" Then
      RS("Content_Brief")=request("Content_Brief")
    Else
      RS("Content_Brief")=null
    End If
    RS("Content_Desc")=request("Content_Desc")
    RS("Content_Redirect")=request("Content_Redirect")
    RS("Content_Redirect_Url")=request("Content_Redirect_Url")
    RS("Content_ImgPosition")="right"
    RS("Content_ImageShow_Type")="barrier"
    RS("Content_RegDate")=date()
    RS("Content_IsPage")="N"
    RS("BranchType")=request("BranchType")
    RS.Update
    RS.Close
    Set RS=Nothing
    session("errnumber")=1
    session("msg")="資料新增成功！"
    SQL="Select @@IDENTITY as ser_no"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,1
    If Not RS.EOF Then
      response.redirect "content_edit.asp?ser_no="&RS("ser_no")
    End If
  Else
    session("errnumber")=1
    session("msg")="您新增的資料已經存在，請確認？(主功能+頁面+標題)"
  End If
  RS1.Close
  Set RS1=Nothing    
End If  
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>網頁內容管理 【新增資料】</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <table border="0" width="830" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
	        <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25" colspan="2"><font color="#FFFFFF"><b>&nbsp;&nbsp;網頁內容管理 </b> <font size="2">【新增資料】</font></font></td>
        </tr>
        <tr>
	        <td width="20%" class="font11" align="center" valign="top" bgcolor="EEEEE3"><iframe name="left" src="contenttype_list.asp" height="220" width="100%" frameborder="1" scrolling="auto" target="_self"></iframe></td>
	        <td width="80%" class="font11" valign="top">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">
   		          <tr>
   		            <td>
                    <table width="100%" border="1" cellspacing="0" style="border-collapse: collapse" cellpadding="2" bordercolor="#E1DFCE">
                      <tr> 
                        <td noWrap align="right" width="12%"><font color="#000080">主功能：</font></td>
                        <td width="16%"> 
                        <%
                          SQL="Select Content_Type=CodeName From CodeFile Where Category='單一網頁' Order By Menu_Seq"
                          call Sel_CodeType (SQL,"Content_Type",request("Content_Type"),"Content_SubType")
                        %>
                        </td>
                        <td width="8%" align="right"><font color="#000080">頁面：</font></td>
                        <td width="30%"><%call Sel_SubType(request("Content_Type"),"Content_SubType",request("Content_SubType"))%></td>
                        <td width="12%" align="right"><font color="#000080">所屬分會：</font></td>
                        <td width="22%">
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
                        <td noWrap align="right"><font color="#000080">網頁標題：</font></td>
                        <td valign="center" colspan="3">
                        	<input class=font9 size=42 name="Content_Subject" maxlength="80" value="<%=request("Content_Subject")%>">&nbsp;&nbsp;&nbsp;&nbsp;
                        	<font color="#000080">排序：</font>
                        	<input class=font9 size=3 name="Content_Seq" maxlength="3" value="<%=request("Content_Seq")%>">	
                        </td>
                        <td noWrap align="right"><font color="#000080">作者：</font></td>
                        <td><input class=font9 size="19" name="Content_Author" maxlength="20" value="<%=session("user_name")%>"></td>
                      </tr>
                      <!--<tr> 
                        <td align="right"><font color="#000080"><span lang="zh-tw">內容摘要</span>：</font></td>
                        <td colspan="5"><textarea rows="6" name="Content_Brief" cols="91" class="font9"><%=request("Content_Brief")%></textarea></td>
                      </tr>-->
                      <tr>
                        <td align="right"><font color="#000080">是否轉址：</font></td>
                        <td colspan="5"> 
                          <input type="radio" name="Content_Redirect" id="redirectYes" value="Y" <%If request("Content_Redirect")="Y" Then%>checked<%End If%> >是<font color="#000080">：
                          <input type="text" name="Content_Redirect_Url" size="60" maxlength="120" class="font9" value="<%=request("Content_Redirect_Url")%>" OnFocus="Redirect_OnFocus()"><br>
                          <input type="radio" name="Content_Redirect" id="redirectNo" value="N" <%If request("Content_Redirect")<>"Y" Then%>checked<%End If%> ></font>否：
                          請填寫以下之訊息內容
                        </td>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2" colspan="5"></td>
                      </tr>
                      <tr> 
                        <td align="right"><font color="#000080">網頁內容：</font></td>
                        <td valign="center" colspan="5" height="200">
                        <!--#include file="../fckeditor/fckeditor.asp" -->
                        <%
                          Dim oFCKeditor
			                    Set oFCKeditor = New FCKeditor
		                      oFCKeditor.BasePath	= "../fckeditor/"
			                    oFCKeditor.Value	= request("Content_Desc")
		                      oFCKeditor.Create "Content_Desc"
		                    %>
                        </td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 class="font9" colspan="6"><%SaveButton%></td>
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
function Save_OnClick(){
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
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  location.href='content.asp';
}
--></script>