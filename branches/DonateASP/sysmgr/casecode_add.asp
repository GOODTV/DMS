<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function CaseCodeList (SQL,HLink,LinkParam,LinkTarget,DelLink,DelParam,DelTarget)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
    Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>刪除</span></font></td>"
  Response.Write "</tr>"
  While Not RS1.EOF
    Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
    For I = 1 To FieldsCount
      If I = 1 Then
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & "<a href='#' onclick=""window.open('" & HLink & RS1(LinkParam) & "','','status=no,scrollbars=no,top=100,left=120,width=450,height=260')"">" & RS1(I) & "</a></span></td>"
      Else
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I) & "</span></td>"
      End If
    Next
    Response.Write "<td><a href='JavaScript:if(confirm(""是否確定要刪除 ?"")){window.location.href=""" & DelLink & RS1(DelParam) & """;}' target='" & DelTarget &"'><img src='../images/x5.gif' border=0 width='16' height='14' alt='刪除'></a></td>"
    RS1.MoveNext
    Response.Write "</tr>"
  Wend
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function

If request("action")="save" Then
  SQL="CASECODE"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,3
  RS1.AddNew
  RS1("CodeType")=request("CodeType")
  RS1("CodeName")=request("CodeName")
  RS1("CodeDesc")=request("CodeDesc")
  RS1("CodeKind")=request("CodeKind")
  RS1("Menu_Url")=request("Menu_Url")
  RS1("Menu_Target")=request("Menu_Target")
  RS1("Image_Sub_Url")=request("Image_Sub_Url")
  If request("Image_Sub_Width")<>"" Then
    RS1("Image_Sub_Width")=request("Image_Sub_Width")
  Else
    RS1("Image_Sub_Width")="0"
  End If
  RS1("Seq")=request("Seq")
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="代碼選項新增成功 ！"
End If

If request("action")="delete" Then
  SQL="Select * From CASECODE Where CodeType='"&request("CodeType")&"'"
  Call QuerySQL(SQL,RS)
  If Not RS.EOF Then
    session("errnumber")=1
    session("msg")="『 "&request("codename")&" 』還有代碼選項，不能刪除！"
    response.redirect "codefile.asp?codetype="&session("codetype")
  Else
    SQL="Delete From CODEFILE Where CodeType='"&request("codetype")&"'"
    Set RS=Conn.Execute(SQL)
    session("errnumber")=1
    session("msg")="代碼主檔刪除成功 ！"
    response.redirect "codefile.asp" 
  End If
End If

SQL="Select * From CODEFILE Where CodeType='"&request("codetype")&"'"
Call QuerySQL(SQL,RS)
If Not RS.EOF Then
  session("codetype")=request("codetype")
  CodeType=RS("CodeType")
  CodeName=RS("CodeName")
  Category=RS("Category")
End If

Max_Seq=1
SQL1="Select Max_Seq=Isnull(Max(Seq),0) From CASECODE Where CodeType='"&request("codetype")&"' And Seq<99"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then 
  Max_Seq=Cint(RS1("Max_Seq"))+1
Else
  SQL2="Select Max_Seq=Isnull(Max(Seq),0) From CASECODE Where CodeType='"&request("codetype")&"' And Seq>=99"
  Call QuerySQL(SQL2,RS2)
  If Not RS2.EOF Then Max_Seq=Cint(RS2("Max_Seq"))+1
  RS2.Close
  Set RS2=Nothing 
End If
RS1.Close
Set RS1=Nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>代碼選項列表</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=tool>
  <form name="form" method="post" action="" target="main">
    <input type="hidden" name="action">
    <table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%" align="center">
      <%If CodeType<>"" Then%>
      <tr>
        <td width="14%" align="right">代碼<span lang="en-us">：</span></td>
        <td width="26%"><input type="text" name="codetype" size="14" readonly class="font9t" value="<%=CodeType%>"></td>
        <td width="14%" align="right">名稱<span lang="en-us">：</span></td>
        <td width="26%"><input type="text" name="CodeName" size="14" value="<%=CodeName%>" readonly class="font9t"> </td>
        <td width="20%"><input type="button" value="刪除主檔" name="delete" class="delbutton" onClick="Delete_OnClick()"></td>
      </tr>
      <tr>
        <td width="100%" colspan="6" bgcolor="#C0C0C0" height="1"> </td>
      </tr>
      <tr>
        <td align="right">選項<span lang="en-us">：</span></td>
        <td colspan="3"><input type="text" name="CodeDesc" size="35" class="font9" maxlength="60" onKeypress='if(event.keyCode==13){JavaScript:Save_OnClick();}'></td>
        <td><input type="button" value="新增選項" name="Save" class="addbutton" onClick="Save_OnClick()"></td>
      </tr>
      <%If Session("Menu_Edit")="Y" And RS("Category")<>"類別選項" Then%>
      <tr>
        <td align="right">連結<span lang="en-us">：</span></td>
        <td colspan="3"><input type="text" name="Menu_Url" size="35" class="font9" maxlength="40" <%If Category="單一網頁" Then%>value="content.asp"<%End if%>></td>
        <td> </td>
      </tr>
      <tr>
        <td align="right">視窗<span lang="en-us">：</span></td>
        <td colspan="3">
          <input type="radio" name="Menu_Target" value="_blank">開啟新視窗&nbsp; 
          <input type="radio" name="Menu_Target" value="_self" checked >原視窗
        </td>
        <td> </td>
      </tr>
      <tr>
        <td align="right">圖檔<span lang="en-us">：</span></td>
        <td colspan="3"><input type="text" name="Image_Sub_Url" size="35" class="font9" maxlength="200"></td>
        <td> </td>
      </tr>
      <tr>
        <td align="right">圖寬<span lang="en-us">：</span></td>
        <td colspan="3"><input type="text" name="Image_Sub_Width" size="35" class="font9"></td>
        <td> </td>
      </tr> 
      <%Else%>
      <input type="hidden" name="Menu_Url">
      <input type="hidden" name="Menu_Target">	
      <input type="hidden" name="Image_Sub_Url">
      <input type="hidden" name="Image_Sub_Width" value="0">
      <%End If%>
      <tr>
        <td align="right">排序<span lang="en-us">：</span></td>
        <td colspan="3"><input type="text" name="Seq" size="8" class="font9" maxlength="3" value="<%=Max_Seq%>"></td>
        <td>&nbsp;</td>
      </tr>
      <%
        If CodeType="Purpose" Or CodeType="Accoun" Then
          SQL_IS="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"' "
          Set RS_IS=Server.CreateObject("ADODB.Recordset")
          RS_IS.Open SQL_IS,Conn,1,1
          IsDonation=RS_IS("IsDonation")
          IsMember=RS_IS("IsMember")
          IsShopping=RS_IS("IsShopping")
          RS_IS.Close
          Set RS_IS=Nothing
          If IsMember="Y" Or IsShopping="Y" Then
      %>
      <tr>
        <td align="right">分類<span lang="en-us">：</span></td>
        <td colspan="4">
          <input type="radio" name="CodeKind" value="D" checked >愛心捐款&nbsp;
          <%If IsMember="Y" Then%><input type="radio" name="CodeKind" value="M">會務繳費&nbsp;<%End If%>
          <%If IsMember="Y" Then%><input type="radio" name="CodeKind" value="A">活動報名&nbsp;<%End If%>
          <%If IsShopping="Y" Then%><input type="radio" name="CodeKind" value="S">商品義賣&nbsp;<%End If%>
          <!--<input type="radio" name="CodeKind" value="V">志工-->	
        </td>
      </tr>
      <%
          Else
            Response.Write "<input type='hidden' name='CodeKind' value='D'>"
          End If
        Else
          Response.Write "<input type='hidden' name='CodeKind' value=''>"
        End If
      %>
      <tr>
        <td width="100%" colspan="5">
        <%
          If CodeType="Purpose" Or CodeType="Accoun" Then
            If Session("Menu_Edit")="Y" And Category<>"類別選項" Then
              SQL="Select Ser_No,代碼選項=CodeDesc,分類=(Case When CodeKind='D' Then '愛心捐款' Else Case When CodeKind='M' Then '會務繳費' Else Case When CodeKind='M' Then '會務繳費' Else Case When CodeKind='A' Then '活動報名' Else '商品義賣' End End End End),連結頁面=Menu_Url,排序=Seq From CASECODE Where CodeType='"&request("codetype")&"' Order By Seq"	
            Else
              'SQL="Select Ser_No,代碼選項=CodeDesc,分類=(Case When CodeKind='D' Then '愛心捐款' Else Case When CodeKind='M' Then '會務繳費' Else Case When CodeKind='M' Then '會務繳費' Else Case When CodeKind='A' Then '活動報名' Else '商品義賣' End End End End),排序=Seq From CASECODE Where CodeType='"&request("codetype")&"' Order By Seq"
              SQL="Select Ser_No,代碼選項=CodeDesc,排序=Seq From CASECODE Where CodeType='"&request("codetype")&"' Order By Seq"
            End If
          Else
            If Session("Menu_Edit")="Y" And Category<>"類別選項" Then
              SQL="Select Ser_No,代碼選項=CodeDesc,連結頁面=Menu_Url,排序=Seq From CASECODE Where CodeType='"&request("codetype")&"' Order By Seq"	
            Else
              SQL="Select Ser_No,代碼選項=CodeDesc,排序=Seq From CASECODE Where CodeType='"&request("codetype")&"' Order By Seq"
            End If
          End If
          HLink="casecode_edit.asp?ser_no="
          LinkParam="ser_no"
          LinkTarget="update"
             
          DelLink="casecode_delete.asp?ser_no="
          DelParam="ser_no"
          DelTarget="detail"

          call CaseCodeList (SQL,HLink,LinkParam,LinkTarget,DelLink,DelParam,DelTarget)
        %>
        </td>
      </tr>
      <%End If%>
    </table>
  </form>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Save_OnClick(){
  <%call CheckStringJ("CodeDesc","代碼選項")%>
  <%call ChecklenJ("CodeDesc",60,"代碼選項")%>
  <%call ChecklenJ("Menu_Url",40,"連結頁面")%>
  <%call CheckStringJ("Seq","排序")%>
  <%call CheckNumberJ("Seq","排序")%>
  document.form.target='detail';
  <%call SubmitJ("save")%>
}
function Delete_OnClick(){
  if(confirm('您是否確定要刪除主檔？')){
    document.form.action.value='delete';
    document.form.submit();
  }
}
--></script>