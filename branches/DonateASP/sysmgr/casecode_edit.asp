<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From CodeFile Where CodeType='"&request("CodeType")&"'"
  Call QuerySQL(SQL,RS)
  CodeTable=RS("CodeTable")
  CodeFile=RS("CodeFile")
  CodeWhere=RS("CodeWhere")
  RS.close
  Set RS=Nothing
 
  SQL="Select * From CASECODE Where ser_no='"&request("ser_no")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("CodeDesc")=request("CodeDesc")
  RS("Menu_Url")=request("Menu_Url")
  RS("Menu_Target")=request("Menu_Target")
  RS("Image_Sub_Url")=request("Image_Sub_Url")
  If request("Image_Sub_Width")<>"" Then
    RS("Image_Sub_Width")=request("Image_Sub_Width")
  Else
    RS("Image_Sub_Width")=null
  End If
  RS("Seq")=request("Seq")
  RS.Update
  RS.Close
  Set RS=Nothing
    
  If CodeTable<>"" Then
    All_Table = Split(CodeTable,"/")
    All_File = Split(CodeFile,"/")
    If CodeWhere<>"" Then 
      All_Where = Split(CodeWhere,"/")
    End If
    For I = 0 to UBound(All_Table)
      SQL=SQL&" Update " & All_Table(I) & " Set " & All_File(I) & "='"&request("CodeDesc")&"' Where " & All_File(I) & "='"&request("CodeDesc_Old")&"' "
      If CodeWhere<>"" Then 
        If UBound(All_Where)>=I Then
          SQL=SQL&" And " & All_Where(I)
        End If
      End If  
    Next
    Call ExecSQL(SQL)
  End If
  
  session("errnumber")=1
  session("msg")="代碼選項修改成功 ！"  
  response.redirect "codefile.asp?codetype="&request("codetype") 
End If

IsMember="N"
SQL1="Select IsMember From DEPT Where Dept_Id='"&session("dept_id")&"'"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then
  If RS1("IsMember")<>"" Then IsMember=RS1("IsMember")
End If
RS1.Close
Set RS1=Nothing

'SQL="Select * From CASECODE Where Ser_No="&request("ser_no")&" "
SQL="Select CODEFILE.Category,CASECODE.* From CODEFILE Join CASECODE On CODEFILE.CodeType=CASECODE.CodeType Where CASECODE.Ser_No="&request("ser_no")&""
Call QuerySQL(SQL,RS)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>修改代碼選項</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=gray>
  <form name="form" method="post" action="" target="main">
    <input type="hidden" name="action">
    <input type="hidden" name="CodeType" value="<%=RS("CodeType")%>">
    <input type="hidden" name="CodeDesc_Old" value="<%=RS("CodeDesc")%>">
    <div align="center"><center>
      <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td width="100%">
            <table width="100%" border=1 cellspacing="0" cellpadding="2">
              <tr>
                <td width="20%" align="right">
                  <font color="#000080">代碼名稱<span lang="en-us">：</span></font>
                </td>
                <td width="80%"><input type="text" name="CodeName" size="13" value="<%=RS("CodeName")%>" readonly class="font9t"></td>                
              </tr>         
              <tr> 
                <td align="right"><font color="#000080">代碼選項<span lang="en-us">：</span></font></td>
                <td><input type="text" name="CodeDesc" size="35" class="font9" maxlength="20" value="<%=RS("CodeDesc")%>" onKeypress='if(event.keyCode==13){JavaScript:Update_OnClick();}'></td>                
              </tr>
              <%If Session("Menu_Edit")="Y" And RS("Category")<>"類別選項" Then%>
              <tr> 
                <td align="right"><font color="#000080">連結<span lang="en-us">：</span></font></td>
                <td><input type="text" name="Menu_Url" size="35" class="font9" maxlength="40" value="<%=RS("Menu_Url")%>"></td>
              </tr>
               <tr> 
                <td align="right"><font color="#000080">視窗<span lang="en-us">：</span></font></td>
                <td>
                  <input type="radio" name="Menu_Target" value="_blank" <%If RS("Menu_Target")="_blank" Then%>checked<%End If%> >開啟新視窗&nbsp; 
                  <input type="radio" name="Menu_Target" value="_self" <%If RS("Menu_Target")<>"_blank" Then%>checked<%End If%> >原視窗
                </td>
              </tr>             
              <tr> 
                <td align="right"><font color="#000080">圖檔<span lang="en-us">：</span></font></td>
                <td><input type="text" name="Image_Sub_Url" size="35" class="font9" maxlength="200" value="<%=RS("Image_Sub_Url")%>"></td>
              </tr>
              <tr> 
                <td align="right"><font color="#000080">圖寬<span lang="en-us">：</span></font></td>
                <td><input type="text" name="Image_Sub_Width" size="35" class="font9" value="<%=RS("Image_Sub_Width")%>"></td>
              </tr>      
              <%Else%>
              <input type="hidden" name="Menu_Url" value="<%=RS("Menu_Url")%>">
              <input type="hidden" name="Menu_Target" value="<%=RS("Menu_Target")%>">	
              <input type="hidden" name="Image_Sub_Url" value="<%=RS("Image_Sub_Url")%>">
              <input type="hidden" name="Image_Sub_Width" value="<%=RS("Image_Sub_Width")%>">	
              <%End If%>
              <tr> 
                <td align="right"><font color="#000080">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;排序<span lang="en-us">：</span></font></td>
                <td><input type="text" name="Seq" size="8" class="font9" maxlength="3" value="<%=RS("Seq")%>" onKeypress='if(event.keyCode==13){JavaScript:Update_OnClick();}'></td>
              </tr>
              <%
                If RS("CodeType")="Purpose" Or RS("CodeType")="Accoun" Then
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
                <td colspan="3">
                  <input type="radio" name="CodeKind" value="D" <%If RS("CodeKind")="D" Then%>checked<%End If%> >愛心捐款&nbsp;
                  <%If IsMember="Y" Then%><input type="radio" name="CodeKind" value="M" <%If RS("CodeKind")="M" Then%>checked<%End If%> >會務繳費&nbsp;<%End If%>
                  <%If IsMember="Y" Then%><input type="radio" name="CodeKind" value="A" <%If RS("CodeKind")="A" Then%>checked<%End If%> >活動報名&nbsp;<%End If%>
                  <%If IsShopping="Y" Then%><input type="radio" name="CodeKind" value="S" <%If RS("CodeKind")="S" Then%>checked<%End If%> >商品義賣&nbsp;<%End If%>
                  <!--<input type="radio" name="CodeKind" value="V">志工-->	
                </td>
                <td> </td>
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
                <td width="100%" height=15 align="center" colspan="2">
                  <button id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onclick='Update_OnClick()'> <img src='../images/update.gif' width='19' height='20' align='absmiddle'> 修改</button>&nbsp;
                  <button id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onclick='window.close();'> <img src='../images/icon6.gif' width='20' height='15' align='absmiddle'> 離開</button>
                </td>
              </tr>   
            </table>
          </td>
        </tr>
      </table>
    </center></div>
  </form>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Update_OnClick(){
  <%call CheckStringJ("CodeDesc","代碼選項")%>
  <%call ChecklenJ("CodeDesc",60,"代碼選項")%>
  <%call CheckStringJ("Seq","排序")%>
  <%call CheckNumberJ("Seq","排序")%>
  <%call SubmitJ("update")%>
  window.close();
}
--></script>