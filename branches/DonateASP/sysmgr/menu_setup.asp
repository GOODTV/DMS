<!--#include file="../include/dbfunction.asp"-->
<%
Sub CheckBoxData (SQL,FName,DataName,DataNum,StrChecked)
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL,Conn,1,1
    session("FieldsetCount")=RS1.RecordCount
    FieldsCount = RS1.Fields.Count-1
    Dim i, j
    j = 1
    K = FieldsCount - DataNum
    Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>選擇</span></font></td>"
    For i = 1 To FieldsCount
	Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i).Name & "</span></font></td>"
    Next
    Response.Write "</tr>"
    While Not RS1.EOF
	Response.Write "<tr bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'><Input Type='checkbox'" & StrChecked & " Name='" & FName & j & "' value='" & RS1(FName) & "'>" & "</span></td>"
	For i = 1 To FieldsCount
            if i > k then
	           Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'><input type='text' name='" & DataName & i & j & "' size='4' style='font-size: 9pt; font-family: 新細明體' value='" & RS1(i) & "'></span></td>"
            else
               Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i) & "</span></td>"
            end if
	Next 
        j = j + 1   
        RS1.MoveNext
	Response.Write "</tr>"
    Wend
    Response.Write "</table>"
    RS1.Close
End sub
if request("action")="save" then
   For I = 1 to session("FieldsetCount")
     if Request("prog_id"&I)<>"" then
        Update_data()
     end if       
   Next
   Response.Redirect "menu_edit.asp?menu_id="&request("menu_id")
end if   

sub Update_data()
   sql="insert into prog_menu values ('"&request("menu_id")&"','"&Request("prog_id"&I)&"')"  
   set RS=Server.CreateObject("ADODB.Recordset")
   RS.Open sql,conn,1,3
end sub
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>程式群組管理--資料輸入</title>
</head>
<BODY>
<p>
<form name="form" action="" method="post">
  <input type="hidden" name="action"> 
  <input type="hidden" name="AddButton" value="menu_add.asp">
<table border="0" width="80%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"  height="25">
        程式群組管理【程式設定】</td>
  </tr>
  <tr>
    <td width="100%">
                    <div align="center">
                      <center>
                    <table width="100%" border=1 cellspacing="1" style="border-collapse: collapse" bordercolor="#6487DC" cellpadding="2">
                      <tr> 
                        <td noWrap align=right width="19%" bgcolor=#FFFFCC height=22>
                        <font color="#FF0000">*</font> <font color="#000080">
                        <span lang="zh-tw">程式群組</span>：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          <input class=font9 size=20 readonly name="menu_id" value="<%=request("menu_id")%>">
          <input type="button" value=" 存檔 " name="save" class="cbutton">
                          <input type="button" value=" 取消 " name="cancel" class="cbutton"></td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 height=22 class="font9" colspan="2">
                        <%SQL="select prog_id,序號=prog_seq,程式代號=prog_id,程式名稱=prog_desc from prog where prog_id NOT IN (select progid from prog_menu where menu_id='"&request("menu_id")&"')"_
       +"                       order by prog_seq"      
                          FName="prog_id"       
                          DataName="" 
                          DataNum=0    
                          call CheckBoxData (SQL,FName,DataName,DataNum,StrChecked)       
                         %>       
                        </td>
                        </table>
                      </center>
                    </div>
      </td>
  </tr>
</table>

<%message%>           
</form>
</BODY>
</HTML>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--

sub save_OnClick
  form.action.value="save"
  form.submit
end sub               
           
sub cancel_OnClick
  history.back
end sub

--></script>