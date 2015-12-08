<!--#include file="../include/dbfunctionJ.asp"-->
<%
News_Type=""
If request("code_id")<>"" Then
  SQL="Select News_Type=CodeDesc From CASECODE Where CodeType='NewsType' And Seq='"&request("code_id")&"' " 
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then News_Type=RS("News_Type")
End If 

Sel_Type="<option value=''> </option>"
If News_Type<>"" Then
  SQL="Select CodeDesc From CASECODE Where CodeType='NewsType' And CodeDesc='"&News_Type&"'"
Else
  SQL="Select CodeDesc From CASECODE Where CodeType='NewsType' Order By Seq"
End If
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.Open SQL,Conn,1,1
Do While Not RS.EOF
  If Cstr(RS("CodeDesc"))=Cstr(request("newstype")) Then
    Sel_Type = Sel_Type & "<option value='" & RS("CodeDesc") & "' selected >" & RS("CodeDesc") & "</option>"
  Else
    Sel_Type = Sel_Type & "<option value='" & RS("CodeDesc") & "'>" & RS("CodeDesc") & "</option>"
  End If
  RS.MoveNext
Loop
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>頁面查詢</title>
  <base target="_self">
</head>
<body class=left style="background-color:#EEEEE3">
  <form method="POST" name="form" action> 
    <table border="1" width="100%"  align="center" cellspacing="0"> 
      <tr> 
        <td width="100%" class="font11" style="background-color: #dcdcba">
          <img border="0" src="../images/icon6.GIF" align="absmiddle">
          <%If request("code_id")<>"" Then%>
          <Select name='NewsType' Size='1' style='font-size: 9pt; font-family: 新細明體' onchange='Type_OnChange2(this.value,<%=request("code_id")%>)'><%=Sel_Type%></Select>
          <%Else%>
          <Select name='NewsType' Size='1' style='font-size: 9pt; font-family: 新細明體' onchange='Type_OnChange(this.value)'><%=Sel_Type%></Select>
          <%End If%>
        </td> 
      </tr> 
    </table> 
    <%
      SQL="Select Top 100 Ser_No,訊息標題=News_Subject From NEWS Where NEWS_Type='"&request("newstype")&"' And News_EndDate >= '"&Date()&"' Order By NEWS_Type,Ser_No Desc" 
      HLink="news_edit.asp?code_id="&request("code_id")&"&ser_no=" 
      LinkParam="ser_no" 
      LinkTarget="main" 
      call SubjectList (SQL,HLink,LinkParam,LinkTarget,LinkType) 
    %>
  </form>        
</body> 
</html> 
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
  function Type_OnChange(DataType){	 
    document.location.href='newstype_list.asp?newstype='+DataType+'&code_id=';
  }	
  function Type_OnChange2(DataType,code_id){	 
    document.location.href='newstype_list.asp?newstype='+DataType+'&code_id='+code_id+'';
  }
--></script> 