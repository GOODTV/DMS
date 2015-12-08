<!--#include file="../include/dbfunctionJ.asp"-->
<%
Sel_Type="<option value=''> </option>"
SQL="Select CodeName From CodeFile Where Category='單一網頁' Order By Menu_Seq"
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.Open SQL,Conn,1,1
Do While Not RS.EOF
  If Cstr(RS("CodeName"))=Cstr(request("contenttype")) Then
    Sel_Type = Sel_Type & "<option value='" & RS("CodeName") & "' selected >" & RS("CodeName") & "</option>"
  Else
    Sel_Type = Sel_Type & "<option value='" & RS("CodeName") & "'>" & RS("CodeName") & "</option>"
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
          <Select name='ContentType' Size='1' style='font-size: 9pt; font-family: 新細明體' onchange='Type_OnChange(this.value)'><%=Sel_Type%></Select>
        </td> 
      </tr> 
    </table> 
    <%
      SQL="Select Top 100 Ser_No,網頁標題=Content_Subject From CONTENT Where Content_Type='"&request("contenttype")&"' Order By Ser_No" 
      HLink="content_edit.asp?ser_no=" 
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
    document.location.href='contenttype_list.asp?contenttype='+DataType+'';
  }
--></script> 