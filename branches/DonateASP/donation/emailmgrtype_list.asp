<!--#include file="../include/dbfunctionJ.asp"-->
<%
Sel_Type="<option value=''> </option>"
SQL="Select CodeDesc From CASECODE Where CodeType='EmailMgrType' Order By Seq"
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.Open SQL,Conn,1,1
Do While Not RS.EOF
  If Cstr(RS("CodeDesc"))=Cstr(request("emailmgrtype")) Then
    Sel_Type = Sel_Type & "<option value='" & RS("CodeDesc") & "' selected >" & RS("CodeDesc") & "</option>"
  Else
    Sel_Type = Sel_Type & "<option value='" & RS("CodeDesc") & "'>" & RS("CodeDesc") & "</option>"
  End If
  RS.MoveNext
Loop
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
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
          <Select name='NewsType' Size='1' style='font-size: 9pt; font-family: 新細明體' onchange='Type_OnChange(this.value)'><%=Sel_Type%></Select>
        </td> 
      </tr> 
    </table> 
    <%
      SQL="Select Top 100 Ser_No,郵件標題=EmailMgr_Subject From EMAILMGR Where EmailMgr_Type='"&request("emailmgrtype")&"' Order By EmailMgr_Type,Ser_No Desc" 
      HLink="emailmgr_edit.asp?ser_no=" 
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
    document.location.href='emailmgrtype_list.asp?emailmgrtype='+DataType+'';
  }
--></script> 