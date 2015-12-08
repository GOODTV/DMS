<%Response.ContentType="text/html; charset=utf-8"%>
<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL="Select EmailMgr_Subject,EmailMgr_Desc From EMAILMGR Where ser_no='"&request("id")&"'"
call QuerySQL(SQL,RS)
If Not RS.EOF Then
  EmailMgr_Subject=RS("EmailMgr_Subject")
  EmailMgr_Desc=RS("EmailMgr_Desc")
End If
If EmailMgr_Desc<>"" Then EmailMgr_Desc=Replace(EmailMgr_Desc,"○○○",Session("user_name"))
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=ProgDesc("other")%></title>
  <style type="text/css">
    body {
      margin-left: 0px;
	    margin-top: 0px;
	    margin-right: 0px;
	    margin-bottom: 0px;
	    background-color: #FFFFFF;
    }
  </style>  
</head>
<body>
  <table width="800" align="center" border="0" cellSpacing="0" cellPadding="1">
    <tbody>
      <tr>
        <td><%=EmailMgr_Desc%></td>
      </tr>
    </tbody>
  </table>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->