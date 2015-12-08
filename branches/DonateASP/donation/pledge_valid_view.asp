<%Response.ContentType="text/html; charset=utf-8"%>
<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL="Select EmailMgr_Subject,EmailMgr_Desc From EMAILMGR Where ser_no='"&request("id")&"'"
call QuerySQL(SQL,RS)
If Not RS.EOF Then
  EmailMgr_Subject=RS("EmailMgr_Subject")
  EmailMgr_Desc=RS("EmailMgr_Desc")
End If
If EmailMgr_Desc<>"" Then 
  YY1=Cint(Year(Date()))
  MM1=Cint(Month(Date()))+1
  If MM1=13 Then
    MM1=1
    YY1=Cint(Year(Date()))+1
  End If
  
  YY2=YY1
  MM2=MM1+1
  If Cint(MM2)=13 Then
    YY2=YY1+1
    MM2=1
  End If
  
  If Len(MM1)=1 Then MM1="0"&MM1
  If Len(MM2)=1 Then MM2="0"&MM2
  
  EmailMgr_Desc=Replace(EmailMgr_Desc,"○○○",Session("user_name"))
  EmailMgr_Desc=Replace(EmailMgr_Desc,"YY1",YY1)
  EmailMgr_Desc=Replace(EmailMgr_Desc,"MM1",MM1)
  EmailMgr_Desc=Replace(EmailMgr_Desc,"YY2",YY2)
  EmailMgr_Desc=Replace(EmailMgr_Desc,"MM2",MM2)
  EmailMgr_Desc=Replace(EmailMgr_Desc,"年月日",Date())
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=ProgDesc("pledge_valid")%></title>
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