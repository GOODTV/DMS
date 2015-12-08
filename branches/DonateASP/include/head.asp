<%Prog_Desc=ProgDesc(Prog_Id)%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <%If request("action")<>"export" And request("action")<>"mobile" And request("action")<>"email" Then%><link rel="stylesheet" type="text/css" href="../include/dms.css"><%End If%>
  <title><%=Prog_Desc%></title>
</head>