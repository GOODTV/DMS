<%
  Prog_Desc=""
  If Prog_Id<>"" Then
    SQL1="Select Prog_Desc From PROG Where Prog_Id='"&Prog_Id&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    If Not RS1.EOF Then Prog_Desc=RS1("Prog_Desc")
    RS1.Close
    Set RS1=Nothing
  End If
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <title><%=Prog_Desc%></title>
</head>