<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL1="Select EmailMgr_Desc From EMAILMGR Where Ser_No='"&Session("EmailMgr_Id")&"'"
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
If Not RS1.EOF Then EmailMgr_Desc=RS1("EmailMgr_Desc")
RS1.Close
Set RS1=Nothing

SQL1=Session("SQL1")
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=ProgDesc("other")%></title>
  <style>
  <!--
    table.TableGrid{
	    font-size: 13px;
	    width:194mm;
    }
    .Tablecell{
	    padding-left: 5mm;
	    padding-right: 5mm;
	    padding-top: 10pt;
	    padding-bottom: 10pt;
	    font-size: 12pt;
	    font-family: 標楷體;
	    width:96mm;
	    height:38mm;
	  }
    .CellMiddle{
      width:2mm;
    }
    .PageBreak {
      page-break-after:alothers;
    }
  -->
  </style>
</head>
<body class=tool <%If Not RS1.EOF Then%>onload='print();'<%End If%>>
  <p><div align="center"><center>
  <%
    If RS1.EOF Then
      Response.Write "<br><br><br><font size=3>沒有符合條件的資料可以列印!!</font>"
    Else
      Row=1
      While Not RS1.EOF
        EmailMgrDesc=EmailMgr_Desc
        EmailMgrDesc=Replace(EmailMgrDesc,"○○○",Data_Minus(RS1("Donor_Name")))
        Response.Write "<table id='grid' align='center' valign='top' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"     
        Response.Write "  <tr><td width='5%' valign='top'> </td><td width='90%' valign='top'>"
        If EmailMgr_Desc<>"" Then Response.Write EmailMgrDesc
        Response.Write "  </td><td width='5%' valign='top'> </td><tr>"	
        Response.Write "<table>"
        If Row<RS1.Recordcount Then Response.Write "<div class='pagebreak'>&nbsp;</div>"
        Row=Row+1
        Response.Flush
        Response.Clear
        RS1.MoveNext
      Wend
    End If
    RS1.Close
    Set RS1=Nothing
    Session("EmailMgr_Id")=Request("EmailMgr_Id")
    Session.Contents.Remove("SQL1")
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->