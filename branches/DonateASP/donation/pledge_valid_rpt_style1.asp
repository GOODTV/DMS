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
  <title><%=ProgDesc("pledge_valid")%></title>
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
      page-break-after:always;
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
        If RS1("Valid_Date")<>"" Then
          YY1=Cint("20"&Right(RS1("Valid_Date"),2))
          MM1=Cint(Left(RS1("Valid_Date"),2))
        Else
          YY1=Cint(Year(Date()))
          MM1=Cint(Month(Date()))
        End If
        
        YY2=YY1
        MM2=MM1+1
        If Cint(MM2)=13 Then
          YY2=YY1+1
          MM2=1
        End If
  
        If Len(MM1)=1 Then MM1="0"&MM1
        If Len(MM2)=1 Then MM2="0"&MM2
        
        EmailMgrDesc=EmailMgr_Desc
        EmailMgrDesc=Replace(EmailMgrDesc,"○○○",RS1("Donor_Name"))
        EmailMgrDesc=Replace(EmailMgrDesc,"YY1",YY1)
        EmailMgrDesc=Replace(EmailMgrDesc,"MM1",MM1)
        EmailMgrDesc=Replace(EmailMgrDesc,"YY2",YY2)
        EmailMgrDesc=Replace(EmailMgrDesc,"MM2",MM2)
        EmailMgrDesc=Replace(EmailMgrDesc,"年月日",Date())
        
        Response.Write "<table id='grid' align='center' valign='top' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"     
        Response.Write "  <tr><td width='5%' valign='top'> </td><td width='90%' valign='top'>"
        If EmailMgr_Desc<>"" Then Response.Write Replace(EmailMgrDesc,"○○○",Data_Minus(RS1("Donor_Name")))
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