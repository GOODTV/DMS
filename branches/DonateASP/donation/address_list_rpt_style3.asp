<!--#include file="../include/dbfunctionJ.asp"-->
<%
'8 x 2 (37.125mm * 105mm)
SQL1=Session("SQL1")
Call QuerySQL(SQL1,RS1)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=ProgDesc("address_list")%></title>
  <style>
  <!--
    table.TableGrid{
	    font-size: 13px;
	    width:194mm;
    }
    .Address{ 
	    width:97mm;
      font-size:20;
      font-family:"新細明體";
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
      Response.Write "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        Row=1
        While Not RS1.EOF
          Address=""
          If Session("Print_Desc")="1" Then
            If Data_Minus(RS1("Invoice_Address"))<>"" Then
              Address=RS1("Invoice_ZipCode")&"<br />"&Data_Minus(RS1("Invoice_Address"))&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"&nbsp;啟"
            Else
              Address=RS1("ZipCode")&"<br />"&RS1("Address")&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"&nbsp;啟"
            End If
          Else
            If RS1("Address")<>"" Then
              Address=RS1("ZipCode")&"<br />"&RS1("Address")&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"&nbsp;啟"
            Else
              Address=RS1("Invoice_ZipCode")&"<br />"&Data_Minus(RS1("Invoice_Address"))&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"&nbsp;啟"
            End If
          End If
          
          If (Row+16) Mod 16=1 Then
            Response.Write "<tr>"&vbcrlf
            Response.Write "  <td>"&vbcrlf
            Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
            Response.Write "      <tr><td style='height:29.125mm;width:89mm' align='left' valign='top' class='Address'>"&Address&"</td><td style='height:29.125mm;width:8mm'></td></tr>"&vbcrlf
            Response.Write "    </table>"&vbcrlf
            Response.Write "  </td>"&vbcrlf		 
          ElseIf (Row+16) Mod 16=2 Then
            Response.Write "  <td>"&vbcrlf
            Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
            Response.Write "      <tr><td style='height:29.125mm;width:8mm'></td><td style='height:29.125mm;width:89mm' align='left' valign='top' class='Address'>"&Address&"</td></tr>"&vbcrlf
            Response.Write "    </table>"&vbcrlf
            Response.Write "  </td>"&vbcrlf	
            Response.Write "</tr>"&vbcrlf
          ElseIf (Row+16) Mod 16=15 Then
            Response.Write "<tr>"&vbcrlf
            Response.Write "  <td>"&vbcrlf
            Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
            Response.Write "      <tr><td style='height:25mm;width:89mm' align='left' valign='top' class='Address'>"&Address&"</td><td style='height:25mm;width:8mm'></td></tr>"&vbcrlf
            Response.Write "    </table>"&vbcrlf
            Response.Write "  </td>"&vbcrlf
          ElseIf (Row+16) Mod 16=0 Then
            Response.Write "  <td>"&vbcrlf
            Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
            Response.Write "      <tr><td style='height:25mm;width:8mm'></td><td style='height:25mm;width:89mm' align='left' valign='top' class='Address'>"&Address&"</td></tr>"&vbcrlf
            Response.Write "    </table>"&vbcrlf
            Response.Write "  </td>"&vbcrlf	
            Response.Write "</tr>"&vbcrlf
            If Row<RS1.Recordcount Then
            	Response.Write "</table>"&vbcrlf
            	Response.Write "<div class='pagebreak'>&nbsp;</div>"
            	Response.Write "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
            End If
          Else
            If (Row+2) Mod 2=1 Then
              Response.Write "<tr>"&vbcrlf
              Response.Write "  <td>"&vbcrlf
              Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
              Response.Write "      <tr><td style='height:8mm;width:97mm' colspan='2'></td></tr>"&vbcrlf
              Response.Write "      <tr><td style='height:29.125mm;width:89mm' align='left' valign='top' class='Address'>"&Address&"</td><td style='height:29.125mm;width:8mm'></td></tr>"&vbcrlf
              Response.Write "    </table>"&vbcrlf
              Response.Write "  </td>"&vbcrlf	
            Else
              Response.Write "  <td>"&vbcrlf
              Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
              Response.Write "      <tr><td style='height:8mm;width:97mm' colspan='2'></td></tr>"&vbcrlf
              Response.Write "      <tr><td style='height:29.125mm;width:8mm'></td><td style='height:29.125mm;width:89mm' align='left' valign='top' class='Address'>"&Address&"</td></tr>"&vbcrlf
              Response.Write "    </table>"&vbcrlf
              Response.Write "  </td>"&vbcrlf	
              Response.Write "</tr>"&vbcrlf
            End If
          End If
          
          '感謝函地址名條列印
          If Instr(SQL1,"IsThanks_Add")>0 Then
            SQL2="Select IsThanks_Add From DONOR Where Donor_Id='"&RS1("Donor_Id")&"'"
            Set RS2 = Server.CreateObject("ADODB.RecordSet")
            RS2.Open SQL2,Conn,1,3
            If Not RS2.EOF Then
              RS2("IsThanks_Add")="1"
              RS2.Update
            End If
            RS2.Close
            Set RS2=Nothing
          End If

          '單筆收據地址名條列印
          If Instr(SQL1,"Invoice_Print_Add")>0 Then
            SQL2="Select Invoice_Print_Add From DONATE Where Donate_Id='"&RS1("Donate_Id")&"'"
            Set RS2 = Server.CreateObject("ADODB.RecordSet")
            RS2.Open SQL2,Conn,1,3
            If Not RS2.EOF Then
              RS2("Invoice_Print_Add")="1"
              RS2.Update
            End If
            RS2.Close
            Set RS2=Nothing
          End If

          '年度匯整收據地址名條列印
          If Instr(SQL1,"Invoice_Print_Yearly_Add")>0 Then
          SQL2="Update DONATE Set Invoice_Print_Yearly_Add='1' "&Session("Donate_Where")&" And Donor_Id='"&RS1("Donor_Id")&"' And Invoice_Title='"&Data_Minus(RS1("Donor_Name"))&"'"
          Set RS2=Conn.Execute(SQL2)
          End If
        
          Row=Row+1
          Response.Flush
          Response.Clear
          RS1.MoveNext
        Wend
        RS1.Close
        Set RS1=Nothing
        Row=Row-1
        
        While (Row+16) Mod 16<>0
          Row=Row+1
          If (Row+16) Mod 16=1 Then
            Response.Write "<tr>"&vbcrlf
            Response.Write "  <td>"&vbcrlf
            Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
            Response.Write "      <tr><td style='height:29.125mm;width:89mm' align='left' valign='top'></td><td style='height:29.125mm;width:8mm'></td></tr>"&vbcrlf
            Response.Write "    </table>"&vbcrlf
            Response.Write "  </td>"&vbcrlf		 
          ElseIf (Row+16) Mod 16=2 Then
            Response.Write "  <td>"&vbcrlf
            Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
            Response.Write "      <tr><td style='height:29.125mm;width:8mm'></td><td style='height:29.125mm;width:89mm' align='left' valign='top'></td></tr>"&vbcrlf
            Response.Write "    </table>"&vbcrlf
            Response.Write "  </td>"&vbcrlf	
            Response.Write "</tr>"&vbcrlf
          ElseIf (Row+16) Mod 16=15 Then
            Response.Write "<tr>"&vbcrlf
            Response.Write "  <td>"&vbcrlf
            Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
            Response.Write "      <tr><td style='height:27mm;width:89mm' align='left' valign='top'></td><td style='height:27mm;width:8mm'></td></tr>"&vbcrlf
            Response.Write "    </table>"&vbcrlf
            Response.Write "  </td>"&vbcrlf
          ElseIf (Row+16) Mod 16=0 Then
            Response.Write "  <td>"&vbcrlf
            Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
            Response.Write "      <tr><td style='height:27mm;width:8mm'></td><td style='height:27mm;width:89mm' align='left' valign='top'></td></tr>"&vbcrlf
            Response.Write "    </table>"&vbcrlf
            Response.Write "  </td>"&vbcrlf	
            Response.Write "</tr>"&vbcrlf
          Else
            If (Row+2) Mod 2=1 Then
              Response.Write "<tr>"&vbcrlf
              Response.Write "  <td>"&vbcrlf
              Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
              Response.Write "      <tr><td style='height:8mm;width:97mm' colspan='2'></td></tr>"&vbcrlf
              Response.Write "      <tr><td style='height:25mm;width:89mm' align='left' valign='top'></td><td style='height:25mm;width:8mm'></td></tr>"&vbcrlf
              Response.Write "    </table>"&vbcrlf
              Response.Write "  </td>"&vbcrlf	
            Else
              Response.Write "  <td>"&vbcrlf
              Response.Write "    <table border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
              Response.Write "      <tr><td style='height:8mm;width:97mm' colspan='2'></td></tr>"&vbcrlf
              Response.Write "      <tr><td style='height:25mm;width:8mm'></td><td style='height:25mm;width:89mm' align='left' valign='top'></td></tr>"&vbcrlf
              Response.Write "    </table>"&vbcrlf
              Response.Write "  </td>"&vbcrlf	
              Response.Write "</tr>"&vbcrlf
            End If
          End If
        Wend
      Response.Write "</table>"&vbcrlf
    End If
    Session.Contents.Remove("Donate_Where")
    Session.Contents.Remove("Print_Type")
    Session.Contents.Remove("Print_Desc")
    Session.Contents.Remove("SQL1")
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->