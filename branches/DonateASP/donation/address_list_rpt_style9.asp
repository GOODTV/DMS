<!--#include file="../include/dbfunctionJ.asp"-->
<%
'格式:13 x 3 (23mm * 70mm)
SQL1=Session("SQL1")
Call QuerySQL(SQL1,RS1)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=ProgDesc("address_list")%></title>
  <!--
    table.TableGrid{
	    font-size: 13px;
	    width:194mm;
    }
    .Address{ 
	    height:22.84615mm;
      font-size:15;
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
      Response.Write "  <tr><td style='height:14.84615mm;width:194mm' colspan='7'></td></tr>"&vbcrlf
      Row=1
      While Not RS1.EOF
        Address=""
        If Session("Print_Desc")="1" Then
          If Data_Minus(RS1("Invoice_Address"))<>"" Then
            Address=RS1("Invoice_ZipCode")&Data_Minus(RS1("Invoice_Address"))&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"&nbsp;啟"
          Else
            Address=RS1("ZipCode")&RS1("Address")&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"&nbsp;啟"
          End If
        Else
          If RS1("Address")<>"" Then
            Address=RS1("ZipCode")&RS1("Address")&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"&nbsp;啟"
          Else
            Address=RS1("Invoice_ZipCode")&Data_Minus(RS1("Invoice_Address"))&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"&nbsp;啟"
          End If
        End If
        If (Row+3) Mod 3=1 Then
          Response.Write "<tr>"&vbcrlf
          Response.Write "  <td style='height:22.84615mm;width:62mm' align='left' valign='center' class='Address'>"&Address&"</td><td style='height:22.84615mm;width:8mm'></td>"&vbcrlf
        ElseIf (Row+3) Mod 3=2 Then
          Response.Write "  <td style='height:22.84615mm;width:8mm'></td><td style='height:22.84615mm;width:62mm' align='left' valign='center' class='Address'>"&Address&"</td><td style='height:22.84615mm;width:8mm'></td>"&vbcrlf
        Else
          Response.Write "  <td style='height:22.84615mm;width:62mm' align='left' valign='center' class='Address'>"&Address&"</td><td style='height:22.84615mm;width:8mm'></td>"&vbcrlf
          Response.Write "</tr>"&vbcrlf	
        End If
        If (Row+33) Mod 33=0 Then 
          If Row<RS1.Recordcount Then
            Response.Write "</table>"&vbcrlf
            Response.Write "<div class='pagebreak'>&nbsp;</div>"          	
            Response.Write "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
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
        
      While (Row+33) Mod 33<>0
        Row=Row+1
        If (Row+3) Mod 3=1 Then
          Response.Write "<tr>"&vbcrlf
          Response.Write "  <td style='height:22.84615mm;width:62mm' align='left' valign='top'></td><td style='height:22.84615mm;width:8mm'></td>"&vbcrlf
        ElseIf (Row+3) Mod 3=2 Then
          Response.Write "  <td style='height:22.84615mm;width:8mm'></td><td style='height:22.84615mm;width:70mm' align='left' valign='top'></td><td style='height:22.84615mm;width:8mm'></td>"&vbcrlf
        Else
          Response.Write "  <td style='height:22.84615mm;width:62mm' align='left' valign='top'></td><td style='height:22.84615mm;width:8mm'></td>"&vbcrlf
          Response.Write "</tr>"&vbcrlf	
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