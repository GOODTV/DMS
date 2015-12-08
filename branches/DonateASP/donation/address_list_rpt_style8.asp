<!--#include file="../include/dbfunctionJ.asp"-->
<%
'格式:10 x 3 (29.7mm * 70mm)
'邊界:左4.23mm 右4.23mm 上4.23mm 下4.23mm
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
	    width:201.54mm;
    }
    .Tablecell_Top{
	    padding-left: 5mm;
	    padding-right: 5mm;
	    font-size: 12pt;
	    font-family: 標楷體;
	    width:70mm;
	    height:25.47mm;
    }    
    .Tablecell{
	    padding-left: 5mm;
	    padding-right: 5mm;
	    font-size: 12pt;
	    font-family: 標楷體;
	    width:70mm;
	    height:29.7mm;
	  }
    .Tablecell_Bottom{
	    padding-left: 5mm;
	    padding-right: 5mm;
	    font-size: 12pt;
	    font-family: 標楷體;
	    width:70mm;
	    height:22mm;
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

      Col=3       '每行筆數
      Page_Row=30 '每頁筆數
      Row=1
      While Not RS1.EOF
        If (Row+Page_Row) Mod Page_Row=1 Then Response.Write "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"
        If (Row+Col) Mod Col=1 Then Response.Write "<tr>"
        
        If Row=1 Or Row=2 Or Row=3 Then
          Response.Write "<td class='Tablecell_Top'>"
        ElseIf Row=28 Or Row=29 Or Row=30 Then
          Response.Write "<td class='Tablecell_Bottom'>"
        Else
          Response.Write "<td class='Tablecell'>"
        End If
        
        If Session("Print_Desc")="1" Then
          If Data_Minus(RS1("Invoice_Address"))<>"" Then
            Response.Write RS1("Invoice_ZipCode")&Data_Minus(RS1("Invoice_Address"))&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"啟"
          Else
            Response.Write RS1("ZipCode")&RS1("Address")&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"啟"
          End If
        Else
          If RS1("Address")<>"" Then
            Response.Write RS1("ZipCode")&RS1("Address")&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"啟"
          Else
            Response.Write RS1("Invoice_ZipCode")&Data_Minus(RS1("Invoice_Address"))&"<br />"&Data_Minus(RS1("Donor_Name"))&"&nbsp;"&RS1("title")&"啟"
          End If
        End If
        If Session("Print_Type")="1" Then Response.Write "<br /><span style='font-size:10pt'>"&RS1("Invoice_No")&"</span>"
        Response.Write "</td>"
        
        If (Row+Col) Mod Col=0 Then Response.Write "</tr>"
        If (Row+Page_Row) Mod Page_Row=0 Then Response.Write "</table><div class='pagebreak'>&nbsp;</div>"

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
        If Row>Page_Row Then Row=1
        Response.Flush
        Response.Clear
        RS1.MoveNext
      Wend
      While (Row+Page_Row) Mod Page_Row<>1
        If (Row+Col) Mod Col=1 Then Response.Write "<tr>"
        If Row=1 Or Row=2 Or Row=3 Then
          Response.Write "<td class='Tablecell_Top'>"
        ElseIf Row=28 Or Row=29 Or Row=30 Then
          Response.Write "<td class='Tablecell_Bottom'>"
        Else
          Response.Write "<td class='Tablecell'>"
        End If
        If (Row+Col) Mod Col=0 Then Response.Write "</tr>"
        If (Row+Page_Row) Mod Page_Row=0 Then Response.Write "</table>"
        Row=Row+1
      Wend
      RS1.Close
      Set RS1=Nothing
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