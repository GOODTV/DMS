<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=accoun.xls"
End If

'邊界:左:8 右:8 上:8 下:8
If Session("DeptId")<>"" Then
  SQL1="Select Comp_ShortName From DEPT Where Dept_Id='"&Session("DeptId")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  Comp_ShortName=RS1("Comp_ShortName")
  RS1.Close
  Set RS1=Nothing
End If
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <%If request("action")<>"export" Then%><link REL="stylesheet" type="text/css" HREF="../include/dms.css"><%End If%>
  <title><%=ProgDesc("donate_accoun_list")%></title>
  <style>
  <!--
    table.TableGrid{
	    font-size: 12px;
	    width:190mm;
    }
    .Month_Name{ 
      font-size:18;
      font-family:"新細明體";
    }       
    .Comp_Name{ 
      font-size:15;
      font-family:"新細明體";
    }
    .Title_Name{ 
      font-size:13;
      font-family:"新細明體";
    }
    .List_Name{ 
      font-size:12;
      font-family:"新細明體";
    }
    .PageBreak {
      page-break-after:always;
    }
  -->
  </style>
</head>
<body class=tool onload='print();'>	
  <p><div align="center"><center>
  <%
    Response.Write "<table id='grid' align='center' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
    Response.Write "  <tr>"&vbcrlf
    Response.Write "    <td align='center' valign='top'>"&vbcrlf
    Response.Write "      <table border='0' style='width:190mm;' cellspacing='0' cellpadding='0'>"&vbcrlf
    Response.Write "  	    <tr>"&vbcrlf
    Response.Write "  		    <td style='height=10mm;' align='center' class='Month_Name'>"&Year(Session("Donate_Date_Begin"))-1911&"年"&Month(Session("Donate_Date_Begin"))&"月份&nbsp;&nbsp;會計科目報表</td>"&vbcrlf
    Response.Write "  	    </tr>"&vbcrlf
    Response.Write "  	    <tr>"&vbcrlf
    Response.Write "  		    <td style='height=5mm;' align='left' class='Title_Name'>統計日期:"&Session("Donate_Date_Begin")&"～"&Session("Donate_Date_End")&"</td>"&vbcrlf
    Response.Write "  	    </tr>"&vbcrlf
    Response.Write "      </table>"&vbcrlf
    Response.Write "    </td>"&vbcrlf
    Response.Write "  </tr>"&vbcrlf

    Response.Write "  <tr>"&vbcrlf
    Response.Write "    <td align='center' valign='top'>"&vbcrlf
    Response.Write "      <table border='0' style='width:190mm;' cellspacing='0' cellpadding='0'>"&vbcrlf
    Response.Write "  	    <tr>"&vbcrlf
    Response.Write "  		    <td style='width:5mm; height=10mm; border-top: 2px solid #000000; border-left: 2px solid #000000' class='Title_Name'>&nbsp;</td>"&vbcrlf
    Response.Write "  		    <td style='width:60mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='left'>會計科目</td>"&vbcrlf
    Response.Write "  		    <td style='width:30mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='right'>捐款金額</td>"&vbcrlf
    Response.Write "  		    <td style='width:30mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='right'>手續費</td>"&vbcrlf
    Response.Write "  		    <td style='width:30mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='right'>實收金額</td>"&vbcrlf
    Response.Write "  		    <td style='width:30mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='right'>捐款筆數</td>"&vbcrlf
    Response.Write "  		    <td style='width:5mm; height=10mm; border-top: 2px solid #000000; border-right: 2px solid #000000' class='Title_Name'>&nbsp;</td>"&vbcrlf
    Response.Write "  	    </tr>"&vbcrlf
    
    Not_Known=""
    DonateTotal=0
    DonateFee=0
    DonateAccou=0
    DonateNo=0
    SQL2="Select Accounting_Title=CodeDesc From CASECODE Where CodeType='Accoun' Order By Seq"
    Set RS2 = Server.CreateObject("ADODB.RecordSet")
    RS2.Open SQL2,Conn,1,1
    While Not RS2.EOF
      If Not_Known="" Then
        Not_Known="'"&RS2("Accounting_Title")&"'"
      Else
        Not_Known=Not_Known&",'"&RS2("Accounting_Title")&"'"
      End If
      Donate_Amt=0
      Donate_No=0
      SQL3="Select Donate_Amt=Isnull(Sum(Donate_Amt),0),Donate_Fee=Isnull(Sum(Donate_Fee),0),Donate_Accou=Isnull(Sum(Donate_Accou),0),Donate_No=Count(*) From DONATE Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id Left Join ACT On DONATE.Act_Id=ACT.Act_Id Where Issue_Type<>'D' And Accounting_Title='"&RS2("Accounting_Title")&"' "&Session("Donate_Where")&" "
      Set RS3 = Server.CreateObject("ADODB.RecordSet")
      RS3.Open SQL3,Conn,1,1
      Donate_Amt=RS3("Donate_Amt")
      Donate_Fee=RS3("Donate_Fee")
      Donate_Accou=RS3("Donate_Accou")
      Donate_No=RS3("Donate_No")
      DonateTotal=DonateTotal+Cdbl(RS3("Donate_Amt"))
      DonateFee=DonateFee+Cdbl(RS3("Donate_Fee"))
      DonateAccou=DonateAccou+Cdbl(RS3("Donate_Accou"))
      DonateNo=DonateNo+Cdbl(RS3("Donate_No"))
      RS3.Close
		  Set RS3=Nothing
		
      Response.Write "  	  <tr>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000; border-left: 2px solid #000000' class='List_Name'>&nbsp;</td>"&vbcrlf	
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='left'>"&RS2("Accounting_Title")&"</td>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(Donate_Amt,0)&"元</td>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(Donate_Fee,0)&"元</td>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(Donate_Accou,0)&"元</td>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(Donate_No,0)&"筆</td>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000; border-right: 2px solid #000000' class='List_Name'>&nbsp;</td>"&vbcrlf	
	    Response.Write "  	  </tr>"&vbcrlf
	    Response.Flush
      Response.Clear
      RS2.MoveNext
    Wend
		RS2.Close
		Set RS2=Nothing
    If Not_Known<>"" Then
      Donate_Total=0
      Donate_No=0
      SQL3="Select Donate_Amt=Isnull(Sum(Donate_Amt),0),Donate_Fee=Isnull(Sum(Donate_Fee),0),Donate_Accou=Isnull(Sum(Donate_Accou),0),Donate_No=Count(*) From DONATE Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id Left Join ACT On DONATE.Act_Id=ACT.Act_Id Where Issue_Type<>'D' And Accounting_Title Not In ("&Not_Known&") "&Session("Donate_Where")&" "
      Set RS3 = Server.CreateObject("ADODB.RecordSet")
      RS3.Open SQL3,Conn,1,1
      Donate_Amt=RS3("Donate_Amt")
      Donate_Fee=RS3("Donate_Fee")
      Donate_Accou=RS3("Donate_Accou")
      Donate_No=RS3("Donate_No")
      DonateTotal=DonateTotal+Cdbl(RS3("Donate_Amt"))
      DonateFee=DonateFee+Cdbl(RS3("Donate_Fee"))
      DonateAccou=DonateAccou+Cdbl(RS3("Donate_Accou"))
      DonateNo=DonateNo+Cdbl(RS3("Donate_No"))
      RS3.Close
		  Set RS3=Nothing
      Response.Write "  	  <tr>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000; border-left: 2px solid #000000' class='List_Name'>&nbsp;</td>"&vbcrlf	
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='left'>科目不詳</td>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(Donate_Amt,0)&"元</td>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(Donate_Fee,0)&"元</td>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(Donate_Accou,0)&"元</td>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(Donate_No,0)&"筆</td>"&vbcrlf
      Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000; border-right: 2px solid #000000' class='List_Name'>&nbsp;</td>"&vbcrlf	
	    Response.Write "  	  </tr>"&vbcrlf
    End If
    
    SQL3="Select Invoice_No=Invoice_Pre+Invoice_No,Donate_Amt_D=Isnull(Donate_Amt_D,0) From DONATE Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id Left Join ACT On DONATE.Act_Id=ACT.Act_Id Where Issue_Type='D' "&Session("Donate_Where")&" "
    Set RS3 = Server.CreateObject("ADODB.RecordSet")
    RS3.Open SQL3,Conn,1,1
    If Not RS3.EOF Then
      Row=1
      Response.Write "  	  <tr>"&vbcrlf
      Response.Write "  	    <td style='height=5mm; border-top: 1px solid #000000;border-left: 2px solid #000000;border-right: 2px solid #000000;border-bottom: 2px solid #000000' colspan='7'>"&vbcrlf
      Response.Write "          <table border='0' cellspacing='0' cellpadding='0'>"&vbcrlf
      Response.Write "  	        <tr>"&vbcrlf
      Response.Write "  	          <td style='width=5mm; height=5mm;' align='left' class='List_Name'>&nbsp;</td>"&vbcrlf	
      Response.Write "  	          <td style='width=180mm; height=5mm;' align='left' class='List_Name'>作廢收據編號如下:<br />"&vbcrlf
      While Not RS3.EOF
        If Row<RS3.Recordcount Then
          Response.Write RS3("Invoice_No")&"&nbsp;"&FormatNumber(RS3("Donate_Amt_D"),0)&"元、"
        Else
          Response.Write RS3("Invoice_No")&"&nbsp;"&FormatNumber(RS3("Donate_Amt_D"),0)&"元"
        End If
        Row=Row+1  
        RS3.MoveNext
      Wend
      Response.Write "<br /><br /><br />"
      Response.Write "  	          </td>"&vbcrlf
      Response.Write "  	          <td style='width=5mm; height=5mm;' align='left' class='List_Name'>&nbsp;</td>"&vbcrlf	
      Response.Write "  	        </tr>"&vbcrlf
      Response.Write "  	      </table>"&vbcrlf
      Response.Write "  	    </td>"&vbcrlf
      Response.Write "  	  </tr>"&vbcrlf
    Else
      Response.Write "  	  <tr>"&vbcrlf
      Response.Write "  	    <td style='height=5mm; border-top: 1px solid #000000;border-left: 2px solid #000000;border-right: 2px solid #000000;border-bottom: 2px solid #000000' colspan='7'>"&vbcrlf
      Response.Write "          <table border='0' cellspacing='0' cellpadding='0'>"&vbcrlf
      Response.Write "  	        <tr>"&vbcrlf
      Response.Write "  	          <td style='width=190mm; height=20mm;' align='left' class='List_Name'>&nbsp;</td>"&vbcrlf	
      Response.Write "  	        </tr>"&vbcrlf
      Response.Write "  	      </table>"&vbcrlf
      Response.Write "  	    </td>"&vbcrlf
      Response.Write "  	  </tr>"&vbcrlf
    End If
    RS3.Close
		Set RS3=Nothing
		      
    Response.Write "  	    <tr>"&vbcrlf
    Response.Write "  	      <td style='height=5mm; border-top: 2px solid #000000;border-left: 2px solid #000000;border-right: 2px solid #000000;border-bottom: 2px solid #000000' colspan='7'>"&vbcrlf
    Response.Write "            <table border='0' cellspacing='0' cellpadding='0'>"&vbcrlf
    Response.Write "  	          <tr>"&vbcrlf
    Response.Write "  	            <td style='width=5mm; height=5mm;' align='left' class='List_Name'>&nbsp;</td>"&vbcrlf	
    Response.Write "  	            <td style='width=50mm; height=5mm;' align='left' class='List_Name'>列印日期:"&Now()&"</td>"&vbcrlf
    Response.Write "  	            <td style='width=10mm; height=5mm;' align='right' class='List_Name'>總計:</td>"&vbcrlf
    Response.Write "  	            <td style='width=30mm; height=5mm;' align='right' class='List_Name'>"&FormatNumber(DonateTotal,0)&"元</td>"&vbcrlf
    Response.Write "  	            <td style='width=30mm; height=5mm;' align='right' class='List_Name'>"&FormatNumber(DonateTotal,0)&"元</td>"&vbcrlf
    Response.Write "  	            <td style='width=30mm; height=5mm;' align='right' class='List_Name'>"&FormatNumber(DonateTotal,0)&"元</td>"&vbcrlf
    Response.Write "  	            <td style='width=30mm; height=5mm;' align='right' class='List_Name'>"&FormatNumber(DonateNo,0)&"</td>"&vbcrlf
    Response.Write "  	            <td style='width=5mm; height=5mm;' align='left' class='List_Name'>&nbsp;</td>"&vbcrlf	
    Response.Write "  	          </tr>"&vbcrlf
    Response.Write "  	        </table>"&vbcrlf
    Response.Write "  	      </td>"&vbcrlf
    Response.Write "  	    </tr>"&vbcrlf
    Response.Write "      </table>"&vbcrlf
    Response.Write "    </td>"&vbcrlf
    Response.Write "  </tr>"&vbcrlf
    Response.Write "</table>"&vbcrlf

    Session.Contents.Remove("Donate_Where")
    Session.Contents.Remove("DeptId")
    Session.Contents.Remove("Donate_Date_Begin")
    Session.Contents.Remove("Donate_Date_End")
    Session.Contents.Remove("Act_Id")
    Session.Contents.Remove("Donate_Payment")
    Session.Contents.Remove("Donate_Purpose")
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->