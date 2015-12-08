<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=donate_month_report5.xls"
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

SQL1=Session("SQL1")
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
SQL2=Session("SQL2")
Set RS2 = Server.CreateObject("ADODB.RecordSet")
RS2.Open SQL2,Conn,1,1
'response.write RS2(2)
'response.end()
Donate_Date_Begin=DateAdd("M",0,Session("Year")&"/"&Session("Month")&"/1")
Donate_Date_End=DateAdd("D",-1,DateAdd("M",1,Session("Year")&"/"&Session("Month")&"/1"))
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <%If request("action")<>"export" Then%><link REL="stylesheet" type="text/css" HREF="../include/dms.css"><%End If%>
  <title><%=ProgDesc("report_list")%></title>
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
<body class=tool <%If Not RS1.EOF Then%>onload='print();'<%End If%>>
  <p><div align="center"><center>
  <%
    If RS1.EOF Then
      Response.Write "<br /><br /><br /><font size=3>沒有符合條件的資料可以列印!!</font>"
    Else
      Response.Write "<table id='grid' align='center' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
      Response.Write "  <tr>"&vbcrlf
      Response.Write "    <td align='center' valign='top'>"&vbcrlf
      Response.Write "      <table border='0' style='width:200mm;' cellspacing='0' cellpadding='0'>"&vbcrlf
      Response.Write "  	    <tr>"&vbcrlf
      Response.Write "  		    <td style='height=10mm;' align='center' class='Month_Name' colspan='2'>當月各日捐款方式統計表</td>"&vbcrlf
      Response.Write "  	    </tr>"&vbcrlf
      Response.Write "  	    <tr>"&vbcrlf
      Response.Write "  		    <td style='height=5mm;' align='left' class='Title_Name'>捐款期間:"&Donate_Date_Begin&"～"&Donate_Date_End&"</td>"&vbcrlf
      Response.Write "  		    <td style='height=5mm;' align='right' class='Title_Name'>依捐款日期排序</td>"&vbcrlf
      Response.Write "  	    </tr>"&vbcrlf
      Response.Write "      </table>"&vbcrlf	
      Response.Write "    </td>"&vbcrlf
      Response.Write "  </tr>"&vbcrlf
      
      Response.Write "  <tr>"&vbcrlf
      Response.Write "    <td align='center' valign='top'>"&vbcrlf
      Response.Write "      <table border='0' style='width:200mm;' cellspacing='0' cellpadding='0'>"&vbcrlf
      Response.Write "  	    <tr>"&vbcrlf
      Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000; border-left: 2px solid #000000' class='Title_Name' align='center'>日期</td>"&vbcrlf
      Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>現金</td>"&vbcrlf
	  Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>劃撥</td>"&vbcrlf
      Response.Write "  		    <td style='width:20mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>信用卡授權書(一般)</td>"&vbcrlf
      Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>郵局轉帳</td>"&vbcrlf
      Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>匯款</td>"&vbcrlf
      Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>支票</td>"&vbcrlf
      Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>實物奉獻</td>"&vbcrlf
	  Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>ATM</td>"&vbcrlf
	  Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>網路信用卡</td>"&vbcrlf
	  Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>ACH</td>"&vbcrlf
	  Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>美國運通</td>"&vbcrlf
	  Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000; border-right: 2px solid #000000' class='Title_Name' align='center'>小計</td>"&vbcrlf
      Response.Write "  	    </tr>"&vbcrlf

      FieldsCount=RS1.Fields.Count-1
      PageSize=45
      PageNow=1
      PageTotal=1
      TolRow=RS1.Recordcount
      If TolRow>PageSize Then
        If TolRow Mod PageSize=0 Then
          PageTotal=TolRow/PageSize 
        Else
          PageTotal=(TolRow\PageSize)+1
        End If
      End If
      Row=1
      Fee_1=0
      Fee_2=0
      Fee_3=0
	  Fee_4=0
	  Fee_5=0
	  Fee_6=0
	  Fee_7=0
	  Fee_8=0
	  Fee_9=0
	  Fee_10=0
	  Fee_11=0
	  Fee_Total=0
      While Not RS1.EOF
        Response.Write "  	  <tr>"&vbcrlf
        'Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000;border-left: 2px solid #000000' class='List_Name' align='center'>"&Row&"</td>"&vbcrlf	
        For I = 0 To FieldsCount
          If I=0 Then
            Response.Write "  	<td style='height=5mm; border-top: 1px solid #000000;border-left: 2px solid #000000' class='List_Name' align='center'>"&Data_Minus(RS1(I))&"</td>"&vbcrlf
          ElseIf I=1 Then
            Fee_1=Fee_1+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
		  ElseIf I=2 Then
            Fee_2=Fee_2+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
          ElseIf I=3 Then
            Fee_3=Fee_3+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
          ElseIf I=4 Then
            Fee_4=Fee_4+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
          ElseIf I=5 Then
            Fee_5=Fee_5+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
		  ElseIf I=6 Then
            Fee_6=Fee_6+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
		  ElseIf I=7 Then
            Fee_7=Fee_7+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
		  ElseIf I=8 Then
            Fee_8=Fee_8+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
		  ElseIf I=9 Then
            Fee_9=Fee_9+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
		  ElseIf I=10 Then
            Fee_10=Fee_10+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
		  ElseIf I=11 Then
            Fee_11=Fee_11+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf                        
          ElseIf I=FieldsCount Then
			Fee_Total=Fee_Total+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000;border-right: 2px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
          'Else
            'Response.Write "  	<td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='right'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
          End If
        Next
        Response.Write "  	  </tr>"&vbcrlf
        
        '換頁處理
        If (Row+PageSize) Mod PageSize=0 Then
          Response.Write "  	<tr>"&vbcrlf
          Response.Write "  	  <td style='height=5mm; border-top: 2px solid #000000;border-left: 2px solid #000000;border-right: 2px solid #000000;border-bottom: 2px solid #000000' colspan='13'>"&vbcrlf
          Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"&vbcrlf
          Response.Write "  	      <tr>"&vbcrlf
          'Response.Write "  	        <td style='width=100mm;' align='left' class='List_Name'>&nbsp;列印日期:"&Now()&"</td>"&vbcrlf
          If Row<RS1.Recordcount Then
			Response.Write "  	        <td style='width=100mm;' align='left' class='List_Name'></td>"&vbcrlf
            'Response.Write "  	      <td style='width=120mm;' align='left' class='List_Name'>累計&nbsp;&nbsp;&nbsp;捐款金額:"&FormatNumber(Donate_Total,0)&"元&nbsp;&nbsp;&nbsp;手續費:"&FormatNumber(Donate_Fee,0)&"元&nbsp;&nbsp;&nbsp;實收金額:"&FormatNumber(Donate_Accou,0)&"元</td>"&vbcrlf
          Else
			Response.Write "  	        <td style='width=100mm;' align='left' class='List_Name'>&nbsp;總筆數:"&TolRow&"&nbsp;筆</td>"&vbcrlf
            'Response.Write "  	      <td style='width=120mm;' align='left' class='List_Name'><b>總計&nbsp;&nbsp;&nbsp;捐款金額:"&FormatNumber(Donate_Total,0)&"元&nbsp;&nbsp;&nbsp;手續費:"&FormatNumber(Donate_Fee,0)&"元&nbsp;&nbsp;&nbsp;實收金額:"&FormatNumber(Donate_Accou,0)&"元</b></td>"&vbcrlf
          End If
          Response.Write "  	        <td style='width=100mm;' align='right' class='List_Name'>頁數:"&PageNow&"/"&PageTotal&"&nbsp;</td>"&vbcrlf
          Response.Write "  	      </tr>"&vbcrlf
          Response.Write "  	    </table>"&vbcrlf
          Response.Write "  	  </td>"&vbcrlf
          Response.Write "  	</tr>"&vbcrlf
          PageNow=PageNow+1
          
          If Row<RS1.Recordcount Then
            Response.Write "      </table>"&vbcrlf
            Response.Write "    </td>"&vbcrlf
            Response.Write "  </tr>"&vbcrlf
            Response.Write "<table />"&vbcrlf
            Response.Write "<div class='pagebreak'>&nbsp;</div>"

            Response.Write "<table id='grid' align='center' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
            'Response.Write "  <tr>"&vbcrlf
            'Response.Write "    <td align='center' valign='top'>"&vbcrlf
            'Response.Write "      <table border='0' style='width:200mm;' cellspacing='0' cellpadding='0'>"&vbcrlf
            'Response.Write "  	    <tr>"&vbcrlf
            'Response.Write "  		    <td style='height=10mm;' align='center' class='Month_Name' colspan='2'>"&Year(Session("Donate_Date_Begin"))-1911&"年"&Month(Session("Donate_Date_Begin"))&"月份&nbsp;&nbsp;當月各日捐款方式統計表</td>"&vbcrlf
            'Response.Write "  	    </tr>"&vbcrlf
            'Response.Write "  	    <tr>"&vbcrlf
            'Response.Write "  		    <td style='height=5mm;' align='left' class='Title_Name'>授權終止起訖日期:"&Session("Donate_Date_Begin")&"～"&Session("Donate_Date_End")&"</td>"&vbcrlf
            'Response.Write "  		    <td style='height=5mm;' align='right' class='Title_Name'>依捐款人編號排序</td>"&vbcrlf
            'Response.Write "  	    </tr>"&vbcrlf
            'Response.Write "      </table>"&vbcrlf	
            'Response.Write "    </td>"&vbcrlf
            'Response.Write "  </tr>"&vbcrlf
      
            Response.Write "  <tr>"&vbcrlf
            Response.Write "    <td align='center' valign='top'>"&vbcrlf
            Response.Write "      <table border='0' style='width:200mm;' cellspacing='0' cellpadding='0'>"&vbcrlf
            'Response.Write "  	    <tr>"&vbcrlf
            'Response.Write "  		    <td style='width:12mm; height=10mm; border-top: 2px solid #000000; border-left: 2px solid #000000' class='Title_Name' align='center'>新授權書號</td>"&vbcrlf
			'Response.Write "  		    <td style='width:10mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>授權起日</td>"&vbcrlf
			'Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>授權迄日</td>"&vbcrlf
			'Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='left'>新授權額</td>"&vbcrlf
			'Response.Write "  		    <td style='width:25mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>授權方式</td>"&vbcrlf
			'Response.Write "  		    <td style='width:10mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>銀行</td>"&vbcrlf
			'Response.Write "  		    <td style='width:10mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>卡別</td>"&vbcrlf
			'Response.Write "  		    <td style='width:10mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>授權書號</td>"&vbcrlf
			'Response.Write "  		    <td style='width:13mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>捐款人編號</td>"&vbcrlf
			'Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>天使姓名</td>"&vbcrlf
			'Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>電話</td>"&vbcrlf
			'Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>手機</td>"&vbcrlf
			'Response.Write "  		    <td style='width:10mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>郵遞區號</td>"&vbcrlf
            'Response.Write "  		    <td style='width:30mm; height=10mm; border-top: 2px solid #000000; border-right: 2px solid #000000' class='Title_Name' align='center'>地址</td>"&vbcrlf
            'Response.Write "  	    </tr>"&vbcrlf            
          End If
        End If
        
	      Row=Row+1
	      Response.Flush
        Response.Clear
        RS1.MoveNext
      Wend
      
      Row=Row-1
      If (Row+PageSize) Mod PageSize<>0 Then
        Response.Write "  	<tr>"&vbcrlf
        Response.Write "  	  <td style='height=5mm; border-top: 2px solid #000000;border-left: 2px solid #000000;border-right: 2px solid #000000;border-bottom: 2px solid #000000' colspan='13'>"&vbcrlf
        Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"&vbcrlf
        Response.Write "  	      <tr>"&vbcrlf
        Response.Write "  		    <td style='width:15mm; class='List_Name' align='center'>總計</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&FormatNumber(Fee_1,0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&FormatNumber(Fee_2,0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:20mm; class='List_Name' align='right'>"&FormatNumber(Fee_3,0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&FormatNumber(Fee_4,0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&FormatNumber(Fee_5,0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&FormatNumber(Fee_6,0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&FormatNumber(Fee_7,0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&FormatNumber(Fee_8,0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&FormatNumber(Fee_9,0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&FormatNumber(Fee_10,0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&FormatNumber(Fee_11,0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&FormatNumber(Fee_Total,0)&"</td>"&vbcrlf
        Response.Write "  	      </tr>"&vbcrlf
		Response.Write "  	      <tr>"&vbcrlf
        Response.Write "  		    <td style='width:15mm; class='List_Name' align='center'>筆數</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&RS2(0)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&RS2(1)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:20mm; class='List_Name' align='right'>"&RS2(2)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&RS2(3)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&RS2(4)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&RS2(5)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&RS2(6)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&RS2(7)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&RS2(8)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&RS2(9)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&RS2(10)&"</td>"&vbcrlf
		  Response.Write "  		    <td style='width:15mm; class='List_Name' align='right'>"&RS2(11)&"</td>"&vbcrlf
        Response.Write "  	      </tr>"&vbcrlf
		'Response.Write "  	      <tr>"&vbcrlf
        'Response.Write "  	        <td style='width=110mm;' align='left' class='List_Name'>&nbsp;總筆數:"&TolRow&"&nbsp;筆</td>"&vbcrlf
        'Response.Write "  	        <td style='width=120mm;' align='left' class='List_Name'><b>總計&nbsp;&nbsp;&nbsp;捐款金額:"&FormatNumber(Donate_Total,0)&"元&nbsp;&nbsp;&nbsp;手續費:"&FormatNumber(Donate_Fee,0)&"元&nbsp;&nbsp;&nbsp;實收金額:"&FormatNumber(Donate_Accou,0)&"元</b></td>"&vbcrlf
        'Response.Write "  	        <td style='width=100mm;' align='right' class='List_Name'>頁數:"&PageNow&"/"&PageTotal&"&nbsp;</td>"&vbcrlf
        'Response.Write "  	      </tr>"&vbcrlf
        Response.Write "  	    </table>"&vbcrlf
        Response.Write "  	  </td>"&vbcrlf
        Response.Write "  	</tr>"&vbcrlf
      End If  
      
      Response.Write "      </table>"&vbcrlf
      Response.Write "    </td>"&vbcrlf
      Response.Write "  </tr>"&vbcrlf
      Response.Write "</table>"&vbcrlf
		  RS1.Close
		  Set RS1=Nothing
		  
      Session.Contents.Remove("SQL")
      Session.Contents.Remove("DeptId")
      Session.Contents.Remove("Year")
      Session.Contents.Remove("Month")
    End If  
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->