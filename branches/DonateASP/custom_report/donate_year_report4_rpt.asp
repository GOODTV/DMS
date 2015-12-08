<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=donate_year_report4.xls"
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
      Response.Write "      <table border='0' style='width:190mm;' cellspacing='0' cellpadding='0'>"&vbcrlf
      Response.Write "  	    <tr>"&vbcrlf
      Response.Write "  		    <td style='height=10mm;' align='center' class='Month_Name' colspan='2'>國內地區捐款總額明細表</td>"&vbcrlf
      Response.Write "  	    </tr>"&vbcrlf
      Response.Write "  	    <tr>"&vbcrlf
      Response.Write "  		    <td style='height=5mm;' align='left' class='Title_Name'>統計期間:"&Session("Donate_Date_Begin")&"～"&Session("Donate_Date_End")&"</td>"&vbcrlf
      Response.Write "  		    <td style='height=5mm;' align='right' class='Title_Name'>依捐款人編號排序</td>"&vbcrlf
      Response.Write "  	    </tr>"&vbcrlf
      Response.Write "      </table>"&vbcrlf	
      Response.Write "    </td>"&vbcrlf
      Response.Write "  </tr>"&vbcrlf
      
      Response.Write "  <tr>"&vbcrlf
      Response.Write "    <td align='center' valign='top'>"&vbcrlf
      Response.Write "      <table border='0' style='width:190mm;' cellspacing='0' cellpadding='0'>"&vbcrlf
      Response.Write "  	    <tr>"&vbcrlf
      Response.Write "  		    <td style='width:20mm; height=10mm; border-top: 2px solid #000000; border-left: 2px solid #000000' class='Title_Name' align='center'>捐款人編號</td>"&vbcrlf
      Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>天使姓名</td>"&vbcrlf
	  Response.Write "  		    <td style='width:10mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>稱謂</td>"&vbcrlf
	  Response.Write "  		    <td style='width:25mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>累計奉獻金額</td>"&vbcrlf
	  Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>奉獻次數</td>"&vbcrlf
	  Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>郵遞區號</td>"&vbcrlf
      Response.Write "  		    <td style='width:30mm; height=10mm; border-top: 2px solid #000000; border-right: 2px solid #000000' class='Title_Name' align='left'>地址</td>"&vbcrlf
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
      Donate_Total=0
      Donate_Fee=0
      Donate_Forign=0
      While Not RS1.EOF
        Response.Write "  	  <tr>"&vbcrlf
        For I = 0 To FieldsCount
		  If I=0 then
			Response.Write "  		  <td style='height=5mm; border-top: 1px solid #000000;border-left: 2px solid #000000' class='List_Name' align='center'>"&Data_Minus(RS1(I))&"</td>"&vbcrlf
          ElseIf I=1 Then
            Response.Write "  	<td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='center'>"&Data_Minus(RS1(I))&"</td>"&vbcrlf
          'ElseIf I=3 Then
            'Response.Write "  	<td style='height=5mm; border-top: 1px solid #000000' class='List_Name'>"&Data_Minus(RS1(I))&"("&Left("00000",5-Len(RS1("Donor_Id")))&RS1("Donor_Id")&")</td>"&vbcrlf
          ElseIf I=3 Then
            Donate_Total=Donate_Total+Cdbl(RS1(I))
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='center'>"&FormatNumber(RS1(I),0)&"</td>"&vbcrlf
		  ElseIf I=FieldsCount Then
            Response.Write "    <td style='height=5mm; border-top: 1px solid #000000;border-right: 2px solid #000000' class='List_Name' align='left'>"&Data_Minus(RS1(I))&"</td>"&vbcrlf
          Else
            Response.Write "  	<td style='height=5mm; border-top: 1px solid #000000' class='List_Name' align='center'>"&Data_Minus(RS1(I))&"</td>"&vbcrlf
          End If
        Next
        Response.Write "  	  </tr>"&vbcrlf
        
		'20140123 Modify:匯出excel移除分頁，報表仍保留分頁
		If request("action")<>"export" Then
			'換頁處理
			If (Row+PageSize) Mod PageSize=0 Then
			  Response.Write "  	<tr>"&vbcrlf
			  Response.Write "  	  <td style='height=5mm; border-top: 2px solid #000000;border-left: 2px solid #000000;border-right: 2px solid #000000;border-bottom: 2px solid #000000' colspan='7'>"&vbcrlf
			  Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"&vbcrlf
			  Response.Write "  	      <tr>"&vbcrlf
			  If Row<RS1.Recordcount Then
				Response.Write "  	        <td style='width=160mm;' align='left' class='List_Name'></td>"&vbcrlf
				'Response.Write "  	      <td style='width=120mm;' align='center' class='List_Name'>累計&nbsp;&nbsp;&nbsp;捐款金額:"&FormatNumber(Donate_Total,0)&"元&nbsp;&nbsp;&nbsp;</td>"&vbcrlf
			  Else
				Response.Write "  	        <td style='width=160mm;' align='left' class='List_Name'>&nbsp;總筆數:"&TolRow&"&nbsp;筆</td>"&vbcrlf
				'Response.Write "  	      <td style='width=120mm;' align='center' class='List_Name'><b>總計&nbsp;&nbsp;&nbsp;捐款金額:"&FormatNumber(Donate_Total,0)&"元&nbsp;&nbsp;&nbsp;</b></td>"&vbcrlf
			  End If
			  Response.Write "  	        <td style='width=30mm;' align='right' class='List_Name'>頁數:"&PageNow&"/"&PageTotal&"&nbsp;</td>"&vbcrlf
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
				'Response.Write "      <table border='0' style='width:190mm;' cellspacing='0' cellpadding='0'>"&vbcrlf
				'Response.Write "  	    <tr>"&vbcrlf
				'Response.Write "  		    <td style='height=10mm;' align='center' class='Month_Name' colspan='2'>國內地區捐款總額明細表</td>"&vbcrlf
				'Response.Write "  	    </tr>"&vbcrlf
				'Response.Write "  	    <tr>"&vbcrlf
				'Response.Write "  		    <td style='height=5mm;' align='left' class='Title_Name'>統計期間:"&Session("Donate_Date_Begin")&"～"&Session("Donate_Date_End")&"</td>"&vbcrlf
				'Response.Write "  		    <td style='height=5mm;' align='right' class='Title_Name'>依捐款人編號排序</td>"&vbcrlf
				'Response.Write "  	    </tr>"&vbcrlf
				'Response.Write "      </table>"&vbcrlf	
				'Response.Write "    </td>"&vbcrlf
				'Response.Write "  </tr>"&vbcrlf
		  
				Response.Write "  <tr>"&vbcrlf
				Response.Write "    <td align='center' valign='top'>"&vbcrlf
				Response.Write "      <table border='0' style='width:190mm;' cellspacing='0' cellpadding='0'>"&vbcrlf
				Response.Write "  	    <tr>"&vbcrlf
				Response.Write "  		    <td style='width:20mm; height=10mm; border-top: 2px solid #000000; border-left: 2px solid #000000' class='Title_Name' align='center'>捐款人編號</td>"&vbcrlf
				Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>天使姓名</td>"&vbcrlf
				Response.Write "  		    <td style='width:10mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>稱謂</td>"&vbcrlf
				Response.Write "  		    <td style='width:25mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>累計奉獻金額</td>"&vbcrlf
				Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>奉獻次數</td>"&vbcrlf
				Response.Write "  		    <td style='width:15mm; height=10mm; border-top: 2px solid #000000' class='Title_Name' align='center'>郵遞區號</td>"&vbcrlf
				Response.Write "  		    <td style='width:30mm; height=10mm; border-top: 2px solid #000000; border-right: 2px solid #000000' class='Title_Name' align='left'>地址</td>"&vbcrlf
				Response.Write "  	    </tr>"&vbcrlf
			  End If
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
        Response.Write "  	  <td style='height=5mm; border-top: 2px solid #000000;border-left: 2px solid #000000;border-right: 2px solid #000000;border-bottom: 2px solid #000000' colspan='7'>"&vbcrlf
        Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"&vbcrlf
        Response.Write "  	      <tr>"&vbcrlf
        Response.Write "  	        <td style='width=70mm;' align='left' class='List_Name'>&nbsp;總筆數:"&TolRow&"&nbsp;筆</td>"&vbcrlf
        Response.Write "  	        <td style='width=120mm;' align='center' class='List_Name'><b>總計&nbsp;&nbsp;&nbsp;捐款金額:"&FormatNumber(Donate_Total,0)&"元&nbsp;&nbsp;&nbsp;</b></td>"&vbcrlf
        '20140123 Modify:匯出excel移除分頁，報表仍保留分頁
		If request("action")<>"export" Then
		Response.Write "  	        <td style='width=30mm;' align='right' class='List_Name'>頁數:"&PageNow&"/"&PageTotal&"&nbsp;</td>"&vbcrlf
		End If       
	    Response.Write "  	      </tr>"&vbcrlf
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
      Session.Contents.Remove("Donate_Date_Begin")
      Session.Contents.Remove("Donate_Date_End")
      Session.Contents.Remove("Donate_Total_Begin")
	  Session.Contents.Remove("Donate_Total_End")
      Session.Contents.Remove("Donate_Purpose")
    End If  
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->