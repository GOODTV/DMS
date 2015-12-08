<!--#include file="../include/dbfunctionJ.asp"-->
<%
'邊界:左:8 右:8 上:8 下:8
SQL1=Session("SQL1")
Donate_Where=Session("Donate_Where")
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
'Response.write SQL1
'Response.write Donate_Where
'Response.end
'20140122 Add by GoodTV Tanya:顯示列印筆數
PrintCount = RS1.RecordCount
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=ProgDesc("invoice_yearly2_prit")%></title>
  <object id="factory" viewastext style="display:none" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="http://mis.npois.com.tw/smsx.cab"></object>
  <script language="javascript">
    function window.onload(){
      factory.printing.header='' ;          //頁首
      factory.printing.footer='' ;          //頁尾
      factory.printing.portrait = true ;    //直印(true),橫印(false)
      factory.printing.leftMargin = 8.0 ;   //左邊界
      factory.printing.topMargin = 8.0 ;    //上邊界
      factory.printing.rightMargin = 8.0 ;  //右邊界
      factory.printing.bottomMargin = 8.0 ; //下邊界
      factory.printing.print();
    }
    //20140122 Add by GoodTV Tanya:顯示列印筆數
    function alertInfomsg()
    {
    	var printCount = document.getElementById('printCount').value;
    	alert('年度證明列印：共'+printCount+'筆');
    }
  </Script>
  <style>
  <!--
    table.TableGrid{
	    width:194mm;
    }
    .Sale_Content{
      font-size:18;
      font-family:"標楷體";
    }    
    .Donate_Desc{
      font-size:14;
      font-family:"標楷體";
    }  
    .CellMiddle{
      height:8mm;
    }
    .PageBreak {
      page-break-after:always;
    }
  -->
  </style>
</head>
<body class=tool <%If Not RS1.EOF And Session("action")="report" Then%>onload='print();alertInfomsg();'<%End If%>>	
  <p><div align="center"><center>
  	<!--20140122 Add by GoodTV Tanya:顯示列印筆數-->
  	<input type="hidden" id="printCount" value="<%=PrintCount%>">
  <%
    If RS1.EOF Then
      Response.Write "<br /><br /><br /><font size=3>沒有符合條件的資料可以列印!!</font>"
    Else
      Sale_Content=""
      SQL2="Select TOP 1 * From DONATE_SALE Where ('"&Date()&"' Between Sale_BeginDate And Sale_EndDate) Order By Sale_BeginDate Desc,Sale_EndDate Desc,Ser_No Desc"
      Set RS2 = Server.CreateObject("ADODB.RecordSet")
      RS2.Open SQL2,Conn,1,1
      If Not RS2.EOF Then
        If RS2("Sale_Content")<>"" Then Sale_Content=Replace(RS2("Sale_Content"),vbcrlf,"<br />")
      End If
      RS2.Close
      Set RS2=Nothing

      While Not RS1.EOF
        Response.Write "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        Response.Write "  <tr><td style='height:7mm'></td></tr>"&vbcrlf
        Response.Write "  <tr>"&vbcrlf
        Response.Write "    <td>"&vbcrlf
        Response.Write "      <table id='grid_1' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        Response.Write "        <tr>"&vbcrlf
				Response.Write "          <td style='width:17mm;height:82mm' valign='top' align='center'>&nbsp;</td>"&vbcrlf	
        Response.Write "          <td style='width:75mm;height:85mm' valign='top' align='left' class='Sale_Content'><font class='Sale_Content'>"& RS1("Donor_Name") &"</font></td>"&vbcrlf
				SQL2="Select IsAbroad_Invoice From DONOR Where Donor_Id='"&RS1("Donor_Id")&"' "
				Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,1
				IsAbroad_Invoice=RS2("IsAbroad_Invoice")
				If IsAbroad_Invoice="N" Then
					SQL2="Select *,地址=(Case When ISNULL(DONOR.Invoice_City,'')='' Then Address Else Case When A.mValue<>B.mValue Then ISNULL(DONOR.Invoice_ZipCode,'')+A.mValue+B.mValue+Invoice_Address Else ISNULL(DONOR.Invoice_ZipCode,'')+A.mValue+Invoice_Address End End) From DONOR Left Join CODECITY As A On DONOR.Invoice_City=A.mCode Left Join CODECITY As B On DONOR.Invoice_Area=B.mCode Where Donor_Id='"&RS1("Donor_Id")&"' "
				ElseIf IsAbroad_Invoice="Y" Then 
					SQL2="Select * From Donor Where Donor_Id='"&RS1("Donor_Id")&"' "
				End if
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,1
				If IsAbroad_Invoice="N" Then
					If len(RS2("地址"))>22 Then 
						Response.Write "          <td style='width:102mm;height:82mm' valign='top' class='Sale_Content'><br><br><br><br><br>"& left(RS2("地址"),22)
						Response.write "<br>"&Right(RS2("地址"),len(RS2("地址"))-22)
						If RS2("Invoice_Attn")<>"" then Response.write "("&RS2("Invoice_Attn")&")"
					Else
						Response.Write "          <td style='width:102mm;height:82mm' valign='top' class='Sale_Content'><br><br><br><br><br>"& left(RS2("地址"),22)
						If RS2("Invoice_Attn")<>"" then Response.write "("&RS2("Invoice_Attn")&")"
						Response.write "<br>"
					End If
					Response.write "<br>"& RS2("Donor_Name") &"　"& RS2("Title") &"　　鈞啟<br>"& RS1("Donor_Id") &"</td>"&vbcrlf
				Elseif IsAbroad_Invoice="Y" Then
					Response.Write "          <td style='width:102mm;height:82mm' valign='top' class='Sale_Content'><br><br><br><br><br>"& RS2("Donor_Name")&"<br>"& RS2("Invoice_OverseasAddress")
					Response.write "<br>"& RS2("Invoice_OverseasCountry")
					Response.write "<br>"& RS1("Donor_Id") &"</td>"&vbcrlf
					'Response.write "</td>"&vbcrlf
				End if
        Response.Write "        </tr>"&vbcrlf    'Sale_Content
        Response.Write "      </table>"&vbcrlf
        Response.Write "    </td>"&vbcrlf
        Response.Write "  </tr>"&vbcrlf
        
        Response.Write "  <tr>"&vbcrlf
        Response.Write "    <td>"&vbcrlf
        Response.Write "      <table id='grid_2' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        Response.Write "        <tr>"&vbcrlf
        Response.Write "          <td style='width:22mm;height:30mm'>&nbsp;</td>"&vbcrlf
        If Session("Invoice_Title_New")<>"" And RS1.Recordcount=1 Then
          Response.Write "          <td style='width:128mm;height:30mm' valign='bottom' class='Sale_Content'>"&Session("Invoice_Title_New")&"</td>"&vbcrlf
        Else
          Response.Write "          <td style='width:128mm;height:30mm' valign='bottom' class='Sale_Content'>"&RS1("Donor_Name")&"</td>"&vbcrlf
        End If
        Response.Write "          <td style='width:44mm;height:30mm' valign='bottom' class='Sale_Content'>"&RS1("Invoice_IDNo")&"</td>"&vbcrlf
        Response.Write "        </tr>"&vbcrlf
        Response.Write "      </table>"&vbcrlf
        Response.Write "    </td>"&vbcrlf
        Response.Write "  </tr>"&vbcrlf
        Response.Write "  <tr><td style='height:13mm'></td></tr>"&vbcrlf
        Response.Write "  <tr>"&vbcrlf
        Response.Write "    <td>"&vbcrlf
        Response.Write "      <table id='grid_3' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        
        Row=0
        '20140210 Modify by GoodTV Tanya:因特殊中文字造成查不到資料在Unicode字串加上前置詞N
        SQL2="Select Donate_Id,Donate_Date=CONVERT(VarChar,Donate_Date,111),Donate_Amt=Isnull(Donate_Amt,0),Donate_Payment,Donate_Purpose,Invoice_Pre,Invoice_No From Donate Join DONOR On Donate.Donor_Id=DONOR.Donor_Id "&Donate_Where&" And Donate.Donor_Id='"&RS1("Donor_Id")&"' And Donate.Invoice_Title=N'"&RS1("Donor_Name")&"' Order By Donate_Date,Donate_Id"        
        'Response.Write SQL2
      	'Response.End
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,1
				If Not RS1.Eof Then
		    		totRec=RS1.Recordcount
				    RS1.PageSize=20
				    totPage=RS1.PageCount
						nowPage=1
				End if
				'20140123 Modify by GoodTV Tanya:修改資料筆數為雙數或1筆時，位置會跑掉的問題
				'20140206 Modify by GoodTV Tanya:捐款日改為西元年
				'Row=1
		
        While Not RS2.EOF 
        	Row=Row+1          
          'Donate_Date_Year=Cstr(Year(RS2("Donate_Date"))-1911)
          Donate_Date_Year=Cstr(Year(RS2("Donate_Date")))
          Donate_Date_Month=Cstr(Month(RS2("Donate_Date")))
          If Len(Donate_Date_Month)=1 Then Donate_Date_Month="0"&Donate_Date_Month
          Donate_Date_Day=Cstr(Day(RS2("Donate_Date")))
          If Len(Donate_Date_Day)=1 Then Donate_Date_Day="0"&Donate_Date_Day
          Donate_Date=Donate_Date_Year&Donate_Date_Month&Donate_Date_Day
          If (Row Mod 2)=1 Then
            Response.Write "        <tr>"&vbcrlf
            Response.Write "          <td style='width:8mm;height:6mm' align='center' valign='center' class='Donate_Desc'></td>"&vbcrlf	
            Response.Write "          <td style='width:25.5mm;height:6mm' align='left' valign='center' class='Donate_Desc'>"&Donate_Date&"</td>"&vbcrlf
            Response.Write "          <td style='width:29mm;height:6mm' align='left' valign='center' class='Donate_Desc'>"&RS2("Invoice_Pre")&RS2("Invoice_No")&"</td>"&vbcrlf
            Response.Write "          <td style='width:22.5mm;height:6mm' align='right' valign='center' class='Donate_Desc'>"&FormatNumber(RS2("Donate_Amt"),0)&"</td>"&vbcrlf
            Response.Write "          <td style='width:22.5mm;height:6mm' align='right' valign='center' class='Donate_Desc'></td>"&vbcrlf
          Else
            Response.Write "          <td style='width:27mm;height:6mm' align='left' valign='center' class='Donate_Desc'>"&Donate_Date&"</td>"&vbcrlf
            Response.Write "          <td style='width:29mm;height:6mm' align='left' valign='center' class='Donate_Desc'>"&RS2("Invoice_Pre")&RS2("Invoice_No")&"</td>"&vbcrlf
            Response.Write "          <td style='width:22.5mm;height:6mm' align='right' valign='center' class='Donate_Desc'>"&FormatNumber(RS2("Donate_Amt"),0)&"</td>"&vbcrlf
            Response.Write "          <td style='width:12mm;height:6mm' align='right' valign='center' class='Donate_Desc'></td>"&vbcrlf
            Response.Write "        </tr>"&vbcrlf
          End If
          
				If Row mod 20 = 0 Then '跳頁處理
				'紀錄跨頁2014/3/14
				OverRecord = OverRecord&"No."&RS1("Donor_Id")&" "&RS1("Donor_Name")&"\n"
				  	Response.Write "      </table>"&vbcrlf
						Response.Write "    </td>"&vbcrlf
		        Response.Write "  </tr>"&vbcrlf
						Response.Write "</table>"&vbcrlf
						Response.Write "<div class='pagebreak'>&nbsp;</div>"
						Response.Write "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
		        Response.Write "  <tr><td style='height:7mm'></td></tr>"&vbcrlf
			      Response.Write "  <tr>"&vbcrlf
			      Response.Write "    <td>"&vbcrlf
			      Response.Write "      <table id='grid_1' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
			      Response.Write "        <tr>"&vbcrlf
						Response.Write "          <td style='width:17mm;height:82mm' valign='top' align='center'>&nbsp;</td>"&vbcrlf	
	        	Response.Write "          <td style='width:75mm;height:85mm' valign='top' align='left' class='Sale_Content'><font class='Sale_Content'>"& RS1("Donor_Name") &"</font></td>"&vbcrlf
	        	SQL3="Select IsAbroad_Invoice From DONOR Where Donor_Id='"&RS1("Donor_Id")&"' "
						Set RS3 = Server.CreateObject("ADODB.RecordSet")
				RS3.Open SQL3,Conn,1,1
						IsAbroad_Invoice=RS3("IsAbroad_Invoice")
						If IsAbroad_Invoice="N" Then
							SQL3="Select *,地址=(Case When ISNULL(DONOR.Invoice_City,'')='' Then Address Else Case When A.mValue<>B.mValue Then ISNULL(DONOR.Invoice_ZipCode,'')+A.mValue+B.mValue+Invoice_Address Else ISNULL(DONOR.Invoice_ZipCode,'')+A.mValue+Invoice_Address End End) From DONOR Left Join CODECITY As A On DONOR.Invoice_City=A.mCode Left Join CODECITY As B On DONOR.Invoice_Area=B.mCode Where Donor_Id='"&RS1("Donor_Id")&"' "
						ElseIf IsAbroad_Invoice="Y" Then 
							SQL3="Select * From Donor Where Donor_Id='"&RS1("Donor_Id")&"' "
						End if
				Set RS3 = Server.CreateObject("ADODB.RecordSet")
				RS3.Open SQL3,Conn,1,1
				If IsAbroad_Invoice="N" Then
					If len(RS3("地址"))>22 Then 
						Response.Write "          <td style='width:102mm;height:82mm' valign='top' class='Sale_Content'><br><br><br><br><br>"& left(RS3("地址"),22)
						Response.write "<br>"&Right(RS3("地址"),len(RS3("地址"))-22)
						If RS3("Invoice_Attn")<>"" then Response.write "("&RS3("Invoice_Attn")&")"
					Else
						Response.Write "          <td style='width:102mm;height:82mm' valign='top' class='Sale_Content'><br><br><br><br><br>"& left(RS3("地址"),22)
						If RS3("Invoice_Attn")<>"" then Response.write "("&RS3("Invoice_Attn")&")"
						Response.write "<br>"
					End If
					Response.write "<br>"& RS3("Donor_Name") &"　"& RS3("Title") &"　　鈞啟<br>"& RS1("Donor_Id") &"</td>"&vbcrlf
				Elseif IsAbroad_Invoice="Y" Then
					Response.Write "          <td style='width:102mm;height:82mm' valign='top' class='Sale_Content'><br><br><br><br><br>"& RS3("Donor_Name")&"<br>"& RS3("Invoice_OverseasAddress")
					Response.write "<br>"& RS3("Invoice_OverseasCountry")
					Response.write "<br>"& RS1("Donor_Id") &"</td>"&vbcrlf
					'Response.write "</td>"&vbcrlf
				End if
		        Response.Write "        </tr>"&vbcrlf    'Sale_Content
		        Response.Write "      </table>"&vbcrlf
		        Response.Write "    </td>"&vbcrlf
		        Response.Write "  </tr>"&vbcrlf
		        Response.Write "  <tr>"&vbcrlf
		        Response.Write "    <td>"&vbcrlf
		        Response.Write "      <table id='grid_2' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
		        Response.Write "        <tr>"&vbcrlf
		        Response.Write "          <td style='width:22mm;height:30mm'>&nbsp;</td>"&vbcrlf
		        If Session("Invoice_Title_New")<>"" And RS1.Recordcount=1 Then
				  Response.Write "          <td style='width:128mm;height:30mm' valign='bottom' class='Sale_Content'>"&Session("Invoice_Title_New")&"</td>"&vbcrlf
				Else
				  Response.Write "          <td style='width:128mm;height:30mm' valign='bottom' class='Sale_Content'>"&RS1("Donor_Name")&"</td>"&vbcrlf
				End If
				Response.Write "          <td style='width:44mm;height:30mm' valign='bottom' class='Sale_Content'>"&RS1("Invoice_IDNo")&"</td>"&vbcrlf
		        Response.Write "        </tr>"&vbcrlf
		        Response.Write "      </table>"&vbcrlf
		        Response.Write "    </td>"&vbcrlf
		        Response.Write "  </tr>"&vbcrlf
		        Response.Write "  <tr><td style='height:13mm'></td></tr>"&vbcrlf
		        Response.Write "  <tr>"&vbcrlf
		        Response.Write "    <td>"&vbcrlf
		        Response.Write "      <table id='grid_3' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf		
		  		End if
          RS2.MoveNext		  		
        Wend
        RS2.Close
        Set RS2=Nothing        
        If Row>0 And (Row Mod 2)=1 Then
          Response.Write "          <td style='width:27mm;height:6mm' align='left' valign='center' class='Donate_Desc'></td>"&vbcrlf
          Response.Write "          <td style='width:28.5mm;height:6mm' align='left' valign='center' class='Donate_Desc'></td>"&vbcrlf
          Response.Write "          <td style='width:27mm;height:6mm' align='right' valign='center' class='Donate_Desc'></td>"&vbcrlf
          Response.Write "          <td style='width:7mm;height:6mm' align='right' valign='center' class='Donate_Desc'></td>"&vbcrlf
          Response.Write "        </tr>"&vbcrlf
        End If
        Response.Write "      </table>"&vbcrlf
        Response.Write "    </td>"&vbcrlf
        Response.Write "  </tr>"&vbcrlf
				Response.Write "  <tr>"&vbcrlf
				'20140206 Modify by GoodTV Tanya:年度改為西元年
'		    Response.Write "    <td style='width:194mm;height:10mm' align='center' valign='center' class='Sale_Content'>"&Cint(Session("Donate_Year"))-1911&"年度捐款金額總計:新台幣"&ChineseMoney(RS1("Donate_Total"))&"&nbsp;(NT$"&FormatNumber(RS1("Donate_Total"),0)&")</td>"&vbcrlf
				Response.Write "    <td style='width:194mm;height:10mm' align='center' valign='center' class='Sale_Content'>"&Cint(Session("Donate_Year"))&"年度捐款金額總計:新台幣"&ChineseMoney(RS1("Donate_Total"))&"&nbsp;(NT$"&FormatNumber(RS1("Donate_Total"),0)&")</td>"&vbcrlf
	    	Response.Write "  </tr>"&vbcrlf
				Response.Write "</table>"&vbcrlf		
        Response.Flush
        Response.Clear
				nowPage=nowPage+1
				'20140213 Add by GoodTV Tanya:紀錄年度證明已列印
				SQL3="Update Donate Set Invoice_Print_Yearly='1' From Donate Join DONOR On Donate.Donor_Id=DONOR.Donor_Id "&Donate_Where&" And Donate.Donor_Id='"&RS1("Donor_Id")&"' And Donate.Invoice_Title=N'"&RS1("Donor_Name")&"'"        
				'Response.Write SQL3
				'Response.End
				Call ExecSQL(SQL3)

        RS1.MoveNext
        If Not RS1.EOF Then Response.Write "<div class='pagebreak'>&nbsp;</div>"
      Wend
      RS1.Close
      Set RS1=Nothing
    End If
    Session.Contents.Remove("DeptId")
    Session.Contents.Remove("DonateDesc")
    Session.Contents.Remove("Rept_Licence")
    Session.Contents.Remove("Print_Desc")
    Session.Contents.Remove("SQL1")
    Session.Contents.Remove("Donate_Where")
    Session.Contents.Remove("Donate_Year")
    Session.Contents.Remove("action")
    Session.Contents.Remove("Invoice_Title_New")
	'顯示跨頁訊息2014/3/14
	If OverRecord <>"" then
		session("errnumber")=1
		session("msg")=OverRecord&"跨頁"
		Message()
	End If
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->