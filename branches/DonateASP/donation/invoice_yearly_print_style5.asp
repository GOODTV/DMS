<!--#include file="../include/dbfunctionJ.asp"-->
<%
'邊界:左:8 右:8 上:8 下:8
SQL1=Session("SQL1")
Donate_Where=Session("Donate_Where")
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=ProgDesc("invoice_yearly_print")%></title>
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
  </Script>
  <style>
  <!--
    table.TableGrid{
	    font-size: 12px;
	    width:194mm;
    }
    .Comp_Name{ 
	    width:194mm;
	    height:12mm;
      font-size:20;
      font-family:"標楷體";
    }
    .Invoice_No{
      height:5mm;
      font-size:15;
      font-family:"標楷體";
    }
    .Donate_Desc{
      height:8mm;
      font-size:15;
      font-family:"標楷體";
    }
    .Donate_Right{
      font-size:12;
      font-family:"標楷體";
    }
    .Donate_Seal{
      height:10mm;
      font-size:15;
      font-family:"標楷體";
    }
    .Donate_Account{
      height:8mm;
      font-size:15;
      font-family:"標楷體";
    }
    .Donate_Foot{
      height:5mm;
      font-size:12;
      font-family:"標楷體";
    }
    .Donate_Seal{
      height:10mm;
      font-size:15;
      font-family:"標楷體";
    }
    .Donate_Amt{
      height:9mm;
      font-size:13;
      font-family:"標楷體";
    }      
    .PageBreak {
      page-break-after:always;
    }
  -->
  </style>
</head>
<body class=tool <%If Not RS1.EOF And Session("action")="report" Then%>onload='print();'<%End If%>>	
  <p><div align="center"><center>
  <%
    If RS1.EOF Then
      Response.Write "<br /><br /><br /><font size=3>沒有符合條件的資料可以列印!!</font>"
    Else
      Row=1
      Max_Page=3
      SQL2="Select *,Dept_Address=Zip_Code+A.mValue+B.mValue+Address,Dept_Address2=Zip_Code2+C.mValue+D.mValue+Address2 From Dept " & _
           "Left Join CodeCity As A On Dept.City_Code=A.mCode " & _
           "Left Join CodeCity As B On Dept.Area_Code=B.mCode " & _
           "Left Join CodeCity As C On Dept.City_Code2=C.mCode " & _
           "Left Join CodeCity As D On Dept.Area_Code2=D.mCode Where Dept_Id='"&Session("DeptId")&"' "
      Set RS2 = Server.CreateObject("ADODB.RecordSet")
      RS2.Open SQL2,Conn,1,1
      If Not RS2.EOF Then
        Invoice_Pre=RS2("Invoice_Pre")
        Invoice_Rule_Type=RS2("Invoice_Rule_Type")
        Invoice_Rule_YMD=RS2("Invoice_Rule_YMD")
        Invoice_Rule_Len=RS2("Invoice_Rule_Len")
        Invoice_Rule_Pub=RS2("Invoice_Rule_Pub")
        Invoice_Pre2=RS2("Invoice_Pre2")
        Invoice_Rule_Type2=RS2("Invoice_Rule_Type2")
        Invoice_Rule_YMD2=RS2("Invoice_Rule_YMD2")
        Invoice_Rule_Len2=RS2("Invoice_Rule_Len2")
        Invoice_Rule_Pub2=RS2("Invoice_Rule_Pub2")
        Donate_Invoice=RS2("Donate_Invoice")
        
        Comp_Name=RS2("Comp_Name")
        Account=RS2("Account")
        Uniform_No=RS2("Uniform_No")
        Tel=RS2("Tel")
        Fax=RS2("Fax")
        EMail=RS2("EMail")
        Fax=RS2("Fax")
        Dept_Address=RS2("Dept_Address")
      End If
      RS2.Close
      Set RS2=Nothing

      '電子印章
      SEAL_1="../images/blank.gif"
      SEAL_2="../images/blank.gif"
      SEAL_3="../images/blank.gif"
      SEAL_D="../images/blank.gif"
      SQL2="Select * From SEAL Where (Ser_No<=4 Or Seal_Flg='Y' Or Seal_Subject='"&session("user_name")&"')"
      call QuerySQL(SQL2,RS2)
      While Not RS2.EOF
        If RS2("Seal_TitleImg")<>"" Then
          If RS2("Ser_No")="1" Then 
            SEAL_1="../upload/"&RS2("Seal_TitleImg")
          ElseIf RS2("Ser_No")="2" Then 
            SEAL_2="../upload/"&RS2("Seal_TitleImg")
          ElseIf RS2("Ser_No")="3" Then 
            SEAL_3="../upload/"&RS2("Seal_TitleImg")
          Else
            If RS2("Seal_Flg")="Y" Then
              If RS2("Seal_TitleImg")<>"" Then SEAL_D="../upload/"&RS2("Seal_TitleImg")
            Else
              If RS2("Seal_TitleImg")<>"" Then SEAL_4="../upload/"&RS2("Seal_TitleImg")
            End If
          End If
        End If
        RS2.MoveNext
      Wend
      RS2.Close
      Set RS2=Nothing
      If SEAL_4="" Then SEAL_4=SEAL_D
                
      While Not RS1.EOF
        '取捐款編號
        If RS1("Invoice_No")="" Then
          Invoice_Pre=Invoice_Pre
          Invoice_No=Get_Invoice_NoY("1",Session("DeptId"),Session("Donate_Year")&"/12/31",Invoice_Rule_Type2,Session("DonateDesc"),Invoice_Rule_Type,Invoice_Rule_YMD,Invoice_Rule_Len,Invoice_Rule_Pub)
        Else
          Invoice_Pre=RS1("Invoice_Pre")
          Invoice_No=RS1("Invoice_No") 
        End If
        
        For Page=1 To Max_Page
          Response.Write "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
          Response.Write "  <tr>"&vbcrlf
          Response.Write "    <td align='center' valign='top' style='height:20mm' colspan='3' class='Comp_Name'><u>"&Comp_Name&"<br />收據</u></td>"&vbcrlf
          Response.Write "  </tr>"&vbcrlf
          Response.Write "  <tr>"&vbcrlf
          Response.Write "    <td align='left' valign='top' style='height:5mm' class='Invoice_No'>捐款年度:"&Session("Donate_Year")&"年度</td>"&vbcrlf
          Response.Write "    <td align='right' valign='top' style='height:5mm' class='Invoice_No'>收據編號:"&Invoice_Pre&Invoice_No&"</td>"&vbcrlf
          Response.Write "    <td align='center' valign='top' style='height:5mm' class='Invoice_No'></td>"&vbcrlf
          Response.Write "  </tr>"&vbcrlf
          Response.Write "  <tr>"&vbcrlf
          Response.Write "    <td colspan='3' valign='top'>"&vbcrlf
          Response.Write "      <table id='grid_2' width='100%' border='1' cellpadding='0' cellspacing='0' bordercolor='#000000' style='border-collapse: collapse'>"&vbcrlf
          Response.Write "        <tr>"&vbcrlf
          Response.Write "          <td style='height:9mm;width:25mm' class='Donate_Desc' align='center'>捐款者</td>"&vbcrlf
          If Session("Invoice_Title_New")<>"" And RS1.Recordcount=1 Then
            Response.Write "        <td style='height:9mm;width:102mm' class='Donate_Desc'>&nbsp;"&Session("Invoice_Title_New")&"</td>"&vbcrlf
          Else
            Response.Write "        <td style='height:9mm;width:102mm' class='Donate_Desc'>&nbsp;"&RS1("Donor_Name")&"</td>"&vbcrlf
          End If
          Response.Write "          <td style='height:9mm;width:30mm' class='Donate_Desc' align='center'>身份證/統編</td>"&vbcrlf
          Response.Write "          <td style='height:9mm;width:29mm' class='Donate_Desc'>&nbsp;"&RS1("Invoice_IDNo")&"</td>"&vbcrlf
          Response.Write "        </tr>"&vbcrlf
          Response.Write "        <tr>"&vbcrlf
          Response.Write "          <td style='height:9mm;width:25mm' class='Donate_Desc' align='center'>總金額</td>"&vbcrlf
          Response.Write "          <td style='height:9mm;width:161mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;新台幣"&ChineseMoney(RS1("Donate_Total"))&"&nbsp;(NT$"&FormatNumber(RS1("Donate_Total"),0)&")</td>"&vbcrlf
          Response.Write "        </tr>"&vbcrlf
          Response.Write "        <tr>"&vbcrlf
          Response.Write "          <td style='height:9mm;width:25mm' align='center' class='Donate_Desc'>收據地址</td>"&vbcrlf
          Response.Write "          <td style='height:9mm;width:161mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;"&RS1("Invoice_Address2")&"</td>"&vbcrlf
          Response.Write "        </tr>"&vbcrlf
          Response.Write "      </table>"&vbcrlf	
          Response.Write "    </td>"&vbcrlf
          If Page=1 Then
            Response.Write "    <td style='height:27mm;width:8mm' align='center' valign='center' class='Donate_Right'>第<br />一<br />聯<br />：<br>捐<br>助<br>人<br>留<br>存</td>"&vbcrlf
          ElseIf Page=2 Then
            Response.Write "    <td style='height:27mm;width:8mm' align='center' valign='center' class='Donate_Right'>第<br />二<br />聯<br />：<br>會<br>計<br>聯</td>"&vbcrlf
          ElseIf Page=3 Then
            Response.Write "    <td style='height:27mm;width:8mm' align='center' valign='center' class='Donate_Right'>第<br />三<br />聯<br />：<br>存<br>根<br>聯</td>"&vbcrlf  
          End If
          Response.Write "  </tr>"&vbcrlf
          Response.Write "  <tr>"&vbcrlf
          Response.Write "    <td align='center' valign='top' style='height:10mm' colspan='3'>"&vbcrlf
          Response.Write "      <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
          Response.Write "        <tr>"&vbcrlf
          Response.Write "          <td style='width:16mm;height:10mm' align='left' class='Donate_Seal'>執行長:</td>"&vbcrlf
          Response.Write "          <td style='width:50mm;height:10mm' align='left' class='Donate_Seal'><img border='0' src='"&SEAL_2&"' style='height:10mm'></td>"&vbcrlf
          Response.Write "          <td style='width:12mm;height:10mm' align='left' class='Donate_Seal'>會計:</td>"&vbcrlf
          Response.Write "          <td style='width:50mm;height:10mm' align='left' class='Donate_Seal'><img border='0' src='"&SEAL_3&"' style='height:10mm'></td>"&vbcrlf
          Response.Write "          <td style='width:19mm;height:10mm' align='left' class='Donate_Seal'>經手人:</td>"&vbcrlf
	        Response.Write "          <td style='width:47mm;height:10mm' align='left' class='Donate_Seal'><img border='0' src='"&SEAL_4&"' style='height:10mm'></td>"&vbcrlf
          Response.Write "        </tr>"&vbcrlf
          Response.Write "      </table>"&vbcrlf
          Response.Write "    </td>"&vbcrlf
          Response.Write "  </tr>"&vbcrlf
          Response.Write "  <tr>"&vbcrlf
          Response.Write "    <td align='center' valign='top' style='height:18mm' colspan='3'>"&vbcrlf
          Response.Write "      <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
          Response.Write "        <tr>"&vbcrlf
          Response.Write "          <td style='width:64mm;height:8mm' align='left' class='Donate_Account'>劃撥帳號:"&Account&"</td>"&vbcrlf
          Response.Write "          <td style='width:130mm;height:8mm' align='left' class='Donate_Account' colspan='2'>戶名:"&Comp_Name&"</td>"&vbcrlf
          Response.Write "        </tr>"&vbcrlf
          Response.Write "        <tr>"&vbcrlf
          Response.Write "          <td style='width:64mm;height:5mm' align='left' class='Donate_Foot'>電話:"&TEL&"</td>"&vbcrlf
          Response.Write "          <td style='width:40mm;height:5mm' align='left' class='Donate_Foot'>傳真:"&Fax&"</td>"&vbcrlf
          Response.Write "          <td style='width:90mm;height:5mm' align='left' class='Donate_Foot'>地址:"&Dept_Address&"</td>"&vbcrlf
          Response.Write "        </tr>"&vbcrlf
          Response.Write "        <tr>"&vbcrlf
          Response.Write "          <td style='width:64mm;height:5mm' align='left' class='Donate_Foot'>EMail:"&EMail&"</td>"&vbcrlf
          Response.Write "          <td style='width:55mm;height:5mm' align='left' class='Donate_Foot'>統一編號:"&Uniform_No&"</td>"&vbcrlf
          'Response.Write "          <td style='width:75mm;height:5mm' align='left' class='Donate_Foot'>勸募許可文號:"&Session("Rept_Licence")&"</td>"&vbcrlf
          Response.Write "          <td style='width:75mm;height:5mm' align='left' class='Donate_Foot'>&nbsp;</td>"&vbcrlf
          Response.Write "        </tr>"&vbcrlf
          Response.Write "      </table>"&vbcrlf
          Response.Write "    </td>"&vbcrlf
          Response.Write "  </tr>"&vbcrl
          Response.Write "  <tr>"&vbcrlf
          Response.Write "    <td colspan='2'>"&vbcrlf
          Response.Write "      <table id='grid_4' width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
          Response.Write "        <tr>"&vbcrlf
          Response.Write "          <td style='height:10mm;width:194mm' align='left' valign='center' class='Donate_Seal'>收據明細如下:</td>"&vbcrlf
          Response.Write "        </tr>"&vbcrlf
          Response.Write "      </table>"&vbcrlf
          Response.Write "    </td>"&vbcrlf
          Response.Write "  </tr>"&vbcrlf
          Response.Write "  <tr>"&vbcrlf
          Response.Write "    <td colspan='2'>"&vbcrlf
	        Response.Write "      <table id='grid_5' width='100%' border='0' cellspacing='0' cellpadding='0' style='border: 1px solid #000000; padding-left: 1px; padding-right: 1px; padding-top: 1px; padding-bottom: 1px'>"&vbcrlf
	        Response.Write "        <tr>"&vbcrlf
		      Response.Write "          <td style='height:5mm;width:20mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>日期</td>"&vbcrlf
		      Response.Write "          <td style='height:5mm;width:20mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>金額</td>"&vbcrlf
		      Response.Write "          <td style='height:5mm;width:29mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>指定用途</td>"&vbcrlf
		      Response.Write "          <td style='height:5mm;width:25mm;border-bottom: 1px solid #000000;border-right: 1px solid #000000' class='Donate_Amt' align='center'>備註</td>"&vbcrlf
		      Response.Write "          <td style='height:5mm;width:20mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>日期</td>"&vbcrlf
		      Response.Write "          <td style='height:5mm;width:20mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>金額</td>"&vbcrlf
		      Response.Write "          <td style='height:5mm;width:29mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>指定用途</td>"&vbcrlf
	        Response.Write "          <td style='height:5mm;width:25mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>備註</td>"&vbcrlf
	        Response.Write "        </tr>"&vbcrlf
        
          Max_Row=30
          Row2=0
          SQL2="Select Donate_Id,Donate_Date=CONVERT(VarChar,Donate_Date,111),Donate_Amt=Isnull(Donate_Amt,0),Donate_Payment,Donate_Purpose,Invoice_Pre,Invoice_No,Invoice_Print_Yearly From Donate Join DONOR On Donate.Donor_Id=DONOR.Donor_Id "&Donate_Where&" And Donate.Donor_Id='"&RS1("Donor_Id")&"' And Donate.Invoice_Title='"&RS1("Donor_Name")&"' And Invoice_Pre='"&Invoice_Pre&"' And Invoice_No='"&Invoice_No&"' Order By Donate_Date,Donate_Id"
          Set RS2 = Server.CreateObject("ADODB.RecordSet")
          RS2.Open SQL2,Conn,1,1
          While Not RS2.EOF
            Row2=Row2+1
            If ((Row2+2) Mod 2)=1 Then Response.Write "<tr>"&vbcrlf
            Response.Write "<td style='height:10mm;' class='Donate_Amt' align='center'>"&RS2("Donate_Date")&"</td>"&vbcrlf
            Response.Write "<td style='height:10mm;' class='Donate_Amt' align='right'>"&FormatNumber(RS2("Donate_Amt"),0)&"</td>"&vbcrlf
            Response.Write "<td style='height:10mm;' class='Donate_Amt' align='left'>&nbsp;"&RS2("Donate_Purpose")&"</td>"&vbcrlf
            If ((Row2+2) Mod 2)=1 Then
              Response.Write "<td style='height:10mm;' class='Donate_Amt' align='left' style='border-right: 1px solid #000000'>&nbsp;</td>"&vbcrlf
            Else
              Response.Write "<td style='height:10mm;' class='Donate_Amt' align='left'>&nbsp;</td>"&vbcrlf
            End If
            If ((Row2+2) Mod 2)=0 Then Response.Write "</tr>"&vbcrlf
            
            '***換頁處理 Str***
            If ((Row2+Max_Row) Mod Max_Row=0) And Row2<RS2.Recordcount Then
              Response.Write "</table></td></tr></table>"&vbcrlf
              Response.Write "<div class='pagebreak'>&nbsp;</div>"
              Response.Write "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
              Response.Write "  <tr>"&vbcrlf
              Response.Write "    <td align='center' valign='top' style='height:20mm' colspan='3' class='Comp_Name'><u>"&Comp_Name&"<br />收據</u></td>"&vbcrlf
              Response.Write "  </tr>"&vbcrlf
              Response.Write "  <tr>"&vbcrlf
              Response.Write "    <td align='left' valign='top' style='height:5mm' class='Invoice_No'>捐款年度:"&Session("Donate_Year")&"年度</td>"&vbcrlf
              Response.Write "    <td align='right' valign='top' style='height:5mm' class='Invoice_No'>收據編號:"&Invoice_Pre&Invoice_No&"</td>"&vbcrlf
              Response.Write "    <td align='center' valign='top' style='height:5mm' class='Invoice_No'></td>"&vbcrlf
              Response.Write "  </tr>"&vbcrlf
              Response.Write "  <tr>"&vbcrlf
              Response.Write "    <td colspan='3'>"&vbcrlf
              Response.Write "      <table id='grid_2' width='100%' border='0' cellpadding='0' cellspacing='0' bordercolor='#000000' style='border-collapse: collapse'>"&vbcrlf
              Response.Write "        <tr>"&vbcrlf
              Response.Write "          <td style='height:9mm;width:25mm' class='Donate_Desc' align='center'>捐款者</td>"&vbcrlf
              If Session("Invoice_Title_New")<>"" And RS1.Recordcount=1 Then
                Response.Write "        <td style='height:9mm;width:102mm' class='Donate_Desc'>&nbsp;"&Session("Invoice_Title_New")&"</td>"&vbcrlf
              Else
                Response.Write "        <td style='height:9mm;width:102mm' class='Donate_Desc'>&nbsp;"&RS1("Donor_Name")&"</td>"&vbcrlf
              End If
              Response.Write "          <td style='height:9mm;width:30mm' class='Donate_Desc' align='center'>身份證/統編</td>"&vbcrlf
              Response.Write "          <td style='height:9mm;width:29mm' class='Donate_Desc'>&nbsp;"&RS1("IDNo")&"</td>"&vbcrlf
              Response.Write "        </tr>"&vbcrlf
              Response.Write "        <tr>"&vbcrlf
              Response.Write "          <td style='height:9mm;width:25mm' class='Donate_Desc' align='center'>總金額</td>"&vbcrlf
              Response.Write "          <td style='height:9mm;width:161mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;新台幣"&ChineseMoney(RS1("Donate_Total"))&"&nbsp;(NT$"&FormatNumber(RS1("Donate_Total"),0)&")</td>"&vbcrlf
              Response.Write "        </tr>"&vbcrlf
              Response.Write "        <tr>"&vbcrlf
              Response.Write "          <td style='height:9mm;width:25mm' align='center' class='Donate_Desc'>收據地址</td>"&vbcrlf
              Response.Write "          <td style='height:9mm;width:161mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;"&RS1("Invoice_Address2")&"</td>"&vbcrlf
              Response.Write "        </tr>"&vbcrlf
              Response.Write "      </table>"&vbcrlf	
              Response.Write "    </td>"&vbcrlf
              If Page=1 Then
                Response.Write "    <td style='height:45mm;width:8mm' align='center' valign='center' class='Donate_Right'>第<br />一<br />聯<br />：<br>捐<br>助<br>人<br>留<br>存</td>"&vbcrlf
              ElseIf Page=2 Then
                Response.Write "    <td style='height:45mm;width:8mm' align='center' valign='center' class='Donate_Right'>第<br />二<br />聯<br />：<br>會<br>計<br>聯</td>"&vbcrlf
              ElseIf Page=3 Then
                Response.Write "    <td style='height:45mm;width:8mm' align='center' valign='center' class='Donate_Right'>第<br />三<br />聯<br />：<br>存<br>根<br>聯</td>"&vbcrlf  
              End If
              Response.Write "  </tr>"&vbcrlf
              Response.Write "  <tr>"&vbcrlf
              Response.Write "    <td align='center' valign='top' colspan='2'>"&vbcrlf
              Response.Write "      <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
              Response.Write "        <tr><td style='width:194mm;height:3mm;' colspan='2'></td></tr>"&vbcrlf
              Response.Write "        <tr>"&vbcrlf
              Response.Write "          <td style='width:90mm;height:5mm' align='left' class='Donate_Foot'>統一編號:"&Uniform_No&"</td>"&vbcrlf
              Response.Write "          <td style='width:104mm;height:5mm' align='left' class='Donate_Foot'>法人登記:玖零證社玖號(90)投府社政字第90018108號</td>"&vbcrlf
              Response.Write "        </tr>"&vbcrlf
              Response.Write "        <tr>"&vbcrlf
              Response.Write "          <td style='width:90mm;height:5mm' align='left' class='Donate_Foot'>電&nbsp;&nbsp;&nbsp;&nbsp;話:"&TEL&"</td>"&vbcrlf
              Response.Write "          <td style='width:104mm;height:5mm' align='left' class='Donate_Foot'>地&nbsp;&nbsp;&nbsp;&nbsp;址:"&Dept_Address&"</td>"&vbcrlf
              Response.Write "        </tr>"&vbcrlf
              Response.Write "        <tr><td style='width:194mm;height:5mm;' colspan='2' class='Donate_Foot'>本收據可於申報所得稅時作為&#65378;列舉扣除額&#65379;之用</td></tr>"&vbcrlf
              Response.Write "      </table>"&vbcrlf
              Response.Write "    </td>"&vbcrlf
              Response.Write "  </tr>"&vbcrl
              Response.Write "  <tr>"&vbcrlf
              Response.Write "    <td colspan='2'>"&vbcrlf
              Response.Write "      <table id='grid_3' width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
              Response.Write "        <tr>"&vbcrlf
              Response.Write "          <td style='width:16mm;height:10mm' align='left' class='Donate_Seal'>執行長:</td>"&vbcrlf
              Response.Write "          <td style='width:50mm;height:10mm' align='left' class='Donate_Seal'><img border='0' src='"&SEAL_2&"' style='height:10mm'></td>"&vbcrlf
              Response.Write "          <td style='width:12mm;height:10mm' align='left' class='Donate_Seal'>會計:</td>"&vbcrlf
              Response.Write "          <td style='width:50mm;height:10mm' align='left' class='Donate_Seal'><img border='0' src='"&SEAL_3&"' style='height:10mm'></td>"&vbcrlf
              Response.Write "          <td style='width:19mm;height:10mm' align='left' class='Donate_Seal'>經手人:</td>"&vbcrlf
	            Response.Write "          <td style='width:47mm;height:10mm' align='left' class='Donate_Seal'><img border='0' src='"&SEAL_4&"' style='height:10mm'></td>"&vbcrlf
              Response.Write "        </tr>"&vbcrlf
              Response.Write "      </table>"&vbcrlf
              Response.Write "    </td>"&vbcrlf
              Response.Write "  </tr>"&vbcrlf
              Response.Write "  <tr>"&vbcrlf
              Response.Write "    <td colspan='2'>"&vbcrlf
              Response.Write "      <table id='grid_4' width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
              Response.Write "        <tr>"&vbcrlf
              Response.Write "          <td style='height:10mm;width:194mm' align='left' valign='center' class='Donate_Seal'>全年度捐款收據明細如下:</td>"&vbcrlf
              Response.Write "        </tr>"&vbcrlf
              Response.Write "      </table>"&vbcrlf
              Response.Write "    </td>"&vbcrlf
              Response.Write "  </tr>"&vbcrlf
              Response.Write "  <tr>"&vbcrlf
              Response.Write "    <td colspan='2'>"&vbcrlf
	            Response.Write "      <table id='grid_5' width='100%' border='0' cellspacing='0' cellpadding='0' style='border: 1px solid #000000; padding-left: 1px; padding-right: 1px; padding-top: 1px; padding-bottom: 1px'>"&vbcrlf
	            Response.Write "        <tr>"&vbcrlf
		          Response.Write "          <td style='height:5mm;width:20mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>日期</td>"&vbcrlf
		          Response.Write "          <td style='height:5mm;width:25mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>收據編號</td>"&vbcrlf
		          Response.Write "          <td style='height:5mm;width:20mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>金額</td>"&vbcrlf
		          Response.Write "          <td style='height:5mm;width:29mm;border-bottom: 1px solid #000000;border-right: 1px solid #000000' class='Donate_Amt' align='center'>指定用途</td>"&vbcrlf
		          Response.Write "          <td style='height:5mm;width:20mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>日期</td>"&vbcrlf
		          Response.Write "          <td style='height:5mm;width:25mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>收據編號</td>"&vbcrlf
		          Response.Write "          <td style='height:5mm;width:20mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>金額</td>"&vbcrlf
		          Response.Write "          <td style='height:5mm;width:29mm;border-bottom: 1px solid #000000' class='Donate_Amt' align='center'>指定用途</td>"&vbcrlf
	            Response.Write "        </tr>"&vbcrlf
            End If
            '***換頁處理 End***
            
            If Page=Max_Page Then
              SQL3="Select * From Donate Where Donate_Id='"&RS2("Donate_Id")&"'"
              Set RS3 = Server.CreateObject("ADODB.RecordSet")
              RS3.Open SQL3,Conn,1,3            
              RS3("Invoice_Print")="1"
              RS3("Invoice_Print_Yearly")="1"
              If RS3("Invoice_Pre")="" Then RS3("Invoice_Pre")=Invoice_Pre
              If RS3("Invoice_No")="" Then RS3("Invoice_No")=Invoice_No
              RS3.Update
              RS3.Close
              Set RS3=Nothing
            End If
            RS2.MoveNext
          Wend
          RS2.Close
          Set RS2=Nothing
        
          While (Row2+Max_Row) Mod Max_Row<>0
            Row2=Row2+1
            If ((Row2+2) Mod 2)=1 Then Response.Write "<tr>"&vbcrlf
            Response.Write "<td style='height:10mm;' class='Donate_Amt' align='center'>&nbsp;</td>"&vbcrlf
            Response.Write "<td style='height:10mm;' class='Donate_Amt' align='center'>&nbsp;</td>"&vbcrlf
            Response.Write "<td style='height:10mm;' class='Donate_Amt' align='right'>&nbsp;</td>"&vbcrlf
            If ((Row2+2) Mod 2)=1 Then
              Response.Write "<td style='height:10mm;' class='Donate_Amt' align='center' style='border-right: 1px solid #000000'>&nbsp;</td>"&vbcrlf
            Else
              Response.Write "<td style='height:10mm;' class='Donate_Amt' align='center'>&nbsp;</td>"&vbcrlf
            End If
            If ((Row2+2) Mod 2)=0 Then Response.Write "</tr>"&vbcrlff
          Wend
          Response.Write "      </table>"&vbcrlf
          Response.Write "    </td>"&vbcrlf
          Response.Write "  </tr>"&vbcrlf
          Response.Write "</table>"&vbcrlf
          If Page=1 Or Page=2 Then Response.Write "<div class='pagebreak'>&nbsp;</div>"  
        Next

        SQL2="Select Donate_Date=CONVERT(VarChar,Donate_Date,111),Donate_Amt=Isnull(Donate_Amt,0),Donate_Payment,Donate_Purpose,Invoice_Pre,Invoice_No,Comment=CONVERT(nvarchar(100), Comment),Invoice_No,Invoice_Print,Invoice_Print_Yearly From Donate Join DONOR On Donate.Donor_Id=DONOR.Donor_Id "&Donate_Where&" And Donate.Donor_Id='"&RS1("Donor_Id")&"' And Donate.Invoice_Title='"&RS1("Donor_Name")&"' Order By Donate_Date,Donate_Id"
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,3
        While Not RS2.EOF
          If RS2("Invoice_No")="" Then
            RS2("Invoice_Pre")=Invoice_Pre
            RS2("Invoice_No")=Invoice_No
          End If
          RS2("Invoice_Print")="1"
          RS2("Invoice_Print_Yearly")="1"
          RS2.Update
          RS2.MoveNext
        Wend
        RS2.Close
        Set RS2=Nothing

        If Row<RS1.Recordcount Then Response.Write "<div class='pagebreak'>&nbsp;</div>"  
        Row=Row+1
        Response.Flush
        Response.Clear
        RS1.MoveNext
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
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->