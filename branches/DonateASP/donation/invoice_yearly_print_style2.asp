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
    table.TableGrid1{
	    width:194mm;
    }
    table.TableGrid2{
	    width:194mm;
    }
    table.TableGrid3{
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
    .Line{
      border-bottom: 1px dashed #000000;
      height:3mm;
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
<body class=tool <%If Not RS1.EOF And Session("action")="report" Then%>onload='print();'<%End If%>>	
  <p><div align="center"><center>
  <%
    If RS1.EOF Then
      Response.Write "<br /><br /><br /><font size=3>沒有符合條件的資料可以列印!!</font>"
    Else
      Dept_Id_Old=""
      Row=1
      While Not RS1.EOF
        If Cstr(RS1("Dept_Id"))<>Dept_Id_Old Then
          SQL2="Select *,Dept_Address=Zip_Code+A.mValue+B.mValue+Address,Dept_Address2=Zip_Code2+C.mValue+D.mValue+Address2 From Dept " & _
               "Left Join CodeCity As A On Dept.City_Code=A.mCode " & _
               "Left Join CodeCity As B On Dept.Area_Code=B.mCode " & _
               "Left Join CodeCity As C On Dept.City_Code2=C.mCode " & _
               "Left Join CodeCity As D On Dept.Area_Code2=D.mCode Where Dept_Id='"&RS1("Dept_Id")&"' "
          Set RS2 = Server.CreateObject("ADODB.RecordSet")
          RS2.Open SQL2,Conn,1,1
          If Not RS2.EOF Then  
            Dept_Id_Old=RS2("Dept_Id")
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
        End If
        
        '取捐款編號
        If RS1("Invoice_No")="" Then
          Invoice_Pre=Invoice_Pre
          Invoice_No=Get_Invoice_NoY("1",Session("DeptId"),Session("Donate_Year")&"/12/31",Invoice_Rule_Type2,Session("DonateDesc"),Invoice_Rule_Type,Invoice_Rule_YMD,Invoice_Rule_Len,Invoice_Rule_Pub)
        Else
          Invoice_Pre=RS1("Invoice_Pre")
          Invoice_No=RS1("Invoice_No") 
        End If

        '收據抬頭改印
        If Session("Invoice_Title_New")<>"" Then
          Donor_Name=Session("Invoice_Title_New")
        Else
          Donor_Name=RS1("Donor_Name")
        End If

        Response.Write "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        For P=1 To 3
          Response.Write "<tr>"&vbcrlf
          Response.Write "  <td>"&vbcrlf
          Response.Write "    <table id='grid1' border='0' cellpadding='0' cellspacing='0' class='TableGrid"&P&"'>"&vbcrlf
          Response.Write "      <tr>"&vbcrlf
          Response.Write "        <td align='center' valign='top' colspan='3' class='Comp_Name'><u>"&Comp_Name&"收據</u></td>"&vbcrlf
          Response.Write "      </tr>"&vbcrlf
          Response.Write "      <tr>"&vbcrlf
          Response.Write "        <td align='left' valign='top' class='Invoice_No'>捐款年度:"&Session("Donate_Year")&"年匯整</td>"&vbcrlf
          Response.Write "        <td align='right' valign='top' class='Invoice_No'>收據編號:"&Invoice_Pre&Invoice_No&"</td>"&vbcrlf
          Response.Write "        <td align='center' valign='top' class='Invoice_No'></td>"&vbcrlf
          Response.Write "      </tr>"&vbcrlf
          Response.Write "      <tr>"&vbcrlf
          Response.Write "        <td align='center' valign='top' style='height:24mm' colspan='2'>"&vbcrlf
          Response.Write "          <table width='100%' border='1' cellpadding='0' cellspacing='0' bordercolor='#000000' style='border-collapse: collapse'>"&vbcrlf
          Response.Write "            <tr>"&vbcrlf
          Response.Write "              <td style='width:25mm;height:8mm' class='Donate_Desc' align='center'>捐款者</td>"&vbcrlf
          Response.Write "              <td style='width:110mm;height:8mm' class='Donate_Desc'>&nbsp;"&Data_Minus(Donor_Name)&"</td>"&vbcrlf
          Response.Write "              <td style='width:30mm;height:8mm' class='Donate_Desc' align='center'>身份證/統編</td>"&vbcrlf
          Response.Write "              <td style='width:29mm;height:8mm' class='Donate_Desc'>&nbsp;"&RS1("Invoice_IDNo")&"</td>"&vbcrlf
          Response.Write "            </tr>"&vbcrlf
          Response.Write "            <tr>"&vbcrlf
          Response.Write "              <td style='width:25mm;height:8mm' align='center' class='Donate_Desc'>收據地址</td>"&vbcrlf
          Response.Write "              <td style='width:169mm;height:8mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;"&Data_Minus(RS1("Invoice_Address1"))&"</td>"&vbcrlf
          Response.Write "            </tr>"&vbcrlf
          Response.Write "            <tr>"&vbcrlf
          Response.Write "            <td style='width:25mm;height:8mm' align='center' class='Donate_Desc'>捐款金額</td>"&vbcrlf
          Response.Write "            <td style='width:169mm;height:8mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;"&ChineseMoney(RS1("Donate_Total"))&"&nbsp;(NT$"&FormatNumber(RS1("Donate_Total"),0)&")"&vbcrlf
          Response.Write "            </tr>"&vbcrlf
          Response.Write "          </table>"&vbcrlf
          Response.Write "        </td>"&vbcrlf
          If P=1 Then
            Response.Write "      <td class='Donate_Right' align='center' valign='center'>第<br />一<br />聯<br />：<br>捐<br>助<br>人</td>"&vbcrlf
          ElseIf P=2 Then
            Response.Write "      <td class='Donate_Right' align='center' valign='center'>第<br />二<br />聯<br />：<br>會<br>計<br>聯</td>"&vbcrlf
          ElseIf P=3 Then
            Response.Write "      <td class='Donate_Right' align='center' valign='center'>第<br />三<br />聯<br />：<br>存<br>根<br>聯</td>"&vbcrlf  
          End If
          Response.Write "      </tr>"&vbcrlf
          Response.Write "      <tr>"&vbcrlf
          Response.Write "        <td align='center' valign='top' style='height:10mm' colspan='3'>"&vbcrlf
          Response.Write "          <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
          Response.Write "            <tr>"&vbcrlf
          Response.Write "              <td style='width:16mm;height:10mm' align='left' class='Donate_Seal'>執行長:</td>"&vbcrlf
          Response.Write "              <td style='width:50mm;height:10mm' align='left' class='Donate_Seal'></td>"&vbcrlf
          Response.Write "              <td style='width:12mm;height:10mm' align='left' class='Donate_Seal'>會計:</td>"&vbcrlf
          Response.Write "              <td style='width:50mm;height:10mm' align='left' class='Donate_Seal'></td>"&vbcrlf
          Response.Write "              <td style='width:19mm;height:10mm' align='left' class='Donate_Seal'>經手人:</td>"&vbcrlf
	        Response.Write "              <td style='width:47mm;height:10mm' align='left' class='Donate_Seal'></td>"&vbcrlf
          Response.Write "            </tr>"&vbcrlf
          Response.Write "          </table>"&vbcrlf
          Response.Write "        </td>"&vbcrlf
          Response.Write "      </tr>"&vbcrlf
          If P=1 Then
            Response.Write "    <tr>"&vbcrlf
            Response.Write "      <td align='center' valign='top' style='height:18mm' colspan='3'>"&vbcrlf
            Response.Write "        <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
            Response.Write "          <tr>"&vbcrlf
            Response.Write "            <td style='width:64mm;height:8mm' align='left' class='Donate_Account'>劃撥帳號:"&Account&"</td>"&vbcrlf
            Response.Write "            <td style='width:130mm;height:8mm' align='left' class='Donate_Account' colspan='2'>戶名:"&Comp_Name&"</td>"&vbcrlf
            Response.Write "          </tr>"&vbcrlf
            Response.Write "          <tr>"&vbcrlf
            Response.Write "            <td style='width:64mm;height:5mm' align='left' class='Donate_Foot'>電話:"&TEL&"</td>"&vbcrlf
            Response.Write "            <td style='width:55mm;height:5mm' align='left' class='Donate_Foot'>傳真:"&Fax&"</td>"&vbcrlf
            Response.Write "            <td style='width:75mm;height:5mm' align='left' class='Donate_Foot'>地址:"&Dept_Address&"</td>"&vbcrlf
            Response.Write "          </tr>"&vbcrlf
            Response.Write "          <tr>"&vbcrlf
            Response.Write "            <td style='width:64mm;height:5mm' align='left' class='Donate_Foot'>EMail:"&EMail&"</td>"&vbcrlf
            Response.Write "            <td style='width:55mm;height:5mm' align='left' class='Donate_Foot'>統一編號:"&Uniform_No&"</td>"&vbcrlf
            Response.Write "            <td style='width:75mm;height:5mm' align='left' class='Donate_Foot'>立案字號:北府社老字第1001825306號</td>"&vbcrlf
            Response.Write "          </tr>"&vbcrlf
            Response.Write "        </table>"&vbcrlf
            Response.Write "      </td>"&vbcrlf
            Response.Write "    </tr>"&vbcrl
            Response.Write "    <tr>"&vbcrlf
            Response.Write "      <td style='height:16mm'></td>"&vbcrlf
            Response.Write "    </tr>"&vbcrlf
          Else  
            Response.Write "    <tr>"&vbcrlf
            Response.Write "      <td style='height:32mm'></td>"&vbcrlf
            Response.Write "    </tr>"&vbcrlf
          End If
          Response.Write "    </table>"&vbcrlf
          Response.Write "  </td>"&vbcrlf
          Response.Write "</tr>"&vbcrlf
          If P<3 Then
            Response.Write "<tr>"&vbcrlf
          	Response.Write "  <td class='Line'> </td>"&vbcrlf
            Response.Write "</tr>"&vbcrlf
            Response.Write "<tr>"&vbcrlf
          	Response.Write "  <td class='CellMiddle'> </td>"&vbcrlf
            Response.Write "</tr>"&vbcrlf
          End If
        Next
        Response.Write "</table>"&vbcrlf	

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