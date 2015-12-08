<!--#include file="../include/dbfunctionJ.asp"-->
<%
'邊界:左:8 右:8 上:8 下:8
SQL1=Session("SQL1")
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,3
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=ProgDesc("invoice_print")%></title>
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
	    height:135mm;
    }
    table.TableGrid2{
	    width:194mm;
	    height:125mm;
    }
    .Comp_Name{ 
	    width:194mm;
	    height:20mm;
      font-size:25;
      font-family:"標楷體";
    }
    .Invoice_No{
      height:5mm; 
      font-size:15;
      font-family:"標楷體";
    }
    .Donate_Desc{
      height:12mm;
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
      height:8mm;
      font-size:12;
      font-family:"標楷體";
    }
    .Line{
      border-bottom: 1px dashed #000000;
      height:5mm;
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
<body class=tool <%If Not RS1.EOF Then%>onload='print();'<%End If%>>	
  <p><div align="center"><center>
  <%
    If RS1.EOF Then
      Response.Write "<br /><br /><br /><font size=3>沒有符合條件的資料可以列印!!</font>"
    Else
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
              SEAL_D="../upload/"&RS2("Seal_TitleImg")
            Else
              SEAL_4="../upload/"&RS2("Seal_TitleImg")
            End If
          End If
        End If
        RS2.MoveNext
      Wend
      RS2.Close
      Set RS2=Nothing
      If SEAL_4="" Then SEAL_4=SEAL_D
      
      DonateDesc="物資捐贈"
      SQL2="Select DonateDesc=CodeDesc From CASECODE Where CodeType='Payment' And CodeDesc Like '%物%' Order By Seq"
      Call QuerySQL(SQL2,RS2)
      If Not RS2.EOF Then DonateDesc=RS2("DonateDesc")
      RS2.Close
      Set RS2=Nothing

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
        
        Response.Write "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        For P=1 To 2
          Response.Write "<tr>"&vbcrlf
          Response.Write "  <td>"&vbcrlf
          Response.Write "    <table id='grid"&P&"' border='0' cellpadding='0' cellspacing='0' class='TableGrid"&P&"'>"&vbcrlf
          Response.Write "      <tr>"&vbcrlf
          Response.Write "        <td align='center' valign='top' style='height:20mm' colspan='3' class='Comp_Name'><u>"&Comp_Name&"<br />收據</u></td>"&vbcrlf
          Response.Write "      </tr>"&vbcrlf
          Response.Write "      <tr>"&vbcrlf
          Response.Write "        <td align='left' valign='top' style='height:5mm' class='Invoice_No'>捐款日期:"&RS1("Donate_Date")&"</td>"&vbcrlf
          Response.Write "        <td align='right' valign='top' style='height:5mm' class='Invoice_No'>收據編號:"&RS1("Invoice_No")&"</td>"&vbcrlf
          Response.Write "        <td align='center' valign='top' style='height:5mm' class='Invoice_No'></td>"&vbcrlf
          Response.Write "      </tr>"&vbcrlf
          Response.Write "      <tr>"&vbcrlf
          Response.Write "        <td align='center' valign='top' style='height:60mm' colspan='2'>"&vbcrlf
          Response.Write "          <table width='100%' border='1' cellpadding='0' cellspacing='0' bordercolor='#000000' style='border-collapse: collapse'>"&vbcrlf
          Response.Write "            <tr>"&vbcrlf
          Response.Write "              <td style='width:25mm;height:12mm' class='Donate_Desc' align='center'>捐款者</td>"&vbcrlf
          Response.Write "              <td style='width:110mm;height:12mm' class='Donate_Desc'>&nbsp;"&Data_Minus(RS1("Donor_Name"))&"</td>"&vbcrlf
          Response.Write "              <td style='width:30mm;height:12mm' class='Donate_Desc' align='center'>身份證/統編</td>"&vbcrlf
          Response.Write "              <td style='width:29mm;height:12mm' class='Donate_Desc'>&nbsp;"&RS1("Invoice_IDNo")&"</td>"&vbcrlf
          Response.Write "            </tr>"&vbcrlf
          Response.Write "            <tr>"&vbcrlf
          Response.Write "              <td style='width:25mm;height:12mm' align='center' class='Donate_Desc'>收據地址</td>"&vbcrlf
          Response.Write "              <td style='width:169mm;height:12mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;"&Data_Minus(RS1("Invoice_Address2"))&"</td>"&vbcrlf
          Response.Write "            </tr>"&vbcrlf
          Response.Write "            <tr>"&vbcrlf
          If Cstr(RS1("Donate_Payment"))=DonateDesc Then
            Response.Write "            <td style='width:25mm;height:12mm' align='center' class='Donate_Desc'>物資捐贈</td>"&vbcrlf
            Response.Write "            <td style='width:169mm;height:12mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;"&RS1("Donate_Desc")&"</td>"&vbcrlf
          Else
            Response.Write "            <td style='width:25mm;height:12mm' align='center' class='Donate_Desc'>捐款金額</td>"&vbcrlf
            If RS1("Donate_Amt")<>"" Then
              Response.Write "          <td style='width:169mm;height:12mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;新台幣"&ChineseMoney(RS1("Donate_Amt"))&"</td>"&vbcrlf
            Else
              Response.Write "          <td style='width:169mm;height:12mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;</td>"&vbcrlf
            End If
          End If
          Response.Write "            </tr>"&vbcrlf
          Response.Write "            <tr>"&vbcrlf
          Response.Write "              <td style='width:25mm;height:12mm' align='center' class='Donate_Desc'>指定用途</td>"&vbcrlf
          Response.Write "              <td style='width:169mm;height:12mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;"&RS1("Donate_Purpose")&"</td>"&vbcrlf
          Response.Write "            </tr>"&vbcrlf
          Response.Write "            <tr>"&vbcrlf
          Response.Write "              <td style='width:25mm;height:12mm' align='center' class='Donate_Desc'>備註</td>"&vbcrlf
          If Cstr(RS1("Donate_Payment"))=DonateDesc Then
            If RS1("Donate_Amt2")<>"" Then
              Response.Write "          <td style='width:169mm;height:12mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;折合現金：新台幣"&ChineseMoney(RS1("Donate_Amt2"))&"</td>"&vbcrlf
            Else
              Response.Write "          <td style='width:169mm;height:12mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;</td>"&vbcrlf
            End If
          Else
            Response.Write "            <td style='width:169mm;height:12mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;"&RS1("Donate_Payment")&"</td>"&vbcrlf
          End If
          Response.Write "            </tr>"&vbcrlf
          Response.Write "          </table>"&vbcrlf
          Response.Write "        </td>"&vbcrlf
          If P=1 Then
            Response.Write "      <td class='Donate_Right' align='center' valign='center'>第<br />一<br />聯<br />：<br>捐<br>助<br>人<br>留<br>存</td>"&vbcrlf
          ElseIf P=2 Then
            Response.Write "      <td class='Donate_Right' align='center' valign='center'>第<br />二<br />聯<br />：<br>會<br>計<br>聯</td>"&vbcrlf
          End If
          Response.Write "      </tr>"&vbcrlf          	
          Response.Write "      <tr>"&vbcrlf
          Response.Write "        <td align='center' valign='top' style='height:10mm' colspan='3'>"&vbcrlf
          Response.Write "          <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
          Response.Write "            <tr>"&vbcrlf
          Response.Write "              <td style='width:16mm;height:10mm' align='left' class='Donate_Seal'>執行長:</td>"&vbcrlf
          Response.Write "              <td style='width:50mm;height:10mm' align='left' class='Donate_Seal'><img border='0' src='"&SEAL_2&"' style='height:10mm'></td>"&vbcrlf
          Response.Write "              <td style='width:12mm;height:10mm' align='left' class='Donate_Seal'>會計:</td>"&vbcrlf
          Response.Write "              <td style='width:50mm;height:10mm' align='left' class='Donate_Seal'><img border='0' src='"&SEAL_3&"' style='height:10mm'></td>"&vbcrlf
          Response.Write "              <td style='width:16mm;height:10mm' align='left' class='Donate_Seal'>經手人:</td>"&vbcrlf
	        Response.Write "              <td style='width:50mm;height:10mm' align='left' class='Donate_Seal'><img border='0' src='"&SEAL_4&"' style='height:10mm'></td>"&vbcrlf
          Response.Write "            </tr>"&vbcrlf
          Response.Write "          </table>"&vbcrlf
          Response.Write "        </td>"&vbcrlf
          Response.Write "      </tr>"&vbcrlf
          If P=1 Then
            Response.Write "    <tr>"&vbcrlf
            Response.Write "      <td align='center' valign='top' style='height:18mm' colspan='3'>"&vbcrlf
            Response.Write "        <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
            Response.Write "          <tr>"&vbcrlf
            Response.Write "            <td style='width:194mm;height:8mm' align='left' class='Donate_Account' colspan='2'>劃撥帳號:"&Account&"</td>"&vbcrlf
            Response.Write "          </tr>"&vbcrlf
            Response.Write "          <tr>"&vbcrlf
            Response.Write "            <td style='width:194mm;height:8mm' align='left' class='Donate_Account' colspan='2'>戶&nbsp;&nbsp;&nbsp;&nbsp;名:"&Comp_Name&"</td>"&vbcrlf
            Response.Write "          </tr>"&vbcrlf
            Response.Write "          <tr>"&vbcrlf
            Response.Write "            <td style='width:97mm;height:8mm' class='Donate_Foot' align='left'>電話:"&TEL&"</td>"&vbcrlf
            Response.Write "            <td style='width:97mm;height:8mm' class='Donate_Foot' align='left'>統一編號:"&Uniform_No&"</td>"&vbcrlf
            Response.Write "          </tr>"&vbcrlf
            Response.Write "          <tr>"&vbcrlf
            Response.Write "            <td style='width:97mm;height:8mm' class='Donate_Foot' align='left'>傳真:"&Fax&"</td>"&vbcrlf
            Response.Write "            <td style='width:97mm;height:8mm' class='Donate_Foot' align='left'>EMail:"&EMail&"</td>"&vbcrlf
            Response.Write "          </tr>"&vbcrlf
            Response.Write "          <tr>"&vbcrlf
            Response.Write "            <td style='width:97mm;height:8mm' class='Donate_Foot' align='left'>地址:"&Dept_Address&"</td>"&vbcrlf
            Response.Write "            <td style='width:97mm;height:8mm' class='Donate_Foot' align='left'>勸募許可文號:"&Session("Rept_Licence")&"</td>"&vbcrlf
            Response.Write "          </tr>"&vbcrlf
            Response.Write "        </table>"&vbcrlf
            Response.Write "      </td>"&vbcrlf
            Response.Write "    </tr>"&vbcrlf
          Else
            Response.Write "    <tr>"&vbcrlf
            Response.Write "      <td style='height:30mm'> </td>"&vbcrlf
            Response.Write "    </tr>"&vbcrlf 
          End If
          Response.Write "    </table>"&vbcrlf
          Response.Write "  </td>"&vbcrlf
          Response.Write "</tr>"&vbcrlf
          If P<2 Then
            Response.Write "<tr>"&vbcrlf
          	Response.Write "  <td class='Line'> </td>"&vbcrlf
            Response.Write "</tr>"&vbcrlf
            Response.Write "<tr>"&vbcrlf
          	Response.Write "  <td class='CellMiddle'> </td>"&vbcrlf
            Response.Write "</tr>"&vbcrlf
          End If
        Next
        Response.Write "</table>"&vbcrlf	
        
        If Row<RS1.Recordcount Then Response.Write "<div class='pagebreak'>&nbsp;</div>"
        RS1("Invoice_Print")="1"
        RS1.Update
        Row=Row+1
        Response.Flush
        Response.Clear
        RS1.MoveNext
      Wend
      RS1.Close
      Set RS1=Nothing
    End If
    Session.Contents.Remove("Rept_Licence")
    Session.Contents.Remove("Print_Desc")
    Session.Contents.Remove("SQL1")
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->