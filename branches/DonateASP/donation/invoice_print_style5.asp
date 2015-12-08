<!--#include file="../include/dbfunctionJ.asp"-->
<%
'邊界:左:8 右:8 上:8 下:8
SQL1=Session("SQL1")
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,3
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>套版客制單筆收據列印</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
   <object id="factory" viewastext style="display:none" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="http://mis.npois.com.tw/smsx.cab"></object>
  <script language="javascript">
    function window.onload(){
      factory.printing.header='' ;          //頁首
      factory.printing.footer='' ;          //頁尾
      factory.printing.portrait = false ;   //直印(true),橫印(false)
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
	    width:194mm;
    }
    .Sale_Content{
      font-size:18;
      font-family:"標楷體";
    }    
    .Invoice_No{
      font-size:16;
      font-family:"標楷體";
    }
    .Post_Donor_Name{
      font-size:15;
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
<body class=tool <%If Not RS1.EOF Then%>onload='print();'<%End If%>>	
  <p><div align="center"><center>
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
        Donate_Date_Year=Cstr(Year(RS1("Donate_Date"))-1911)
        Donate_Date_Month=Cstr(Month(RS1("Donate_Date")))
        If Len(Donate_Date_Month)=1 Then Donate_Date_Month="0"&Donate_Date_Month
        Donate_Date_Day=Cstr(Day(RS1("Donate_Date")))
        If Len(Donate_Date_Day)=1 Then Donate_Date_Day="0"&Donate_Date_Day
        Donate_Date="中華民國"&Donate_Date_Year&"年"&Donate_Date_Month&"月"&Donate_Date_Day&"日"
        
        Response.Write "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        Response.Write "  <tr><td style='height:7mm'></td></tr>"&vbcrlf
        Response.Write "  <tr>"&vbcrlf
        Response.Write "    <td>"&vbcrlf
        Response.Write "      <table id='grid_1' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        Response.Write "        <tr>"&vbcrlf
        Response.Write "          <td style='width:2mm;height:82mm'></td>"&vbcrlf
        Response.Write "          <td style='width:192mm;height:82mm' valign='top' class='Sale_Content'>"&Sale_Content&"</td>"&vbcrlf
        Response.Write "        </tr>"&vbcrlf
        Response.Write "      </table>"&vbcrlf
        Response.Write "    </td>"&vbcrlf
        Response.Write "  </tr>"&vbcrlf
        
        Response.Write "  <tr><td style='height:19mm'></td></tr>"&vbcrlf	
        Response.Write "  <tr>"&vbcrlf
        Response.Write "    <td>"&vbcrlf
        Response.Write "      <table id='grid_2' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        Response.Write "        <tr>"&vbcrlf
        Response.Write "          <td style='width:82mm;height:8mm'></td>"&vbcrlf
        Response.Write "          <td style='width:112mm;height:8mm' valign='center' class='Invoice_No'>"&RS1("Invoice_No")&"</td>"&vbcrlf
        Response.Write "        </tr>"&vbcrlf
        Response.Write "        <tr>"&vbcrlf
        Response.Write "          <td style='width:82mm;height:8mm'></td>"&vbcrlf
        Response.Write "          <td style='width:112mm;height:8mm' valign='center' class='Invoice_No'>"&Donate_Date&"</td>"&vbcrlf
        Response.Write "        </tr>"&vbcrlf
        Response.Write "        <tr>"&vbcrlf
        Response.Write "          <td style='width:82mm;height:8mm'></td>"&vbcrlf
        Response.Write "          <td style='width:112mm;height:8mm' valign='center' class='Invoice_No'>"&RS1("Donor_Name")&"</td>"&vbcrlf
        Response.Write "        </tr>"&vbcrlf
        Response.Write "        <tr>"&vbcrlf
        Response.Write "          <td style='width:82mm;height:8mm'></td>"&vbcrlf
        Response.Write "          <td style='width:112mm;height:8mm' valign='center' class='Invoice_No'>新台幣"&ChineseMoney(RS1("Donate_Amt"))&"&nbsp;(NT$"&FormatNumber(RS1("Donate_Amt"),0)&")"&vbcrlf
        Response.Write "        </tr>"&vbcrlf
        Response.Write "      </table>"&vbcrlf
        Response.Write "    </td>"&vbcrlf
        Response.Write "  </tr>"&vbcrlf

        Response.Write "  <tr><td style='height:73mm'></td></tr>"&vbcrlf	
        Response.Write "  <tr>"&vbcrlf
        Response.Write "    <td>"&vbcrlf
        Response.Write "      <table id='grid_3' border='1' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        Response.Write "        <tr>"&vbcrlf
        Response.Write "          <td style='width:52.5mm;height:10mm'></td>"&vbcrlf
        Response.Write "          <td style='width:46mm;height:10mm' valign='center' class='Post_Donor_Name'>"&RS1("Post_Donor_Name")&"</td>"&vbcrlf
        Response.Write "          <td style='width:95.5mm;height:10mm'></td>"&vbcrlf
        Response.Write "        </tr>"&vbcrlf
        Response.Write "        <tr>"&vbcrlf
        Response.Write "          <td style='width:194mm;height:5mm' colspan='3'></td>"&vbcrlf
        Response.Write "        </tr>"&vbcrlf
        Response.Write "        <tr>"&vbcrlf
        Response.Write "          <td style='width:52.5mm;height:21.5mm'></td>"&vbcrlf
        Response.Write "          <td style='width:46mm;height:21.5mm' valign='top' class='Post_Donor_Name'>"&RS1("ZipCode")&"<br />"&RS1("Address")&"</td>"&vbcrlf
        Response.Write "          <td style='width:95.5mm;height:21.5mm'></td>"&vbcrlf
        Response.Write "        </tr>"&vbcrlf
        Response.Write "        <tr>"&vbcrlf
        Response.Write "          <td style='width:52.5mm;height:8mm'></td>"&vbcrlf
        Response.Write "          <td style='width:46mm;height:8mm' valign='center' class='Post_Donor_Name'>"&RS1("TEL")&"</td>"&vbcrlf
        Response.Write "          <td style='width:95.5mm;height:8mm'></td>"&vbcrlf
        Response.Write "        </tr>"&vbcrlf
        Response.Write "      </table>"&vbcrlf
        Response.Write "    </td>"&vbcrlf
        Response.Write "  </tr>"&vbcrlf
        Response.Write "</table>"&vbcrlf
        
        RS1("Invoice_Print")="1"
        RS1.Update
        Response.Flush
        Response.Clear
        RS1.MoveNext
        If Not RS1.EOF Then Response.Write "<div class='pagebreak'>&nbsp;</div>"
      Wend
      RS1.Close
      Set RS1=Nothing
       
      Session.Contents.Remove("Rept_Licence")
      Session.Contents.Remove("Print_Desc")
      Session.Contents.Remove("SQL1")
      Response.End
    End If
    Session.Contents.Remove("Rept_Licence")
    Session.Contents.Remove("Print_Desc")
    Session.Contents.Remove("SQL1")
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->