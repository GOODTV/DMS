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
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>套版客制單筆收據列印</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
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
    .Donor_Name{
      font-size:14;
      font-family:"標楷體";
    }
    .Invoice_No{
      font-size:19;
      font-family:"標楷體";
    }
    .Remark{
      font-size:15;
      font-family:"標楷體";
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
      Sale_Subject=""
      Sale_Content=""
      SQL2="Select TOP 1 * From DONATE_SALE Where ('"&Date()&"' Between Sale_BeginDate And Sale_EndDate) Order By Sale_BeginDate Desc,Sale_EndDate Desc,Ser_No "
      Set RS2 = Server.CreateObject("ADODB.RecordSet")
      RS2.Open SQL2,Conn,1,1
      If Not RS2.EOF Then
        Sale_Subject=RS2("Sale_Subject")
        Sale_Content=RS2("Sale_Content")
        If Sale_Content<>"" Then Sale_Content=Replace(Sale_Content,vbcrlf,"<br>")
      End If
      RS2.Close
      Set RS2=Nothing
      
      Row=0
      While Not RS1.EOF
        Goods_Name=""
        SQL2="Select Goods_Name,Goods_Qty,Goods_Unit From CONTRIBUTEDATA Where Contribute_Id='"&RS1("Contribute_Id")&"' Order By Ser_No"
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,1
        If Not RS2.EOF Then
          While Not RS2.EOF
            If Goods_Name="" Then
              Goods_Name=RS2("Goods_Name")&"("&FormatNumber(RS2("Goods_Qty"),0)&RS2("Goods_Unit")&")"
            Else
              Goods_Name=Goods_Name&"、"&RS2("Goods_Name")&"("&FormatNumber(RS2("Goods_Qty"),0)&RS2("Goods_Unit")&")"
            End If
            RS2.MoveNext
          Wend
          Goods_Name=Goods_Name&"。"  
        End If
        RS2.Close
        Set RS2=Nothing
        If RS1("Invoice_PrintComment")<>"" Then 
          If Goods_Name="" Then 
            Goods_Name=RS1("Invoice_PrintComment")
          Else
            Goods_Name=Goods_Name&"<br />"&RS1("Invoice_PrintComment")
          End If
        End If
        Row=Row+1
        Response.Write "<table id='grid' border='0' align='center' valign='top' cellpadding='0' cellspacing='0' class='TableGrid'>"&vbcrlf
        Response.Write "   <tr><td style='height:31mm'></td></tr>"&vbcrlf
        Response.Write "   <tr>"&vbcrlf
        Response.Write "     <td align='center' valign='top'>"&vbcrlf
        Response.Write "       <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:10mm;width:51mm'></td>"&vbcrlf
        Response.Write "           <td style='height:10mm;width:50mm' align='left' valign='top' class='Donor_Name'>"&RS1("Post_Donor_Name")&"</td>"&vbcrlf
        Response.Write "           <td style='height:10mm;width:93mm'></td>"&vbcrlf
        Response.Write "         </tr>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:4mm;width51mm'></td>"&vbcrlf
        Response.Write "           <td style='height:4mm;width:50mm' align='left' valign='bottom' class='ZipCode'> </td>"&vbcrlf
        Response.Write "           <td style='height:4mm;width:93mm'></td>"&vbcrlf
        Response.Write "         </tr>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:19mm;width:51mm'></td>"&vbcrlf
        Response.Write "           <td style='height:19mm;width:50mm' align='left' valign='top' class='Donor_Name'>"&RS1("ZipCode")&"<br />"&RS1("Address")&"</td>"&vbcrlf
        Response.Write "           <td style='height:19mm;width:93mm'></td>"&vbcrlf
        Response.Write "         </tr>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:8mm;width:41mm'></td>"&vbcrlf
        Response.Write "           <td style='height:8mm;width:50mm' align='left' valign='top' class='Donor_Name'>"&RS1("TEL")&"</td>"&vbcrlf
        Response.Write "           <td style='height:8mm;width:93mm'></td>"&vbcrlf
        Response.Write "         </tr>"&vbcrlf
        Response.Write "       </table>"&vbcrlf
        Response.Write "     </td>"&vbcrlf
        Response.Write "   </tr>"&vbcrlf
        Response.Write "   <tr><td style='height:58mm'></td></tr>"&vbcrlf
        Response.Write "   <tr>"&vbcrlf
        Response.Write "     <td align='center' valign='top'>"&vbcrlf
        Response.Write "       <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:12mm;width:26mm'></td>"&vbcrlf
        Response.Write "           <td style='height:12mm;width:168mm' align='left' valign='bottom' class='Invoice_No'>"&RS1("Invoice_No")&"</td>"&vbcrlf
        Response.Write "         </tr>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:11.5mm;width:26mm'></td>"&vbcrlf
        Response.Write "           <td style='height:11.5mm;width:168mm' align='left' valign='bottom' class='Invoice_No'>"&RS1("Donor_Name")&"</td>"&vbcrlf
        Response.Write "         </tr>"&vbcrlf
        Response.Write "       </table>"&vbcrlf
        Response.Write "     </td>"&vbcrlf
        Response.Write "   </tr>"&vbcrlf
        Response.Write "   <tr>"&vbcrlf
        Response.Write "     <td align='center' valign='top'>"&vbcrlf
        Response.Write "       <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:12mm;width:54mm'></td>"&vbcrlf
        Response.Write "           <td style='height:12mm;width:140mm' align='left' valign='bottom' class='Invoice_No'>"&RS1("Invoice_IDNo")&"</td>"&vbcrlf
        Response.Write "         </tr>"&vbcrlf
        Response.Write "       </table>"&vbcrlf
        Response.Write "     </td>"&vbcrlf
        Response.Write "   </tr>"&vbcrlf
        Response.Write "   <tr>"&vbcrlf
        Response.Write "     <td align='center' valign='top'>"&vbcrlf
        Response.Write "       <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:12mm;width:26mm'></td>"&vbcrlf
        If Csng(RS1("Contribute_Amt"))>0 Then
          Response.Write "         <td style='height:12mm;width:168mm' align='left' valign='bottom' class='Invoice_No'>折合現金新台幣"&ChineseMoney(RS1("Contribute_Amt"))&"&nbsp;(NT$"&FormatNumber(RS1("Contribute_Amt"),0)&")</td>"&vbcrlf
        Else
          Response.Write "         <td style='height:12mm;width:168mm' align='left' valign='bottom' class='Invoice_No'></td>"&vbcrlf
        End If
        Response.Write "         </tr>"&vbcrlf
        Response.Write "       </table>"&vbcrlf
        Response.Write "     </td>"&vbcrlf
        Response.Write "   </tr>"&vbcrlf
        Response.Write "   <tr>"&vbcrlf
        Response.Write "     <td align='center' valign='top'>"&vbcrlf
        Response.Write "       <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:12mm;width:26mm'></td>"&vbcrlf
        Response.Write "           <td style='height:12mm;width:168mm' align='left' valign='bottom' class='Invoice_No'>"&RS1("Contribute_Payment")&"</td>"&vbcrlf
        Response.Write "         </tr>"&vbcrlf
        Response.Write "       </table>"&vbcrlf
        Response.Write "     </td>"&vbcrlf
        Response.Write "   </tr>"&vbcrlf
        Response.Write "   <tr>"&vbcrlf
        Response.Write "     <td align='center' valign='top'>"&vbcrlf
        Response.Write "       <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:13.5mm;width:5mm'></td>"&vbcrlf
        Response.Write "           <td style='height:13.5mm;width:189mm' align='left' valign='top' class='Remark'>"&Goods_Name&"</td>"&vbcrlf
        Response.Write "         </tr>"&vbcrlf
        Response.Write "       </table>"&vbcrlf
        Response.Write "       <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:4mm;width:161mm'></td>"&vbcrlf
        Response.Write "           <td style='height:4mm;width:33mm' align='left' valign='bottom' class='Invoice_No'>"&RS1("Contribute_Date")&"</td>"&vbcrlf
        Response.Write "         </tr>"&vbcrlf
        Response.Write "       </table>"&vbcrlf
        Response.Write "     </td>"&vbcrlf
        Response.Write "   </tr>"&vbcrlf
        Response.Write "   <tr><td style='height:9mm'></td></tr>"&vbcrlf
        Response.Write "   <tr>"&vbcrlf
        Response.Write "     <td align='center' valign='top'>"&vbcrlf
        Response.Write "       <table width='100%' border='0' cellpadding='0' cellspacing='0'>"&vbcrlf
        Response.Write "         <tr>"&vbcrlf
        Response.Write "           <td style='height:60mm;width:12mm'></td>"&vbcrlf
        Response.Write "           <td style='height:60mm;width:90mm' align='left' valign='center' class='Invoice_No'>"&RS1("Invoice_ZipCode")&"<br />"&RS1("Invoice_Address")&"<br /><br />"&RS1("Donor_Name")&"&nbsp;"&RS1("Title2")&"啟</td>"&vbcrlf
        Response.Write "           <td style='height:60mm;width:5mm'></td>"&vbcrlf
        Response.Write "           <td style='height:60mm;width:87mm' valign='top' class='Donor_Name'><br />"&Sale_Subject&"<br/>"&Sale_Content&"</td>"&vbcrlf
        Response.Write "         </tr>"&vbcrlf
        Response.Write "       </table>"&vbcrlf
        Response.Write "     </td>"&vbcrlf
        Response.Write "   </tr>"&vbcrlf
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
    Session.Contents.Remove("SQL1")
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->