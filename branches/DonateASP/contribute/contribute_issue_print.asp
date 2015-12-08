<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL1="Select * From CONTRIBUTE_ISSUE Where Issue_Id='"&request("issue_id")&"'"
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,3
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <title>物品領用列印</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <style>
  .issue{
    FONT-FAMILY: 標楷體;
    FONT-SIZE: 12pt;
  }
  .PageBreak { page-break-after:always; }
  </style>  	
</head>
<body class=tool <%If Not RS1.EOF Then%>onload='print();'<%End If%>>
  <p><div align="center"><center>
  <%
    If RS1.EOF Then
      Response.Write "<br /><br /><br /><font size=3>沒有符合條件的資料可以列印!!</font>"
    Else
      Max_Row=30
      Row=0
      While Not RS1.EOF
        Row=Row+1
        Response.Write "<table id='grid1' width='720' valign='top' align='center' border='0' cellpadding='4' cellspacing='0' class='TableGrid'>"&vbcrlf
		    Response.Write "  <tr>"&vbcrlf
		    Response.Write "	  <td align='center' colspan='2'><span style='font-size: 14pt;font-family: 標楷體'>"&Session("comp_ShortName")&"物品領用單</span></td>"&vbcrlf
		    Response.Write "  </tr>"&vbcrlf
		    Response.Write "  <tr>"&vbcrlf
		    Response.Write "	  <td width='420' class='issue'>領用編號："&RS1("Issue_Pre")&RS1("Issue_No")&"</td>"&vbcrlf
		    Response.Write "	  <td width='300' class='issue'>領用日期："&Year(RS1("Issue_Date"))&"年"&Month(RS1("Issue_Date"))&"月"&Day(RS1("Issue_Date"))&"日</td>"&vbcrlf
		    Response.Write "  </tr>"&vbcrlf
		    Response.Write "  <tr>"&vbcrlf
		    Response.Write "	  <td class='issue'>領用用途："&RS1("Issue_Purpose")&"</td>"&vbcrlf
		    Response.Write "	  <td class='issue'>出貨單位："&RS1("Issue_Org")&"</td>"&vbcrlf
		    Response.Write "  </tr>"&vbcrlf
		    Response.Write "  <tr>"&vbcrlf
		    Response.Write "	  <td class='issue'>&nbsp;</td>"&vbcrlf		
		    Response.Write "	  <td class='issue'>列印時間："&Now()&"</td>"&vbcrlf
		    Response.Write "  </tr>"&vbcrlf
	      Response.Write "</table>"&vbcrlf
        Response.Write "<table id='grid2' id='grid1' width='720' valign='top' align='center' border='2' bordercolor='#111111' cellspacing='0' cellpadding='1'>"&vbcrlf
        Response.Write "  <tr>"&vbcrlf
        Response.Write "    <td>"&vbcrlf
        Response.Write "      <table border='1' width='720' id='table1' bordercolor='#808080' cellspacing='0' cellpadding='2' class='fonts'>"&vbcrlf
	      Response.Write "        <tr>"&vbcrlf
	      Response.Write "	        <td width='80' height='26' align='center' class='issue' bgcolor='#E5ECFF'>物品代號</td>"&vbcrlf
	      Response.Write "	        <td width='340' height='26' align='center' class='issue' bgcolor='#E5ECFF'>物品名稱</td>"&vbcrlf
	      Response.Write "	        <td width='100' height='26' align='center' class='issue' bgcolor='#E5ECFF'>數量單位</td>"&vbcrlf
	      Response.Write "	        <td width='200' height='26' align='center' class='issue' bgcolor='#E5ECFF'>備註</td>"&vbcrlf
	      Response.Write "        </tr>"&vbcrlf
	      Row2=0
        SQL2="Select * From CONTRIBUTE_ISSUEDATA Where Issue_Id='"&request("issue_id")&"' Order By Ser_No"
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,1
        While Not RS2.EOF
          Row2=Row2+1
	        Response.Write "<tr>"&vbcrlf
	        Response.Write "  <td height='26' class='issue'>"&RS2("Goods_Id")&"</td>"&vbcrlf
	        Response.Write "  <td height='26' class='issue'>"&RS2("Goods_Name")&"</td>"&vbcrlf
	        Response.Write "  <td height='26' class='issue'>"&RS2("Goods_Qty")&RS2("Goods_Unit")&"</td>"&vbcrlf
	        Response.Write "  <td height='26' class='issue'>"&RS2("Goods_Comment")&"</td>"&vbcrlf
	        Response.Write "</tr>"&vbcrlf
	        If (Row2+Max_Row) Mod Max_Row=0 And Row2<RS2.Recordcount Then
	          Response.Write "      </table>"&vbcrlf
	          Response.Write "    </td>"&vbcrlf
	          Response.Write "  </tr>"&vbcrlf
	          Response.Write "</table>"&vbcrlf
            Response.Write "<table id='grid3' id='grid1' width='720' valign='top' align='center' border='0' cellpadding='4' cellspacing='0' class='TableGrid'>"&vbcrlf
		        Response.Write "  <tr>"&vbcrlf
		        Response.Write "	  <td height='26 class='issue' colspan='3'>&nbsp;</td>"&vbcrlf
		        Response.Write "  </tr>"&vbcrlf
		        Response.Write "  <tr>"&vbcrlf
		        Response.Write "	  <td width='240' class='issue'>取貨人：</td>"&vbcrlf
		        Response.Write "	  <td width='240' class='issue'>領取人：</td>"&vbcrlf
		        Response.Write "	  <td width='240' class='issue'>主管簽核：</td>"&vbcrlf
		        Response.Write "  </tr>"&vbcrlf
	          Response.Write "</table>"&vbcrlf
            Response.Write "<div class='pagebreak'>&nbsp;</div>"
            Response.Write "<table id='grid1' width='720' valign='top' align='center' border='0' cellpadding='4' cellspacing='0' class='TableGrid'>"&vbcrlf
		        Response.Write "  <tr>"&vbcrlf
		        Response.Write "	  <td align='center' colspan='2'><span style='font-size: 14pt;font-family: 標楷體'>"&Session("comp_ShortName")&"物品領用單</span></td>"&vbcrlf
		        Response.Write "  </tr>"&vbcrlf
		        Response.Write "  <tr>"&vbcrlf
		        Response.Write "	  <td width='420' class='issue'>領用編號："&RS1("Issue_Pre")&RS1("Issue_No")&"</td>"&vbcrlf
		        Response.Write "	  <td width='300' class='issue'>領用日期："&RS1("Issue_Date")&"</td>"&vbcrlf
		        Response.Write "  </tr>"&vbcrlf
		        Response.Write "  <tr>"&vbcrlf
		        Response.Write "	  <td class='issue'>領用用途："&RS1("Issue_Purpose")&"</td>"&vbcrlf
		        Response.Write "	  <td class='issue'>出貨單位："&RS1("Issue_Org")&"</td>"&vbcrlf
		        Response.Write "  </tr>"&vbcrlf
		        Response.Write "  <tr>"&vbcrlf
		        Response.Write "	  <td class='issue'>&nbsp;</td>"&vbcrlf	
		        Response.Write "	  <td class='issue'>列印時間："&Now()&"</td>"&vbcrlf
		        Response.Write "  </tr>"&vbcrlf
	          Response.Write "</table>"&vbcrlf
            Response.Write "<table id='grid2' id='grid1' width='720' valign='top' align='center' border='2' bordercolor='#111111' cellspacing='0' cellpadding='1'>"&vbcrlf
            Response.Write "  <tr>"&vbcrlf
            Response.Write "    <td>"&vbcrlf
            Response.Write "      <table border='1' width='720' id='table1' bordercolor='#808080' cellspacing='0' cellpadding='2' class='fonts'>"&vbcrlf
	          Response.Write "        <tr>"&vbcrlf
	          Response.Write "	        <td width='80' height='26' align='center' class='issue' bgcolor='#E5ECFF'>物品代號</td>"&vbcrlf
	          Response.Write "	        <td width='340' height='26' align='center' class='issue' bgcolor='#E5ECFF'>物品名稱</td>"&vbcrlf
	          Response.Write "	        <td width='100' height='26' align='center' class='issue' bgcolor='#E5ECFF'>數量單位</td>"&vbcrlf
	          Response.Write "	        <td width='200' height='26' align='center' class='issue' bgcolor='#E5ECFF'>備註</td>"&vbcrlf
	          Response.Write "        </tr>"&vbcrlf
	        End If
          RS2.MoveNext
        Wend
        RS2.Close
        Set RS2=Nothing
        While (Row2+Max_Row) Mod Max_Row<>0
          Row2=Row2+1
	        Response.Write "<tr>"&vbcrlf
	        Response.Write "  <td height='26' class='issue_date'>&nbsp;</td>"&vbcrlf
	        Response.Write "  <td height='26' class='issue_date'>&nbsp;</td>"&vbcrlf
	        Response.Write "  <td height='26' class='issue_date'>&nbsp;</td>"&vbcrlf
	        Response.Write "  <td height='26' class='issue_date'>&nbsp;</td>"&vbcrlf
	        Response.Write "</tr>"&vbcrlf
	      Wend
	      Response.Write "      </table>"&vbcrlf
	      Response.Write "    </td>"&vbcrlf
	      Response.Write "  </tr>"&vbcrlf
	      Response.Write "</table>"&vbcrlf
        Response.Write "<table id='grid3' id='grid1' width='720' valign='top' align='center' border='0' cellpadding='4' cellspacing='0' class='TableGrid'>"&vbcrlf
		    Response.Write "  <tr>"&vbcrlf
		    Response.Write "	  <td height='26 class='issue' colspan='3'>&nbsp;</td>"&vbcrlf
		    Response.Write "  </tr>"&vbcrlf
		    Response.Write "  <tr>"&vbcrlf
		    Response.Write "	  <td width='240' class='issue'>取貨人：</td>"&vbcrlf
		    Response.Write "	  <td width='240' class='issue'>領取人：</td>"&vbcrlf
		    Response.Write "	  <td width='240' class='issue'>主管簽核：</td>"&vbcrlf
		    Response.Write "  </tr>"&vbcrlf
	      Response.Write "</table>"&vbcrlf

        RS1("Issue_Print")="1"
        RS1.Update
        If Row<RS1.Recordcount Then Response.Write "<div class='pagebreak'>&nbsp;</div>"
        Response.Flush
        Response.Clear
        RS1.MoveNext
      Wend
    End If
    RS1.Close
    Set RS1=Nothing    
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->