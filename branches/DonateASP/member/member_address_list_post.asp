<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL1=Session("SQL1")
Call QuerySQL(SQL1,RS1)
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>郵局掛號函件執據</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <style>
  <!--
   .style1 {font-size: 15px}
   .style3 {font-size: 10pt}
   .PageBreak { page-break-after:always; }
  -->
  </style>
</head>
<body class=tool <%If Not RS1.EOF Then%>onload='print();'<%End If%>>
<p><div align="center"><center>
  <%
    If RS1.EOF Then
      Response.write "<br><br><br><font size=3>沒有符合條件的資料可以列印!!</font>"
    Else
      '機構連絡資訊
      If Session("DeptId")<>"" Then 
        SQL2="Select *,Dept_Address=(Case When A.mValue=B.mValue Then Zip_Code+A.mValue+Address Else Zip_Code+A.mValue+B.mValue+Address End) From Dept " & _
             "Left Join CodeCity As A On Dept.City_Code=A.mCode " & _
             "Left Join CodeCity As B On Dept.Area_Code=B.mCode " & _
             "Where Dept_Id='"&Session("DeptId")&"' "
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,1
        If Not RS2.EOF Then
          Comp_Name=RS2("Comp_Name")
	        Dept_Address=RS2("Dept_Address")
	        Dept_Tel=RS2("TEL")
	        Dept_ContaCtor=RS2("ContaCtor")
        End If
        RS2.Close
        Set RS2=Nothing
      End If
      Response.write "<table width='720'  border='0' cellspacing='0'>"
      Response.write "  <tr>"
      Response.write "    <td width='40%' height='120' valign='bottom' colspan='2'>"
      Response.write "      <table width='100%' border='0' cellspacing='0'>"
      Response.write "        <tr>"
      Response.write "          <td height='25'><span class='style3'>中華民國&nbsp;"&Year(Date())-1911&"&nbsp;年&nbsp;"&Month(Date())&"&nbsp;月&nbsp;"&Day(Date())&"&nbsp;日</span></td>"
      Response.write "        </tr>"
      Response.write "        <tr>"
      Response.write "          <td height='25'><span class='style3'>寄件人："&Session("user_name")&"</span></td>"
      Response.write "        </tr>"
      Response.write "        <tr>"
      Response.write "          <td height='25'>名&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;稱："&Comp_Name&"</td>"
      Response.write "        </tr>"
      Response.write "      </table>"
      Response.write "    </td>"
      Response.write "    <td width='30%' height='120' align='center'>"
      Response.write "      <p class='style1'><u>中&nbsp; 華&nbsp; 民&nbsp; 國&nbsp; 郵&nbsp; 政</u></p>"
      Response.write "      <p class='style1'><u>限時掛號</u></p>"
      Response.write "      <p class='style1'> 交寄大宗<u>掛&nbsp;&nbsp;&nbsp; 號</u>函件執據</p>"
      Response.write "      <p class='style1'> <u>快捷郵件</u> </p>"
      Response.write "    </td>"
      Response.write "    <td width='30%' height='120' align='center'>"
      Response.write "      <p><img src='../images/post.gif' width='82' height='80'></p>"
      Response.write "      <p class='style3'>郵局郵戳</p>"
      Response.write "    </td>"
      Response.write "  </tr>"
      Response.write "  <tr>"
      Response.write "    <td width='21%' height='25'><span class='style3'>寄件人代表："&Dept_ContaCtor&"</span></td>"
      Response.write "    <td width='48%' height='25' colspan='2'><span class='style3'>詳細地址："&Dept_Address&"</span></td>"
      Response.write "    <td width='30%' height='25'><span class='style3'>電話號碼："&Dept_Tel&"</span></td>"
      Response.write "  </tr>"
      Response.write "</table>"
      Response.write "<table width='720'  border='1' cellspacing='0' bordercolor='#000000' style='border-collapse: collapse'>"
      Response.write "  <tr>"
      Response.write "    <td width='5%' height='25' rowspan='2'><div align='center' class='style3'>順序號碼</div></td>"
      Response.write "    <td width='10%' height='25' rowspan='2'><div align='center' class='style3'>掛號號碼</div></td>"
      Response.write "    <td height='25' colspan='2'><div align='center' class='style3'>收件人</div></td>"
      Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否回執(V)</div></td>"
      Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否航空(V)</div></td>"
      Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否印刷物(V)</div></td>"
      Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>重量</div></td>"
      Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>郵資</div></td>"
      Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>備考</div></td>"
      Response.write "  </tr>"
      Response.write "  <tr>"
      Response.write "    <td width='10%' height='25'><div align='center' class='style3'>姓名</div></td>"
      Response.write "    <td width='38%' height='25'><p align='center' class='style3'>寄達第名(或地址)</p>    </td>"
      Response.write "  </tr>"
      Row=1
      LastRow=0
      While Not RS1.EOF
        Response.write "<tr>"
        Response.write "  <td width='5%' height='30'><div align='center' class='style3'>"&Row&"</div></td>"
        Response.write "  <td width='10%' height='30'><span class='style3'></span></td>"
        Response.write "  <td width='10%' height='30'><span class='style3'></span>"&RS1("Donor_Name")&"</td>"
        If Session("Print_Desc")="1" Then
          Response.write "  <td width='38%' height='30'><span class='style3'></span>"&RS1("ZipCode")&RS1("Address")&"</td>"
        Else
          Response.write "  <td width='38%' height='30'><span class='style3'></span>"&RS1("Invoice_ZipCode")&RS1("Invoice_Address")&"</td>"
        End If
        Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
        Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
        Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
        Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
        Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
        Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
        Response.write "</tr>" 
        If (Row+20) Mod 20=0 Then
          Response.write "</table>"
          Response.write "<table width='720'  border='0' cellspacing='10'>"
          Response.write "  <tr>"
          Response.write "    <td width='58%' valign='top'>"
          Response.write "      <span class='style3'>"
          Response.write "        限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br />"
          Response.write "	      函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br />"
          Response.write "	      將本埠與外埠函件分別列單交寄。<br />"
          Response.write "	      此單由郵局免費供給，應由寄件人清晰填寫一式二份。<br />"
          Response.write "	      如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局經辦員逐件核對。<br />"
          Response.write "	      日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br />"
          Response.write "	      錢鈔或有價證券請利用報值或保價交寄。<br />"
          Response.write "	    </span>"
          Response.write "	  </td>"
          Response.write "    <td width='38%'>"
          Response.write "      <p class='style3'>上開　限時掛號 </p>"
          Response.write "      <p class='style3'>掛號函件 / 共　　20　　件 照收無誤　快捷郵件 </p>"
          Response.write "	    <p class='style3'>　</p>"
          Response.write "      <p class='style3'><font size='3'>郵資共計　　　　&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 元</font> </p>"
          Response.write "      <p align='center' class='style3'>經辦員簽署</p>"
          Response.write "    </td>"
          Response.write "  </tr>"
          Response.write "</table>"
          If Row<RS1.Recordcount Then
          	'換頁處理
          	Response.Write "<div class='pagebreak'>&nbsp;</div>"
          	Response.write "<table width='720'  border='0' cellspacing='0'>"
          	Response.write "  <tr>"
          	Response.write "    <td width='40%' height='120' valign='bottom' colspan='2'>"
          	Response.write "      <table width='100%' border='0' cellspacing='0'>"
          	Response.write "        <tr>"
          	Response.write "          <td height='25'><span class='style3'>中華民國&nbsp;"&Year(Date())-1911&"&nbsp;年&nbsp;"&Month(Date())&"&nbsp;月&nbsp;"&Day(Date())&"&nbsp;日</span></td>"
          	Response.write "        </tr>"
          	Response.write "        <tr>"
          	Response.write "          <td height='25'><span class='style3'>寄件人："&Session("user_name")&"</span></td>"
          	Response.write "        </tr>"
          	Response.write "        <tr>"
          	Response.write "          <td height='25'>名&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;稱："&Comp_Name&"</td>"
          	Response.write "        </tr>"
          	Response.write "      </table>"
          	Response.write "    </td>"
          	Response.write "    <td width='30%' height='120' align='center'>"
          	Response.write "      <p class='style1'><u>中&nbsp; 華&nbsp; 民&nbsp; 國&nbsp; 郵&nbsp; 政</u></p>"
          	Response.write "      <p class='style1'><u>限時掛號</u></p>"
          	Response.write "      <p class='style1'> 交寄大宗<u>掛&nbsp;&nbsp;&nbsp; 號</u>函件執據</p>"
          	Response.write "      <p class='style1'> <u>快捷郵件</u> </p>"
          	Response.write "    </td>"
          	Response.write "    <td width='30%' height='120' align='center'>"
          	Response.write "      <p><img src='../images/post.gif' width='82' height='80'></p>"
          	Response.write "      <p class='style3'>郵局郵戳</p>"
          	Response.write "    </td>"
          	Response.write "  </tr>"
          	Response.write "  <tr>"
          	Response.write "    <td width='21%' height='25'><span class='style3'>寄件人代表："&Dept_ContaCtor&"</span></td>"
          	Response.write "    <td width='48%' height='25' colspan='2'><span class='style3'>詳細地址："&Dept_Address&"</span></td>"
          	Response.write "    <td width='30%' height='25'><span class='style3'>電話號碼："&Dept_Tel&"</span></td>"
          	Response.write "  </tr>"
          	Response.write "</table>"
          	Response.write "<table width='720'  border='1' cellspacing='0' bordercolor='#000000' style='border-collapse: collapse'>"
          	Response.write "  <tr>"
          	Response.write "    <td width='5%' height='25' rowspan='2'><div align='center' class='style3'>順序號碼</div></td>"
          	Response.write "    <td width='10%' height='25' rowspan='2'><div align='center' class='style3'>掛號號碼</div></td>"
          	Response.write "    <td height='25' colspan='2'><div align='center' class='style3'>收件人</div></td>"
          	Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否回執(V)</div></td>"
          	Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否航空(V)</div></td>"
          	Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否印刷物(V)</div></td>"
          	Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>重量</div></td>"
          	Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>郵資</div></td>"
          	Response.write "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>備考</div></td>"
          	Response.write "  </tr>"
          	Response.write "  <tr>"
          	Response.write "    <td width='10%' height='25'><div align='center' class='style3'>姓名</div></td>"
          	Response.write "    <td width='38%' height='25'><p align='center' class='style3'>寄達第名(或地址)</p>    </td>"
          	Response.write "  </tr>"
          End If
          LastRow=0
        Else
          LastRow=LastRow+1
        End If
        Row=Row+1
        Response.Flush
        Response.Clear
        RS1.MoveNext
      Wend
      RS1.Close
      Set RS1=Nothing
      Row=Row-1
      If (Row+20) Mod 20<>0 Then
        While (Row+20) Mod 20<>0
          Response.write "<tr>"
          Response.write "  <td width='5%' height='30'><div align='center' class='style3'> </div></td>"
          Response.write "  <td width='10%' height='30'><span class='style3'></span></td>"
          Response.write "  <td width='10%' height='30'><span class='style3'></span></td>"
          Response.write "  <td width='38%' height='30'><span class='style3'></span></td>"
          Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
          Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
          Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
          Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
          Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
          Response.write "  <td width='6%' height='30'><span class='style3'></span></td>"
          Response.write "</tr>"
          Row=Row+1
        Wend
        Response.write "</table>"
        Response.write "<table width='720'  border='0' cellspacing='10'>"
        Response.write "  <tr>"
        Response.write "    <td width='58%' valign='top'>"
        Response.write "      <span class='style3'>"
        Response.write "        限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br />"
        Response.write "	      函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br />"
        Response.write "	      將本埠與外埠函件分別列單交寄。<br />"
        Response.write "	      此單由郵局免費供給，應由寄件人清晰填寫一式二份。<br />"
        Response.write "	      如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局經辦員逐件核對。<br />"
        Response.write "	      日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br />"
        Response.write "	      錢鈔或有價證券請利用報值或保價交寄。<br />"
        Response.write "	    </span>"
        Response.write "	  </td>"
        Response.write "    <td width='38%'>"
        Response.write "      <p class='style3'>上開　限時掛號 </p>"
        Response.write "      <p class='style3'>掛號函件 / 共　　"&LastRow&"　　件 照收無誤　快捷郵件 </p>"
        Response.write "	    <p class='style3'>　</p>"
        Response.write "      <p class='style3'><font size='3'>郵資共計　　　　&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 元</font> </p>"
        Response.write "      <p align='center' class='style3'>經辦員簽署</p>"
        Response.write "    </td>"
        Response.write "  </tr>"
        Response.write "</table>"
      End If
    End If
    Session.Contents.Remove("DeptId")
    Session.Contents.Remove("Donate_Where")
    Session.Contents.Remove("Print_Type")
    Session.Contents.Remove("Print_Desc")
    Session.Contents.Remove("SQL1")
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->