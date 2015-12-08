<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL_IS="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"' "
Set RS_IS=Server.CreateObject("ADODB.Recordset")
RS_IS.Open SQL_IS,Conn,1,1
IsDonation=RS_IS("IsDonation")
IsContribute=RS_IS("IsContribute")
IsMember=RS_IS("IsMember")
IsMemberNo=RS_IS("IsMemberNo")
IsMemberPre=RS_IS("IsMemberPre")
IsShopping=RS_IS("IsShopping")
RS_IS.Close
Set RS_IS=Nothing

SQL="Select *,Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+Address Else DONOR.ZipCode+A.mValue+Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then DONOR.Invoice_Address Else Case When C.mValue<>B.mValue Then DONOR.Invoice_ZipCode+C.mValue+D.mValue+DONOR.Invoice_Address Else DONOR.Invoice_ZipCode+C.mValue+DONOR.Invoice_Address End End) From DONOR " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _ 
    "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
    "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
    "Where Donor_Id='"&request("donor_id")&"'"
Call QuerySQL(SQL,RS)
%>
<%Prog_Id="donate"%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link rel="stylesheet" type="text/css" href="../include/dms.css">
  <title><%=Prog_Desc%></title>
</head>
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3" size="20">
        <tr>
          <td>
            <table width="780" border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"></td>
                <td width="95%">
  		            <table width="40%"  border="0" cellspacing="0" cellpadding="0">
		                <tr>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                      <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                    </tr>
                    <tr>
                      <td class="table62-bg">&nbsp;</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;"><%=Prog_Desc%></td>
                      <td class="table63-bg">&nbsp;</td>
                    </tr>  
    	            </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td> 
	          <table width="880" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">&nbsp;</td>
                <td valign="top">
                  <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                    <tr>
                      <td width="99%">
                        <table border="1" cellpadding="2" style="border-collapse: collapse" width="100%" height="25" cellspacing="1">
                          <tr>
                            <td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="donor_edit.asp?donor_id=<%=request("donor_id")%>">捐款人資料</a></td>
                            <td class="button2-bg"><img border="0" src="../images/red_arrow.gif" align="texttop">捐款紀錄</td>
                            <td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="pledge_data.asp?donor_id=<%=request("donor_id")%>">轉帳授權書記錄</a></td>
                            <%If IsContribute="Y" Then%><td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="../contribute/contribute_data.asp?donor_id=<%=request("donor_id")%>">公關贈品記錄</a></td><%End If%>
                            <%If IsMember="Y" Then%><td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="../member/signup_data.asp?donor_id=<%=request("donor_id")%>">活動報名記錄</a></td><%End If%>
                            <%If IsShopping="Y" Then%><td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="../shopping/shopping_data.asp?donor_id=<%=request("donor_id")%>">商品義賣記錄</a></td><%End If%>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                           <tr>
                            <td align="right">捐款人：</td>
                            <td align="left">
                              <input type="text" name="Donor_Name" size="15" class="font9t" readonly value="<%=Data_Minus(RS("Donor_Name"))%>&nbsp;<%=RS("Title")%>">
                               &nbsp;捐款人編號：
                              <input type="text" name="DonorId" size="7" class="font9t" readonly value="<%=request("donor_id")%>">
							  &nbsp;收據開立：
                              <input type="text" name="InvoiceType" size="8" class="font9t" readonly value="<%=RS("Invoice_Type")%>">
                              &nbsp;身分證/統編：
                              <input type="text" name="IDNo" size="10" class="font9t" readonly value="<%=RS("IDNo")%>">
                              &nbsp;收據地址：
                              <input type="text" name="Invoice_Address2" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Invoice_Address2"))%>">	
                            </td>
                          </tr>
                          <tr>
                            <td class="td02-c" colspan="2" width="100%"><iframe name="detail" src="donate_data_list.asp?donor_id=<%=request("donor_id")%>" height="420" width="100%" frameborder="0" scrolling="auto"></iframe></td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
                <td class="table63-bg">&nbsp;</td>
              </tr>
              <tr>
                <td style="background-color:#EEEEE3"><img src="../images/table06_06.gif" width="10" height="10"></td>
                <td class="table64-bg"><img src="../images/table06_07.gif" width="1" height="10"></td>
                <td style="background-color:#EEEEE3"><img src="../images/table06_08.gif" width="10" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>   
    </form>
  </center></div></p>
  <%Message()%>
</body>
</html>