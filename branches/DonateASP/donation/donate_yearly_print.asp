<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL1="Select Invoice_Prog2,Rept_Licence From DEPT Where Dept_Id='"&session("dept_id")&"'"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then 
  Invoice_Prog2=RS1("Invoice_Prog2")
  Rept_Licence=RS1("Rept_Licence")
End If
RS1.Close
Set RS1=Nothing

DonateDesc="物"
SQL1="Select DonateDesc=CodeDesc From CASECODE Where CodeType='Payment' And CodeDesc Like '%物%' Order By Seq"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then DonateDesc=RS1("DonateDesc")
RS1.Close
Set RS1=Nothing

InvoiceTypeY="年度收據"
SQL1="Select InvoiceTypeY=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Like '%年%' Order By Seq"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then InvoiceTypeY=RS1("InvoiceTypeY")
RS1.Close
Set RS1=Nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>年度捐款證明名單</title>
  <link rel="stylesheet" type="text/css" href="../include/dms.css">
</head>
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="invoice_yearly_print_rpt.asp" target="main">
      <input type="hidden" name="action">
      <input type="hidden" name="Invoice_Prog2" value="<%=Invoice_Prog2%>">
      <input type="hidden" name="Rept_Licence" value="<%=Rept_Licence%>">
      <input type="hidden" name="DonateDesc" value="<%=DonateDesc%>">
      <input type="hidden" name="InvoiceTypeY" value="<%=InvoiceTypeY%>">
      <table width="700" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="100%"  border="0" cellspacing="0" cellpadding="0" align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"></td>
                <td width="95%">
  		            <table width="60%"  border="0" cellspacing="0" cellpadding="0">
		                <tr>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                      <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                    </tr>
                    <tr>
                      <td class="table62-bg">&nbsp;</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">年度捐款證明名單</td>
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
	          <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">&nbsp;</td>
                <td valign="top">
                  <table width="100%" border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111">
                    <tr>
                      <td class="td02-c" align="right" width="120">機構：</td>
                      <td class="td02-c" width="660">
                        <%
                          If Session("comp_label")="1" Then
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' Order By Comp_Label,Dept_Id"
                          Else
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id='"&Session("dept_id")&"' Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
                          End If
                          FName="Dept_Id"
                          Listfield="Comp_ShortName"
                          menusize="1"
                          BoundColumn=Session("dept_id")
                          call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                        %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款年度：</td>
                      <td class="td02-c">
                      <%
                        SQL1="Select Begin_Year=Isnull(Min(Year(Donate_Date)),"&Year(Date())&") From DONATE"
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        If Not RS1.EOF Then Begin_Year=RS1("Begin_Year")
                        RS1.Close
                        Set RS1=Nothing
                        Response.Write "<SELECT Name='Donate_Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        For I= Cint(Year(Date())) To Cint(Begin_Year) Step -1
                          If I=Cint(Year(Date())-1) Then
                            Response.Write "<OPTION selected value='"&I&"'>"&I&"</OPTION>"
                          Else
                            Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
                          End If
                        Next
                        Response.Write "</SELECT>"
                      %>	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款人：</td>
                      <td class="td02-c">
                      	<input type="text" name="Donor_Name" size="24" class="font9">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款人編號：</td>
                      <td class="td02-c"><input type="text" name="Donor_Id_Begin" size="10" class="font9"> ~ <input type="text" name="Donor_Id_End" size="10" class="font9"></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款人姓名改印：</td>
                      <td class="td02-c"><input type="text" name="Invoice_Title" size="24" class="font9">&nbsp;&nbsp;<font color="#FF0000">(&nbsp;限單一捐款人，才可改印捐款人姓名。&nbsp;)</font></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right"><font color="#000080">勸募許可文號：</font></td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL1="Select Licence=CodeDesc From CASECODE Where CodeType='Licence' Order By Seq"
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        Response.Write "<SELECT Name='Rept_Licence' size='1' style='font-size: 9pt; font-family: 新細明體'>" 
                        Response.Write "<OPTION>" & " " & "</OPTION>"
                        While Not RS1.EOF
                          If Cstr(RS1("Licence"))=Cstr(Rept_Licence) Then
                            Response.Write "<OPTION selected value='"&RS1("Licence")&"'>"&RS1("Licence")&"</OPTION>"
                          Else
                            Response.Write "<OPTION value='"&RS1("Licence")&"'>"&RS1("Licence")&"</OPTION>"
                          End If
                          RS1.MoveNext
                        Wend
                        Response.Write "</SELECT>"
                        RS1.Close
                        Set RS1=Nothing
                      %>&nbsp;<font color="#ff0000">(&nbsp;證明名單內容如須加印請選擇核准文號&nbsp;)</font>
                      </td>
                    </tr>
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">名條內容：</td>
                      <td class="td02-c">
                        <input type="radio" name="Print_Desc" id="Print_Desc1" value="1" checked >收據地址&nbsp;
                        <input type="radio" name="Print_Desc" id="Print_Desc2" value="2">通訊地址	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">名條格式：</td>
                      <td class="td02-c">
                      <%
                        SQL1="Select Label From DEPT Where Dept_Id='"&Session("dept_id")&"' "
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        If Not RS1.EOF Then Label=RS1("Label")
                        RS1.Close
                        Set RS1=Nothing
                        
                        SQL1="Select Seq,Label=CodeDesc From CASECODE Where CodeType='Label' Order By Seq"
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        Response.Write "<SELECT Name='Label' size='1' style='font-size: 9pt; font-family: 新細明體'>" 
                        Response.Write "<OPTION>" & " " & "</OPTION>"
                        While Not RS1.EOF
                          If Cstr(RS1("Seq"))=Cstr(Label) Then
                            Response.Write "<OPTION selected value='"&RS1("Seq")&"'>"&RS1("Label")&"</OPTION>"
                          Else
                            Response.Write "<OPTION value='"&RS1("Seq")&"'>"&RS1("Label")&"</OPTION>"
                          End If
                          RS1.MoveNext
                        Wend
                        Response.Write "</SELECT>"
                        RS1.Close
                        Set RS1=Nothing
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">排序方式：</td>
                      <td class="td02-c">
                        <input type="radio" name="OrderBy_Type" id="OrderBy_Type1" value="1" checked >郵遞區號&nbsp;
                        <input type="radio" name="OrderBy_Type" id="OrderBy_Type2" value="2">捐款人編號&nbsp;
                        <input type="radio" name="OrderBy_Type" id="OrderBy_Type3" value="3">收據編號&nbsp;	
                      </td>
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
                        Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick()'><img src='../images/print.gif' width='19' height='20'><br>收據</button>&nbsp;"
                        Response.Write "<button id='address' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Address_OnClick()'><img src='../images/print.gif' width='20' height='20'><br>名條</button>&nbsp;"	
                        Response.Write "<button id='help' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Help_OnClick()'><img src='../images/help.gif' width='19' height='20'><br>說明</button>"
    			            %>
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
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Report_OnClick(){

}
function Address_OnClick(){

}
function Help_OnClick(){

}
--></script>