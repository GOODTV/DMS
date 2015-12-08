<!--#include file="../include/dbfunctionJ.asp"-->
<%
  Invoice_Pre=""
  Donate_PostPrint="Y"
  SQL1="Select Invoice_Pre,Donate_PostPrint From DEPT Where Dept_Id='"&Session("dept_id")&"'"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then
    Invoice_Pre=RS1("Invoice_Pre")
    Donate_PostPrint=RS1("Donate_PostPrint")
  End If
  RS1.Close
  Set RS1=Nothing
  
  InvoiceTypeN="不寄收據"
  SQL1="Select InvoiceTypeN=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Like '%不%' Order By Seq"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then InvoiceTypeN=RS1("InvoiceTypeN")
  RS1.Close
  Set RS1=Nothing
  
  InvoiceTypeY="年度收據"
  SQL1="Select InvoiceTypeY=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Like '%年%' Order By Seq"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then InvoiceTypeY=RS1("InvoiceTypeY")
  RS1.Close
  Set RS1=Nothing
  
  InvoiceTypeO=""
  SQL1="Select InvoiceTypeO=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc <>'"&InvoiceTypeN&"' And CodeDesc <>'"&InvoiceTypeY&"' Order By Seq"
  Call QuerySQL(SQL1,RS1)
  While Not RS1.EOF
    If InvoiceTypeO="" Then 
      InvoiceTypeO=RS1("InvoiceTypeO")
    Else
      InvoiceTypeO=InvoiceTypeO&"、"&RS1("InvoiceTypeO")
    End If
    RS1.MoveNext
  Wend
  RS1.Close
  Set RS1=Nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>文宣品名條列印</title>
  <link rel="stylesheet" type="text/css" href="../include/dms.css">
</head>
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="address_list_rpt.asp" target="report">
      <input type="hidden" name="action">
      <input type="hidden" name="InvoiceTypeN" value="<%=InvoiceTypeN%>">
      <input type="hidden" name="InvoiceTypeY" value="<%=InvoiceTypeY%>">
      <table width="700" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" style="background-color:#EEEEE3">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">文宣品名條列印</td>
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
                      <td class="td02-c" align="right">機構：</td>
                      <td class="td02-c" colspan="3">
                        <%
                          If Session("comp_label")="1" Then
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' And Dept_Id In ("&Session("all_dept_type")&") Order By Comp_Label,Dept_Id"
                          Else
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id In ("&Session("all_dept_type")&") Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
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
                      <td class="td02-c" align="right">類別：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL="Select Category=CodeDesc From CASECODE Where CodeType='Category' Order By Seq"
                        FName="Category"
                        Listfield="Category"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">身份別：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL="Select Donor_Type=CodeDesc From CASECODE Where CodeType='DonorType' Order By Seq"
                        FName="Donor_Type"
                        Listfield="Donor_Type"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width="100">通訊地址：</td>
                      <td class="td02-c" width="230"><%call CodeArea ("form","City","","Area","","Y")%></td>
                      <td class="td02-c" align="right" width="90">收據地址：</td>
                      <td class="td02-c" width="340"><%call CodeArea ("form","Invoice_City","","Invoice_Area","","N")%></td>
                    </tr>
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">名條用途：</td>
                      <td class="td02-c" colspan="3">
                      	<input type="radio" name="Print_Type" id="Print_Type2" value="2" checked >郵寄會訊
                      	<input type="radio" name="Print_Type" id="Print_Type3" value="3" >郵寄年報
                      	<input type="radio" name="Print_Type" id="Print_Type4" value="4" >郵寄生日卡
                      	<input type="radio" name="Print_Type" id="Print_Type5" value="5" >郵寄賀卡
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">名條內容：</td>
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="Print_Desc" id="Print_Desc2" value="2" checked >通訊地址&nbsp;	
                        <input type="radio" name="Print_Desc" id="Print_Desc1" value="1">收據地址
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">名條格式：</td>
                      <td class="td02-c" colspan="3">
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
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="OrderBy_Type" id="OrderBy_Type1" value="1" checked >郵遞區號&nbsp;
                        <input type="radio" name="OrderBy_Type" id="OrderBy_Type2" value="2">捐款人編號&nbsp;
                        <!--<input type="radio" name="OrderBy_Type" id="OrderBy_Type3" value="3">收據編號&nbsp;<font color="#ff0000">(&nbsp;限郵寄收據名條&nbsp;)</font>-->
                      </td>
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
                        Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick()'><img src='../images/print.gif' width='19' height='20'><br>名條</button>&nbsp;"
                        If Donate_PostPrint="Y" Then Response.Write "<button id='post' style='position:relative;left:30;width:60;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Post_OnClick()'><img src='../images/print.gif' width='20' height='20'><br>大宗掛號</button>&nbsp;"
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
  <%call CheckStringJ("Label","標籤格式")%>
  if(confirm('您是否確定要將查詢結果列印？')){
    document.form.target='report';
    document.form.action.value='report';
    window.open('','report','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600');
    document.form.submit();
  }
}
function Post_OnClick(){
  if(confirm('您是否確定要將查詢結果列印？')){
    document.form.target='report';
    document.form.action.value='post';
    window.open('','report','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600');
    document.form.submit();
  }
}
function Help_OnClick(){
  document.form.target='report';
  document.form.action.value='help';
  window.open('','report','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=300');
  document.form.submit();
}
--></script>