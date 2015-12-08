<!--#include file="../include/dbfunctionJ.asp"-->
<%
Invoice_Prog=""
Invoice_Pre=""
Rept_Licence=""
'20140122 Modify by GoodTV Tanya:隱藏「收據開立」：年度證明選項
Donate_Invoice="N"
Donate_AddPrint="Y"
Donate_PostPrint="Y"
SQL1="Select Invoice_Prog,Invoice_Pre,Rept_Licence,Donate_Invoice,Donate_AddPrint,Donate_PostPrint From DEPT Where Dept_Id='"&session("dept_id")&"'"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then
  If RS1("Invoice_Prog")<>"" Then Invoice_Prog=RS1("Invoice_Prog")
  If RS1("Invoice_Pre")<>"" Then Invoice_Pre=RS1("Invoice_Pre")
  If RS1("Rept_Licence")<>"" Then Rept_Licence=RS1("Rept_Licence")
'  If RS1("Donate_Invoice")<>"" Then Donate_Invoice=RS1("Donate_Invoice")
  If RS1("Donate_AddPrint")<>"" Then Donate_AddPrint=RS1("Donate_AddPrint")
  If RS1("Donate_PostPrint")<>"" Then Donate_PostPrint=RS1("Donate_PostPrint")
End If
RS1.Close
Set RS1=Nothing

DonateDesc="物資捐贈"
SQL1="Select DonateDesc=CodeDesc From CASECODE Where CodeType='Payment2' And CodeDesc Like '%物%' Order By Seq"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then DonateDesc=RS1("DonateDesc")
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
  
InvoiceType0=""
SQL1="Select InvoiceTypeO=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc <>'"&InvoiceTypeN&"' And CodeDesc <>'"&InvoiceTypeY&"' Order By Seq"
Call QuerySQL(SQL1,RS1)
While Not RS1.EOF
  If InvoiceTypeO="" Then 
    InvoiceTypeO=RS1("InvoiceTypeO")
  Else
    InvoiceTypeO=InvoiceTypeO&","&RS1("InvoiceTypeO")
  End If
  RS1.MoveNext
Wend
RS1.Close
Set RS1=Nothing
%>
<%Prog_Id="invoice_print"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="invoice_print_rpt.asp" target="main">
      <input type="hidden" name="action">
      <input type="hidden" name="DonateDesc" value="<%=DonateDesc%>">	
      <input type="hidden" name="InvoiceTypeO" value="<%=InvoiceTypeO%>">
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
                      <td class="td02-c">
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
                      <td class="td02-c" align="right">經手人：</td>
                      <td class="td02-c">
                        <%
                          If Session("user_group")<>"系統管理" And Session("user_group")<>"捐款管理" Then
                            SQL="Select Create_User=user_name From userfile Where user_name='"&Session("user_name")&"' Order By dept_id,user_group"
                            FName="Create_User"
                            Listfield="Create_User"
                            menusize="1"
                            BoundColumn=Session("user_name")
                            BoundNull="N"
                            call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,BoundNull)
                          Else
                            If Session("user_id")<>"npois" Then
                              SQL="Select Create_User=user_name From userfile Where user_id<>'npois' Order By dept_id,user_group"
                            Else
                              SQL="Select Create_User=user_name From userfile Order By dept_id,user_group"
                            End If
                            FName="Create_User"
                            Listfield="Create_User"
                            menusize="1"
                            call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                          End If
                        %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width="100">捐款日期：</td>
                      <td class="td02-c" width="230"><%call Calendar("Donate_Date_Begin","")%> ~ <%call Calendar("Donate_Date_End","")%></td>
                      <!--20131202 Modify by GoodTV Tanya:改為查詢捐款人編號-->
                      <td class="td02-c" align="right" width="90">捐款人編號：</td>
                      <td class="td02-c" width="370">
                      	<input type="text" name="Donor_Id" size="11" class="font9">&nbsp;&nbsp;&nbsp;
                      	<!--<input type="text" name="Donor_Name" size="10" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Report_OnClick();}'>&nbsp;&nbsp;&nbsp;-->
                        <input type="checkbox" name="RePrint" value="Y"><font color="#000080">收據重印</font>
                      	<%If Donate_AddPrint="Y" Then%><input type="checkbox" name="ReAddPrint" value="Y"><font color="#000080">地址名條重印</font><%End If%>                      	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">募款活動：</td>
                      <td class="td02-c">
                      <%
                        SQL="Select Act_Id,Act_ShortName From ACT Order By Act_id Desc"
                        FName="Act_Id"
                        Listfield="Act_ShortName"
                        menusize="1"
                        BoundColumn=""
                        call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                      %>
                      </td>
                      <td class="td02-c" align="right">收據編號：</td>
                      <td class="td02-c"><input type="text" name="Invoice_No_Begin" size="20" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Report_OnClick();}'> ~ <input type="text" name="Invoice_No_End" size="20" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Report_OnClick();}'><%If Invoice_Pre<>"" Then%>&nbsp;</br><font color="#ff0000">(&nbsp;機構前置碼不需輸入&nbsp;)</font><%End If%></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width="100">收據開立：</td>
                      <td class="td02-c" width="680" colspan="3">
                      <%
                        If Donate_Invoice="Y" Then
                          SQL1="Select Invoice_Type=CodeDesc From CASECODE Where CodeType='InvoiceType' Order By Seq"
                        Else
                          SQL1="Select Invoice_Type=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc<>'"&InvoiceTypeY&"' Order By Seq"
                        End If
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        Row=0
                        While Not RS1.EOF
                          Row=Row+1
                          If Instr(InvoiceTypeO,RS1("Invoice_Type"))>0 Then
                            Response.Write "<Input Type='checkbox' Name='Invoice_Type' ID='Invoice_Type_"&Row&"' value='"&RS1("Invoice_Type")&"' checked >"&RS1("Invoice_Type")&"&nbsp;" 
                          Else
                            Response.Write "<Input Type='checkbox' Name='Invoice_Type' ID='Invoice_Type_"&Row&"' value='"&RS1("Invoice_Type")&"'>"&RS1("Invoice_Type")&"&nbsp;"
                          End If
                          RS1.MoveNext
                        Wend
                        RS1.Close
                        Set RS1=Nothing
                      %>                      
                      </td>                                                                 
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width="100">捐款方式：</td>
                      <td class="td02-c" width="680" colspan="3">
                      <%
                        SQL="Select Donate_Payment=CodeDesc From CASECODE Where CodeType='Payment' Order By Seq"
                        FName="Donate_Payment"
                        Listfield="Donate_Payment"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款用途：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL="Select Donate_Purpose=CodeDesc From CASECODE Where CodeType='Purpose' And CodeKind In ('D') Order By Seq"
                        FName="Donate_Purpose"
                        Listfield="Donate_Purpose"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">其他：</td>
                      <td class="td02-c" colspan="3">
                        <input type="checkbox" name="Issue_TypeM" value="M">手開收據
                        <input type="checkbox" name="Issue_TypeD" value="D">作廢收據
                      </td>
                    </tr>
                    <input type="hidden" name="Print_Pint" value="1">
                    <%If Invoice_Prog="5" Then%>
                    <input type="hidden" name="Invoice_Prog" value="<%=Invoice_Prog%>">	
                    <%Else%>
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">收據樣式：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL1="Select Seq,Invoice_Prog=CodeDesc From CASECODE Where CodeType='Invoice' Order By Seq"
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        Response.Write "<SELECT Name='Invoice_Prog' size='1' style='font-size: 9pt; font-family: 新細明體'>" 
                        Response.Write "<OPTION>" & " " & "</OPTION>"
                        While Not RS1.EOF
                          If Cstr(RS1("Seq"))=Cstr(Invoice_Prog) Then
                            Response.Write "<OPTION selected value='"&RS1("Seq")&"'>"&RS1("Invoice_Prog")&"</OPTION>"
                          Else
                            Response.Write "<OPTION value='"&RS1("Seq")&"'>"&RS1("Invoice_Prog")&"</OPTION>"
                          End If
                          RS1.MoveNext
                        Wend
                        Response.Write "</SELECT>"
                        RS1.Close
                        Set RS1=Nothing
                      %>
                      </td>
                    </tr>
                    <%End If%>
                    <input type="hidden" name="Licence">
					<tr>
                      <td class="td02-c" align="right">地址：</td>
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="Print_Address" id="Print_Address1" value="1" checked>國內&nbsp;
                        <input type="radio" name="Print_Address" id="Print_Address2" value="2">海外	
                      </td>
                    </tr>
                    <%If Donate_AddPrint="Y" Then%>
                    <tr>
                      <td class="td02-c" align="right">名條內容：</td>
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="Print_Desc" id="Print_Desc1" value="1" checked>收據地址&nbsp;
                        <input type="radio" name="Print_Desc" id="Print_Desc2" value="2">通訊地址	
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
                    <%Else%>
                    <input type="hidden" name="Print_Desc" value="1">
                    <input type="hidden" name="Label" value="1">	
                    <%End If%>
                    <tr>
                      <td class="td02-c" align="right">排序方式：</td>
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="OrderBy_Type" id="OrderBy_Type3" value="1" checked >收據編號&nbsp;	
                        <input type="radio" name="OrderBy_Type" id="OrderBy_Type2" value="2">捐款人編號&nbsp;
                        <input type="radio" name="OrderBy_Type" id="OrderBy_Type1" value="3">郵遞區號&nbsp;
                        <!--20140122 Add by GoodTV Tanya:開放身份別「主知名」可選擇是否列印-->
                      	<Input Type='checkbox' Name='Print_NoName' ID='Print_NoName' value='Y' checked><font color="#000080">不含身份別主知名</font>
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
                        If Donate_AddPrint="Y" Then Response.Write "<button id='address' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Address_OnClick()'><img src='../images/print.gif' width='20' height='20'><br>名條</button>&nbsp;"
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
function Window_OnLoad(){
  document.form.Donate_Date_Begin.focus();
}	
function Report_OnClick(){
  if(document.form.Donate_Date_Begin.value!=''){
    <%call CheckDateJ("Donate_Date_Begin","捐款日期起")%>
  }
  if(document.form.Donate_Date_End.value!=''){
    <%call CheckDateJ("Donate_Date_End","捐款日期迄")%>
  }
  <%call CheckStringJ("Invoice_Prog","收據樣式")%>
  if(confirm('您是否確定要將查詢結果列印？')){
    document.form.target='report';
    document.form.action.value='report';
    window.open('','report','toolbar=no,location=no,status=yes,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600');
    document.form.submit();
  }
}
function Address_OnClick(){
  if(document.form.Donate_Date_Begin.value!=''){
    <%call CheckDateJ("Donate_Date_Begin","捐款日期起")%>
  }
  if(document.form.Donate_Date_End.value!=''){
    <%call CheckDateJ("Donate_Date_End","捐款日期迄")%>
  }
  <%call CheckStringJ("Label","名條格式")%>
  if(confirm('您是否確定要將查詢結果列印？')){
    document.form.target='report';
    document.form.action.value='address';
    window.open('','report','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600');
    document.form.submit();
  }
}
function Post_OnClick(){
  if(document.form.Donate_Date_Begin.value!=''){
    <%call CheckDateJ("Donate_Date_Begin","捐款日期起")%>
  }
  if(document.form.Donate_Date_End.value!=''){
    <%call CheckDateJ("Donate_Date_End","捐款日期迄")%>
  }
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
  window.open('','report','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=0,left=0,width=500,height=300');
  document.form.submit();
}
--></script>