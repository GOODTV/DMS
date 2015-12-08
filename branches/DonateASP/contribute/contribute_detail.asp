<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function ContributeDataGrid (SQL)
  Goods_Amt=0
  Row=0
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id=grid border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
  Response.Write "<tr>"
  Response.Write "<td align=""center"" bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>序號</span></font></td>"
  For I = 1 To FieldsCount
	  Response.Write "<td align=""center"" bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr>"
	  For I = 1 To FieldsCount
	    If I=1 Then
        Row=Row+1
        Response.Write "<td align=""center""><span style='font-size: 9pt; font-family: 新細明體'>" & Row & "</span></td>"
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I) & "</span></td>"
      ElseIf I=2 Then
        Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(I),0) & "</span></td>"
      ElseIf I=4 Then
        Goods_Amt=Goods_Amt+CDbl(RS1(I))
        Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(I),0) & "</span></td>"
      ElseIf I=5 Or I=7 Then
        Response.Write "<td align=""center""><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I) & "</span></td>"
      Else
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I) & "</span></td>"
      End If
    Next
    RS1.MoveNext
    Response.Write "</tr>"
  Wend
  Response.Write "<tr><td colspan=""4"" align=""right""><b>折合現金合計</b></td><td align=""right""><b>"&FormatNumber(Goods_Amt,0)&"</b></td><td colspan=""3""></td></tr>"
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function

If request("action")="void" Then
  Check_Void=True
  Goods_Name=""
  SQL1="Select CONTRIBUTEDATA.*,Storage_Qty=Isnull(GOODS.Goods_Qty,0),Goods_IsStock=GOODS.Goods_IsStock From CONTRIBUTEDATA Left Join GOODS On CONTRIBUTEDATA.Goods_Id=GOODS.Goods_Id Where CONTRIBUTEDATA.Contribute_Id='"&request("contribute_id")&"'"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then
    While Not RS1.EOF
      If RS1("Contribute_IsStock")="Y" And RS1("Goods_IsStock")="Y" And CDbl(RS1("Goods_Qty"))>CDbl(RS1("Storage_Qty")) Then
        Check_Void=False
        If Goods_Name="" Then
          Goods_Name=RS1("Goods_Name")
        Else
          Goods_Name=Goods_Name&"、"&RS1("Goods_Name")
        End If
      End If
      RS1.MoveNext
    Wend
    If Check_Void=False Then
      session("errnumber")=1
      session("msg")="『"&Goods_Name&"』 \n\n庫存量不足無法收據無法作廢 ！"
    End If
  End If
  RS1.Close
  Set RS1=Nothing

  If Check_Void Then
    SQL1="Select * From CONTRIBUTEDATA Where Contribute_Id='"&request("Contribute_Id")&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,3
    If Not RS1.EOF Then
      While Not RS1.EOF
        Goods_Qty=CDbl(RS1("Goods_Qty"))
        Goods_Amt=CDbl(RS1("Goods_Amt"))
        If RS1("Contribute_IsStock")="Y" And RS1("Goods_Id")<>"" Then
          SQL2="Select Goods_Qty From GOODS Where Goods_Id='"&RS1("Goods_Id")&"' And Goods_IsStock='Y'"
          Set RS2 = Server.CreateObject("ADODB.RecordSet")
          RS2.Open SQL2,Conn,1,3
          If Not RS2.EOF Then
            RS2("Goods_Qty")=CDbl(RS2("Goods_Qty"))-CDbl(RS1("Goods_Qty"))
            RS2.Update
          End If
          RS2.Close
          Set RS2=Nothing
        End If
        RS1("Goods_Qty")="0"
        RS1("Goods_Amt")="0"
        RS1("Goods_Qty_D")=Goods_Qty
        RS1("Goods_Amt_D")=Goods_Amt
        RS1.MoveNext
      Wend
    End If
    RS1.Close
    Set RS1=Nothing
    
    SQL1="Select * From CONTRIBUTE Where contribute_id='"&request("contribute_id")&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,3
    Donor_Id=RS1("Donor_Id")
    Contribute_Amt=RS1("Contribute_Amt")
    RS1("Contribute_Amt")="0"
    RS1("Contribute_Amt_D")=Contribute_Amt
    RS1("Issue_Type")="D"
    RS1.Update
    RS1.Close
    Set RS1=Nothing
  
    '修改捐款人捐款紀錄
    call Declare_DonorId (request("donor_id"))

    session("errnumber")=1
    session("msg")="收據已作廢 ！"
  End If
End If

If request("action")="return" Then
  SQL1="Select * From CONTRIBUTEDATA Where Contribute_Id='"&request("Contribute_Id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  If Not RS1.EOF Then
    While Not RS1.EOF
      Goods_Qty_D=CDbl(RS1("Goods_Qty_D"))
      Goods_Amt_D=CDbl(RS1("Goods_Amt_D"))
      If RS1("Contribute_IsStock")="Y" And RS1("Goods_Id")<>"" Then
        SQL2="Select Goods_Qty From GOODS Where Goods_Id='"&RS1("Goods_Id")&"' And Goods_IsStock='Y'"
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,3
        If Not RS2.EOF Then
          RS2("Goods_Qty")=CDbl(RS2("Goods_Qty"))+CDbl(RS1("Goods_Qty_D"))
          RS2.Update
        End If
        RS2.Close
        Set RS2=Nothing
      End If
      RS1("Goods_Qty")=Goods_Qty_D
      RS1("Goods_Amt")=Goods_Amt_D
      RS1("Goods_Qty_D")="0"
      RS1("Goods_Amt_D")="0"
      RS1.MoveNext
    Wend
  End If
  RS1.Close
  Set RS1=Nothing
    
  SQL1="Select * From CONTRIBUTE Where contribute_id='"&request("contribute_id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  Donor_Id=RS1("Donor_Id")
  Contribute_Amt_D=RS1("Contribute_Amt_D")
  RS1("Contribute_Amt")=RS1("Contribute_Amt_D")
  RS1("Contribute_Amt_D")="0"
  RS1("Issue_Type")=RS1("Issue_Type_Keep")
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  
  '修改捐款人捐款紀錄
  call Declare_DonorId (request("donor_id"))
  
  session("errnumber")=1
  session("msg")="收據已還原 ！"
End If

DonateDesc="物"
SQL1="Select DonateDesc=CodeDesc From CASECODE Where CodeType='Payment2' And CodeDesc Like '%物%' Order By Seq"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then DonateDesc=RS1("DonateDesc")
RS1.Close
Set RS1=Nothing

InvoiceTypeY="年度匯整"
SQL1="Select InvoiceTypeY=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Like '%年%' Order By Seq"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then InvoiceTypeY=RS1("InvoiceTypeY")
RS1.Close
Set RS1=Nothing

SQL="Select CONTRIBUTE.*,Donor_Name=DONOR.Donor_Name,Title=DONOR.Title,Category=DONOR.Category,Donor_Type=DONOR.Donor_Type,IDNo=DONOR.IDNo, " & _
    "Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+DONOR.Address Else DONOR.ZipCode+A.mValue+DONOR.Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then DONOR.Address Else Case When C.mValue<>D.mValue Then DONOR.Invoice_ZipCode+C.mValue+D.mValue+DONOR.Invoice_Address Else DONOR.Invoice_ZipCode+C.mValue+DONOR.Address End End), " & _
    "DEPT.Comp_ShortName,ACT.Act_ShortName,DEPT.Donate_Invoice,DEPT.Invoice_Prog3,DEPT.Rept_Licence " &_ 
    "From CONTRIBUTE " & _
    "Join DONOR On CONTRIBUTE.Donor_Id=DONOR.Donor_Id " & _
    "Join Dept On CONTRIBUTE.Dept_Id=DEPT.Dept_Id " & _
    "Left Join ACT On CONTRIBUTE.Act_Id=ACT.Act_Id " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _
    "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
    "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
    "Where Contribute_Id='"&request("contribute_id")&"' "
Call QuerySQL(SQL,RS)

Check_Close=Get_Close("2",Cstr(RS("Dept_Id")),Cstr(RS("Contribute_Date")),Cstr(Session("user_id")))
%>
<%Prog_Id="contribute"%>
<!--#include file="../include/head.asp"-->
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
    	<input type="hidden" name="action" value="">
    	<input type="hidden" name="ctype" value="<%=request("ctype")%>">
    	<input type="hidden" name="Donor_Id" value="<%=RS("Donor_Id")%>">
      <input type="hidden" name="contribute_id" value="<%=RS("Contribute_Id")%>">
      <input type="hidden" name="donate_invoice" value="<%=RS("Donate_Invoice")%>">
      <input type="hidden" name="Invoice_Prog" value="<%=RS("Invoice_Prog3")%>">
      <input type="hidden" name="Rept_Licence" value="<%=RS("Rept_Licence")%>">
      <input type="hidden" name="DonateDesc" value="<%=DonateDesc%>">		
      <%If Cstr(RS("Contribute_Payment"))=Cstr(DonateDesc) Then%>	
      <input type="hidden" name="Print_Pint" value="2">
      <%Else%>
      <input type="hidden" name="Print_Pint" value="1">
      <%End If%>
      <%If RS("Invoice_Print")="1" Then%>	
      <input type="hidden" name="RePrint" value="Y">
      <%Else%>
      <input type="hidden" name="RePrint" value="">
      <%End If%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="780"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"> </td>
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
	          <table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
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
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
                          <tr>
                            <td align="right">捐贈人：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Donor_Name" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Donor_Name"))%>&nbsp;<%=RS("Title")%>">
                              &nbsp;類&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;別：
                              <input type="text" name="Category" size="10" class="font9t" readonly value="<%=RS("Category")%>">
                              &nbsp;身份別：
                              <input type="text" name="Donor_Type" size="42" class="font9t" readonly value="<%=RS("Donor_Type")%>">
                            </td>
                          </tr>
                          <tr>
                            <td align="right">收據地址：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Invoice_Address2" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Invoice_Address2"))%>">
                              &nbsp;身分證/統編：
                              <input type="text" name="IDNo" size="10" class="font9t" readonly value="<%=RS("IDNo")%>">
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td width="10%" align="right">捐贈日期：</td>
                            <td width="16% align="left"><input type="text" name="Contribute_Date" size="12" class="font9t" readonly value="<%=RS("Contribute_Date")%>"></td>
                            <td width="10%" align="right">捐贈方式：</td>
                            <td width="16%"><input type="text" name="Contribute_Payment" size="12" class="font9t" readonly value="<%=RS("Contribute_Payment")%>"></td>
                            <td width="10%" align="right">捐贈用途：</td>
                            <td width="14%"><input type="text" name="Contribute_Purpose" size="12" class="font9t" readonly value="<%=RS("Contribute_Purpose")%>"></td>
                            <td width="10%" align="right">收據開立：</td>
                            <td width="14%"><input type="text" name="Invoice_Type" size="12" class="font9t" readonly value="<%=RS("Invoice_Type")%>"></td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">機構名稱：</td>
                            <td><input type="text" name="Comp_ShortName" size="12" class="font9t" readonly value="<%=RS("Comp_ShortName")%>"></td>
                            <td align="right">收據抬頭：</td>
                            <td colspan="2"><input type="text" name="Invoice_Title" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Invoice_Title"))%>"></td>
                            <td><input type="checkbox" name="Issue_Type" value="M" <%If RS("Invoice_No")="M" Then Response.Write "checked" End If%> disabled >手開收據</td>
                            <td align="right">收據編號：</td>
                            <td><input type="text" name="Invoice_No" size="12" class="font9t" readonly value="<%=RS("Invoice_Pre")&RS("Invoice_No")%>"></td>
                          </tr> 
                          <tr>
                            <td align="right">沖帳日期：</td>
                            <td align="left"><input type="text" name="Accoun_Date" size="12" class="font9t" readonly value="<%=RS("Accoun_Date")%>"></td>
                            <td align="right">會計科目：</td>
                            <td><input type="text" name="Accounting_Title" size="12" class="font9t" readonly value="<%=RS("Accounting_Title")%>"></td>
                            <td align="right">募款活動：</td>
                            <td colspan="3"><input type="text" name="Act_ShortName" size="42" class="font9t" readonly value="<%=RS("Act_ShortName")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">捐贈備註：</td>
                            <td align="left" colspan="3"><textarea rows="3" name="Comment" cols="42" class="font9t" readonly ><%=RS("Comment")%></textarea></td>
                            <td align="right">收據備註：<br />(列印收據用)</td>
                            <td align="left" colspan="3"><textarea rows="3" name="Invoice_PrintComment" cols="40" class="font9t" readonly ><%=RS("Invoice_PrintComment")%></textarea></td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">捐贈內容：</td>
                            <td align="left" colspan="7">
                            <%
                              SQL="Select ser_no,物品名稱=Goods_Name,數量=Goods_Qty,單位=Goods_Unit,折合現金=Goods_Amt,保存期限=CONVERT(VarChar,Goods_DueDate,111),備註=Goods_Comment,寫入庫存=(Case When Contribute_IsStock='Y' Then 'V' Else '' End) " & _
                                  "From CONTRIBUTEDATA Where Contribute_Id='"&RS("contribute_id")&"' Order By Ser_No "
                              HLink="contribute_data_edit.asp?contribute_id="&RS("contribute_id")&"&ser_no="
                              LinkParam="ser_no"
                              LinkTarget="main"
                              LinkType="window"
                              Call ContributeDataGrid (SQL)
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="center" colspan="8">
                              <%
                                If RS("Issue_Type")<>"D" Then
                                  If Check_Close Then Response.Write"<input type=""button"" value=""修改捐贈資料"" name=""Update"" class=""cbutton"" style=""cursor:hand"" onClick=""Update_OnClick()"">&nbsp;&nbsp;"
                                End If
                                If Cstr(RS("Invoice_Type"))=Cstr(InvoiceTypeY) Then
                                  Response.Write"<input type=""button"" value=""收據列印"" name=""Invoice"" class=""cbutton"" disabled >&nbsp;&nbsp;"
                                ElseIf RS("Issue_Type")<>"D" Then
                                  If RS("Invoice_Print")="1" Then
                                    Response.Write"<input type=""button"" value=""收據補印"" name=""Invoice"" class=""cbutton"" style=""cursor:hand"" onClick=""Invoice_OnClick()"">&nbsp;&nbsp;"
                                  Else
                                    Response.Write"<input type=""button"" value=""收據列印"" name=""Invoice"" class=""cbutton"" style=""cursor:hand"" onClick=""Invoice_OnClick()"">&nbsp;&nbsp;"
                                  End If
                                End If
                                If Check_Close Then
                                  If RS("Issue_Type")<>"D" Then
                                    Response.Write"<input type=""button"" value=""收據作廢"" name=""Void"" class=""cbutton"" style=""cursor:hand"" onClick=""Void_OnClick()"">&nbsp;&nbsp;"
                                  Else
                                    Response.Write"<input type=""button"" value=""還原作廢收據"" name=""Return"" class=""cbutton"" style=""cursor:hand"" onClick=""Return_OnClick()"">&nbsp;&nbsp;"
                                  End If
                                End If
                              %>
				                      <input type="button" value=" 取消 " name="cancel" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">&nbsp;&nbsp;
				                      <input type="button" value="續增本人捐贈資料" name="Add" class="addbutton" style="cursor:hand" onClick="Add_OnClick()">&nbsp;&nbsp;
				                      <input type="button" value="新增他人捐贈資料" name="Add" class="delbutton" style="cursor:hand" onClick="Add2_OnClick()">
                            </td>
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
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Update_OnClick(){
  location.href='contribute_edit.asp?ctype='+document.form.ctype.value+'&contribute_id='+document.form.contribute_id.value+'';
}
function Invoice_OnClick(){
  if(confirm('您是否確定要列印收據？')){
    window.open('contribute_invoice_print_rpt.asp?action=report&invoice_prog='+document.form.Invoice_Prog.value+'&contribute_id='+document.form.contribute_id.value+'','invoice_rpt','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600');
  }
}
function Void_OnClick(){
  if(document.form.donate_invoice.value=='N'&&document.form.Invoice_No.value==''){
    alert('年度收據無法作廢！');
    return;
  }
  if(document.form.RePrint.value==''){
    alert('作廢收據前請先列印！');
    return;
  }
  if(confirm('您是否確定要作廢收據？')){
    document.form.action.value='void';
    document.form.submit();
  }
}
function Return_OnClick(){
  if(confirm('您是否確定要還原作廢收據？')){
    document.form.action.value='return';
    document.form.submit();
  }
}
function Cancel_OnClick(){
  if(document.form.ctype.value=='contribute_data'){
    location.href='contribute_data.asp?donor_id='+document.form.Donor_Id.value+'';
  }else{
    location.href='contribute.asp';
  }
}	
function Add_OnClick(){
  location.href='contribute_add.asp?ctype=contribute_input&donor_id='+document.form.Donor_Id.value+'';
}
function Add2_OnClick(){
  location.href='contribute_input.asp';
}
--></script>	