<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="void" Then
  SQL1="Select * From DONATE Where Donate_Id='"&request("donate_id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  Donor_Id=RS1("Donor_Id")
  CodeKind=RS1("Donate_Purpose_Type")
  RS1("Donate_Amt_D")=RS1("Donate_Amt")
  RS1("Donate_Fee_D")=RS1("Donate_Fee")
  RS1("Donate_Accou_D")=RS1("Donate_Accou")
  RS1("Donate_Amt2_D")=RS1("Donate_Amt2")
  RS1("Donate_RateS_D")=RS1("Donate_RateS")
  RS1("Donate_Amt")="0"
  RS1("Donate_Fee")="0" 
  RS1("Donate_Accou")="0"
  RS1("Donate_Amt"&CodeKind&"")="0"
  RS1("Donate_Fee"&CodeKind&"")="0"
  RS1("Donate_Accou"&CodeKind&"")="0"
  If CodeKind="S" Then RS1("Donate_RateS")="0"
  RS1("Issue_Type")="D"
  RS1("Export")="N"
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  
  '修改捐款人捐款紀錄
  call Declare_DonorId (request("Donor_Id"))
    
  session("errnumber")=1
  session("msg")="收據已作廢 ！"
End If

If request("action")="return" Then
  SQL1="Select * From DONATE Where Donate_Id='"&request("donate_id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  CodeKind=RS1("Donate_Purpose_Type")
  Donate_RateS=RS1("Donate_RateS_D")
  RS1("Donate_Amt")=RS1("Donate_Amt_D")
  RS1("Donate_Fee")=RS1("Donate_Fee_D")
  RS1("Donate_Accou")=RS1("Donate_Accou_D")
  RS1("Donate_Amt"&CodeKind&"")=RS1("Donate_Amt_D")
  RS1("Donate_Fee"&CodeKind&"")=RS1("Donate_Fee_D")
  RS1("Donate_Accou"&CodeKind&"")=RS1("Donate_Accou_D")
  If CodeKind="S" Then RS1("Donate_RateS")=Donate_RateS
  RS1("Donate_Amt_D")="0"
  RS1("Donate_Fee_D")="0"
  RS1("Donate_Accou_D")="0"
  RS1("Donate_Amt2_D")="0"
  RS1("Donate_RateS_D")="0"
  RS1("Issue_Type")=RS1("Issue_Type_Keep")
  RS1("Export")="N"
  RS1.Update
  RS1.Close
  Set RS1=Nothing

  '修改捐款人捐款紀錄
  call Declare_DonorId (request("Donor_Id"))
    
  session("errnumber")=1
  session("msg")="收據已還原 ！"
End If

DonateDesc="物資捐贈"
SQL1="Select DonateDesc=CodeDesc From CASECODE Where CodeType='Payment' And CodeDesc Like '%物%' Order By Seq"
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
'20131202 Modify by GoodTV Tanya:增加顯示捐款人編號
SQL="Select DONATE.*,Donor_Id=DONOR.Donor_Id,Donor_Name=DONOR.Donor_Name,Title=DONOR.Title,Category=DONOR.Category,Donor_Type=DONOR.Donor_Type,IDNo=DONOR.Invoice_IDNo, " & _
    "Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+DONOR.Address Else DONOR.ZipCode+A.mValue+DONOR.Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then DONOR.Address Else Case When C.mValue<>D.mValue Then DONOR.Invoice_ZipCode+C.mValue+D.mValue+DONOR.Invoice_Address Else DONOR.Invoice_ZipCode+C.mValue+DONOR.Address End End), " & _
    "DEPT.Comp_ShortName,ACT.Act_ShortName,DEPT.Donate_Invoice,DEPT.Invoice_Prog,DEPT.Rept_Licence " &_ 
    "From DONATE " & _
    "Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id " & _
    "Join Dept On DONATE.Dept_Id=DEPT.Dept_Id " & _
    "Left Join ACT On DONATE.Act_Id=ACT.Act_Id " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _
    "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
    "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
    "Where Donate_Id='"&request("donate_id")&"' "
Call QuerySQL(SQL,RS)

Check_Close=Get_Close("1",Cstr(RS("Dept_Id")),Cstr(RS("Donate_Date")),Cstr(Session("user_id")))
%>
<%Prog_Id="donate"%>
<!--#include file="../include/head.asp"-->
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
    	<input type="hidden" name="action" value="">
    	<input type="hidden" name="ctype" value="<%=request("ctype")%>">
    	<input type="hidden" name="DonorId" value="<%=RS("Donor_Id")%>">
      <input type="hidden" name="donate_id" value="<%=RS("Donate_Id")%>">
      <input type="hidden" name="donate_invoice" value="<%=RS("Donate_Invoice")%>">
      <input type="hidden" name="Invoice_Prog" value="<%=RS("Invoice_Prog")%>">
      <input type="hidden" name="Rept_Licence" value="<%=RS("Rept_Licence")%>">
      <input type="hidden" name="DonateDesc" value="<%=DonateDesc%>">		
      <%If Cstr(RS("Donate_Payment"))=Cstr(DonateDesc) Then%>	
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
                        	<!--<tr>
                            <td align="right">收據編號：</td>
                            <td align="left" colspan="7">                           
                              <input type="text" name="Category" size="20" class="font9t" readonly value="<%=RS("Invoice_Pre")+RS("Invoice_No")%>">                         
                            </td>
                          </tr>-->
						  						<tr>
                            <td align="right">捐款人：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Donor_Name" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Donor_Name"))%>&nbsp;<%=RS("Title")%>">
                              <!--20131202 Modify by GoodTV Tanya:隱藏「類別」並增加顯示「捐款人編號」-->
                              <!--&nbsp;類&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;別：
                              <input type="text" name="Category" size="10" class="font9t" readonly value="<%=RS("Category")%>">-->
                              &nbsp;&nbsp;捐款人編號：
                              <input type="text" name="Donor_Id" size="10" class="font9t" readonly value="<%=RS("Donor_Id")%>">
                              &nbsp;身份別：
                              <input type="text" name="Donor_Type" size="40" class="font9t" readonly value="<%=RS("Donor_Type")%>">
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
                            <td width="10%" align="right">捐款日期：</td>
                            <td width="15% align="left"><input type="text" name="IDNo" size="12" class="font9t" readonly value="<%=RS("Donate_Date")%>"></td>
                            <td width="10%" align="right">捐款方式：</td>
                            <td width="16%"><input type="text" name="Donate_Payment" size="15" class="font9t" readonly value="<%=RS("Donate_Payment")%>"></td>
                            <td width="10%" align="right">捐款用途：</td>
                            <td width="15%"><input type="text" name="Donate_Purpose" size="14" class="font9t" readonly value="<%=RS("Donate_Purpose")%>"></td>
                            <td width="10%" align="right">收據開立：</td>
                            <td width="14%"><input type="text" name="Invoice_Type" size="12" class="font9t" readonly value="<%=RS("Invoice_Type")%>"></td>
                          </tr>
                          <%
                            If Cstr(RS("Donate_Payment"))<>Cstr(DonateDesc) Then
                              Response.Write "<tr id=""donateamt"" style=""display:block"">"
                            Else
                              Response.Write "<tr id=""donateamt"" style=""display:none"">"
                            End If
                          %>
                            <td align="right">捐款金額：</td>
                            <td><input type="text" name="Donate_Amt" size="12" class="font9t" readonly value="<%if isnumeric(RS("Donate_Amt")) Then Response.write FormatNumber(RS("Donate_Amt"),0)%>" style="text-align: right"></td>
                            <td align="right">手續費：</td>
                            <td><input type="text" name="Donate_Fee" size="15" class="font9t" readonly value="<%if isnumeric(RS("Donate_Fee")) Then Response.write FormatNumber(RS("Donate_Fee"),0)%>" style="text-align: right"></td>
                            <td align="right">實收金額：</td>
                            <td><input type="text" name="Donate_Accou" size="14" class="font9t" readonly value="<%if isnumeric(RS("Donate_Accou")) Then Response.write FormatNumber(RS("Donate_Accou"),0)%>" style="text-align: right"></td>
                            <td align="right">外幣：</td>
                            <td><input type="text" name="Donate_Forign" size="12" class="font9t" readonly value="<%=RS("Donate_Forign")%>"></td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <%
                            Check_Line=False
                            If Instr(Cstr(RS("Donate_Payment")),"信用卡")>0 And (Instr(Cstr(RS("Donate_Payment")),"扣款")>0 Or Instr(Cstr(RS("Donate_Payment")),"授權書")>0) Then
                              Check_Line=True
                              Response.Write "<tr id=""donatecard1"" style=""display:block"">"
                            Else
                              Response.Write "<tr id=""donatecard1"" style=""display:none"">"
                            End If
                          %>
                            <td align="right">銀行別：</td>
                            <td align="left"><input type="text" name="Card_Bank" size="12" class="font9t" readonly value="<%=RS("Card_Bank")%>"></td></td>
                            <td align="right">信用卡別：</td>
                            <td colspan="3">
                            	<input type="text" name="Card_Type" size="12" class="font9t" readonly value="<%=RS("Card_Type")%>">
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;信用卡號:
                              <input type="text" name="Account_No" size="19" class="font9t" readonly value="<%=RS("Account_No")%>">
                            </td>	
                            <td align="right">有效月年：</td>
                            <td><input type="text" name="Valid_Date" size="12" class="font9t" readonly value="<%=RS("Valid_Date")%>"></td>
                          </tr>
                          <%
                            If Instr(Cstr(RS("Donate_Payment")),"信用卡")>0 And (Instr(Cstr(RS("Donate_Payment")),"扣款")>0 Or Instr(Cstr(RS("Donate_Payment")),"授權書")>0) Then
                              Check_Line=True
                              Response.Write "<tr id=""donatecard2"" style=""display:block"">"
                            Else
                              Response.Write "<tr id=""donatecard2"" style=""display:none"">"
                            End If
                          %>
                            <td align="right">持卡人：</td>
                            <td align="left"><input type="text" name="Card_Owner" size="12" class="font9t" readonly value="<%=RS("Card_Owner")%>"></td>
                            <td colspan="4">
                            	<input type="checkbox" name="Same_Owner" value="Y" disabled >同捐款人
                              &nbsp;&nbsp;&nbsp;持卡人身分證：
                              <input type="text" name="Owner_IDNo" size="10" class="font9t" readonly value="<%=RS("Owner_IDNo")%>">
                              &nbsp;&nbsp;&nbsp;與捐款人關係：
                              <input type="text" name="Relation" size="4" class="font9t" readonly value="<%=RS("Relation")%>">
                            </td> 
                            <td align="right">授權碼：</td>
                            <td align="left"><input type="text" name="Authorize" size="12" class="font9t" readonly value="<%=RS("Authorize")%>"></td>
                          </tr>
                          <%
                            If Instr(Cstr(RS("Donate_Payment")),"支票")>0 Then
                              Check_Line=True
                              Response.Write "<tr id=""checkno"" style=""display:block"">"
                            Else
                              Response.Write "<tr id=""checkno"" style=""display:none"">"
                            End If
                          %>
                            <td align="right">支票號碼：</td>
                            <td><input type="text" name="Check_No" size="12" class="font9t" readonly value="<%=RS("Check_No")%>"></td>
                            <td align="right">到期日：</td>
                            <td colspan="5"><input type="text" name="Check_ExpireDate" size="12" class="font9t" readonly value="<%=RS("Check_ExpireDate")%>"></td>
                          </tr>
                          <%
                            If Instr(Cstr(RS("Donate_Payment")),"郵局")>0 And (Instr(Cstr(RS("Donate_Payment")),"轉帳")>0 Or Instr(Cstr(RS("Donate_Payment")),"授權書")>0) Then
                              Check_Line=True
                              Response.Write "<tr id=""post"" style=""display:block"">"	
                            Else
                              Response.Write "<tr id=""post"" style=""display:none"">"
                            End If
                          %>
                            <td align="right">存簿戶名：</td>
                            <td><input type="text" name="Post_Name" size="12" class="font9t" readonly value="<%=RS("Post_Name")%>"></td>
                            <td colspan="6">
                            	<input type="checkbox" name="Same_Post" value="Y" readonly >同捐款人
                            	&nbsp;&nbsp;&nbsp;存簿局號：
                            	<input type="text" name="Post_SavingsNo" size="12" class="font9t" readonly value="<%=RS("Post_SavingsNo")%>">
                              &nbsp;&nbsp;&nbsp;存簿帳號：
                              <input type="text" name="Post_AccountNo" size="12" class="font9t" readonly value="<%=RS("Post_AccountNo")%>">
                            </td>
                            <td colspan="2"> </td>
                          </tr>
                          <%If Check_Line Then%>
                          <tr id="line" style="display:block">
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <%End If%>
                          <tr>
                            <td align="right">機構名稱：</td>
                            <td><input type="text" name="Comp_ShortName" size="12" class="font9t" readonly value="<%=RS("Comp_ShortName")%>"></td>
                            <td align="right">收據抬頭：</td>
                            <td colspan="2"><input type="text" name="Invoice_Title" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Invoice_Title"))%>"></td>
                            <td><input type="checkbox" name="Issue_Type" value="M" <%If RS("Issue_Type")="M" Then Response.Write "checked" End If %> disabled >手開收據</td>
                            <td align="right">收據編號：</td>
                            <td><input type="text" name="Invoice_No" size="12" class="font9t" readonly value="<%=RS("Invoice_Pre")&RS("Invoice_No")%>"></td>
                          </tr> 
                          <tr>
                            <td align="right">請款日期：</td>
                            <td align="left"><input type="text" name="Request_Date" size="12" class="font9t" readonly value="<%=RS("Request_Date")%>"></td>
                            <td align="right">入帳銀行：</td>
                            <td><input type="text" name="Accoun_Bank" size="15" class="font9t" readonly value="<%=RS("Accoun_Bank")%>"></td>
                            <td align="right">沖帳日期：</td>
                            <td><input type="text" name="Accoun_Date" size="15" class="font9t" readonly value="<%=RS("Accoun_Date")%>"></td>
                            <td align="right">捐款類別：</td>
                            <td><input type="text" name="Donate_Type" size="12" class="font9t" readonly value="<%=RS("Donate_Type")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">傳票號碼：</td>
                            <td align="left"><input type="text" name="Donation_NumberNo" size="12" class="font9t" readonly value="<%=RS("Donation_NumberNo")%>"></td> 
                            <td align="left" colspan="2">&nbsp;&nbsp;劃撥&nbsp;/&nbsp;匯款單號：<input type="text" name="Donation_SubPoenaNo" size="11" class="font9t" readonly value="<%=RS("Donation_SubPoenaNo")%>"></td>
                            <td align="right">會計科目：</td>
                            <td align="left"><input type="text" name="Accounting_Title" size="15" class="font9t" readonly value="<%=RS("Accounting_Title")%>"></td>
                          	<td align="right">收據寄送：</td>
                            <td align="left"><input type="text" name="InvoceSend_Date" size="12" class="font9t" readonly value="<%=RS("InvoceSend_Date")%>"></td> 
                          </tr>
                          <tr>
                            <td align="right">專案活動：</td>
                            <td align="left"><input type="text" name="Act_ShortName" size="12" class="font9t" readonly value="<%=RS("Act_ShortName")%>"></td>
							<td align="center">收據備註：(列印用)</td>
                            <td align="left" colspan="5"><input type="text" name="Invoice_PrintComment" size="15" class="font9t" readonly value="<%=RS("Invoice_PrintComment")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">捐款備註：</td>
                            <td align="left" colspan="6"><textarea rows="3" name="Comment" cols="44" class="font9t" readonly ><%=RS("Comment")%></textarea></td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <!--20131120 GoodTV Tanya:顯示收據列印日期-->
                          <tr>
                          	<td align="right">首印日期：</td>
                          	<td align="left" colspan="2"><input type="text" rows="3" name="Invoice_Print_Date" class="font9t" readonly value="<%=RS("Invoice_Print_Date")%>"></td>
                          	<td align="right" colspan="2">最後補印日期：</td>
                          	<td align="left" colspan="3"><input type="text" rows="3" name="Invoice_RePrint_Date" class="font9t" readonly value="<%=RS("Invoice_RePrint_Date")%>"></td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="center" colspan="8">
                            	<!--20140108 GoodTV Tanya:收據狀態Issue_Type為空時仍可「修改捐款資料」-->
                              <%
                                If RS("Issue_Type")<>"D" Then
                                  If Check_Close Then Response.Write"<input type=""button"" value=""修改捐款資料"" name=""Update"" class=""cbutton"" style=""cursor:hand"" onClick=""Update_OnClick()"">&nbsp;&nbsp;"
                                End If
                                If Cstr(RS("Invoice_Type"))=Cstr(InvoiceTypeY) Then
                                  Response.Write"<input type=""button"" value=""收據列印"" name=""Invoice"" class=""cbutton"" disabled >&nbsp;&nbsp;"
                                ElseIf RS("Issue_Type")<>"D" or isnull(RS("Issue_Type")) Then
                                  If RS("Invoice_Print")="1" Then
                                    Response.Write"<input type=""button"" value=""收據補印"" name=""Invoice"" class=""cbutton"" style=""cursor:hand"" onClick=""Invoice_OnClick()"">&nbsp;&nbsp;"
                                  Else
                                    Response.Write"<input type=""button"" value=""收據列印"" name=""Invoice"" class=""cbutton"" style=""cursor:hand"" onClick=""Invoice_OnClick()"">&nbsp;&nbsp;"
                                  End If
                                End If
                                If Check_Close Then
                                  If RS("Issue_Type")<>"D" or isnull(RS("Issue_Type")) Then
                                    Response.Write"<input type=""button"" value=""收據作廢"" name=""Void"" class=""cbutton"" style=""cursor:hand"" onClick=""Void_OnClick()"">&nbsp;&nbsp;"
                                    Response.Write"<input type=""button"" value=""重新取號"" name=""Return"" class=""rbutton"" style=""cursor:hand"" onClick=""New_InvoiceNo_OnClick()"">&nbsp;&nbsp;"
                                  Else
                                    Response.Write"<input type=""button"" value=""還原作廢收據"" name=""Return"" class=""cbutton"" style=""cursor:hand"" onClick=""Return_OnClick()"">&nbsp;&nbsp;"
                                  End If
                                End If
                              %>
				                      <input type="button" value=" 取消 " name="cancel" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">&nbsp;&nbsp;
				                      <input type="button" value="續增本人捐款資料" name="Add" class="addbutton" style="cursor:hand" onClick="Add_OnClick()">&nbsp;&nbsp;
				                      <input type="button" value="新增他人捐款資料" name="Add" class="delbutton" style="cursor:hand" onClick="Add2_OnClick()">
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
  location.href='donate_edit.asp?ctype='+document.form.ctype.value+'&donate_id='+document.form.donate_id.value+'';
}
function Invoice_OnClick(){
  if(confirm('您是否確定要列印收據？')){
    window.open('invoice_print_rpt.asp?action=report&invoice_prog='+document.form.Invoice_Prog.value+'&donate_id='+document.form.donate_id.value+'&print_pint='+document.form.Print_Pint.value+'&reprint='+document.form.RePrint.value+'&donatedesc='+document.form.DonateDesc.value+'&Rept_Licence='+document.form.Rept_Licence.value+'','invoice_rpt','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600');
  }
}
function Void_OnClick(){
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
function New_InvoiceNo_OnClick(){
  location.href='donate_new_invoiceno.asp?ctype='+document.form.ctype.value+'&donate_id='+document.form.donate_id.value+'';
}
function Cancel_OnClick(){
  if(document.form.ctype.value=='donate_data'){
    location.href='donate_data.asp?donor_id='+document.form.Donor_Id.value+'';
  }else{
    location.href='donate.asp';
  }
}	
function Add_OnClick(){
  location.href='donate_add.asp?ctype=donate_input&donor_id='+document.form.Donor_Id.value+'';
}
function Add2_OnClick(){
  location.href='donate_input.asp';
}
--></script>	