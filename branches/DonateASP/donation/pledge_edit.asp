<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  CodeKind=M
  SQL1="Select CodeKind From CASECODE Where CodeType='Purpose' And CodeDesc='"&request("Donate_Purpose")&"' "
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  CodeKind=RS1("CodeKind")
  RS1.Close
  Set RS1=Nothing
  
  SQL1="Select * From PLEDGE Where Pledge_id='"&request("pledge_id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  Donor_Id=RS1("Donor_Id")
  RS1("Donate_Payment")=request("Donate_Payment")
  RS1("Donate_Purpose")=request("Donate_Purpose")
  RS1("Donate_Purpose_Type")=CodeKind
  If request("Donate_Period")="單筆" Then
    RS1("Donate_Type")="單次捐款"
  Else
    RS1("Donate_Type")="長期捐款"
  End If
  If request("Donate_Amt")<>"" Then
    RS1("Donate_Amt")=request("Donate_Amt")
  Else
    RS1("Donate_Amt")="0"
  End If
'20131105 Mark by GoodTV Tanya:目前用不到先隱藏「下次強制扣款」  
'  If request("Donate_FirstAmt")<>"" Then
'    RS1("Donate_FirstAmt")=request("Donate_FirstAmt")
'  Else
'    RS1("Donate_FirstAmt")="0"
'  End If
 	RS1("Donate_FirstAmt")="0"
  If request("Donate_FromDate")<>"" Then
    RS1("Donate_FromDate")=request("Donate_FromDate")
  Else
    RS1("Donate_FromDate")=null
  End If
  If request("Donate_ToDate")<>"" Then
    RS1("Donate_ToDate")=request("Donate_ToDate")
  Else
    RS1("Donate_ToDate")=null
  End If
  RS1("Donate_Period")=request("Donate_Period")   
  If request("Next_DonateDate")<>"" Then
    RS1("Next_DonateDate")=request("Next_DonateDate")
  Else
    RS1("Next_DonateDate")=null
  End If
  RS1("Card_Bank")=request("Card_Bank")
  RS1("Card_Type")=request("Card_Type")
  RS1("Account_No")=request("Account_No1")&request("Account_No2")&request("Account_No3")&request("Account_No4")
  RS1("Valid_Date")=request("Valid_Month")&request("Valid_Year")
  RS1("Card_Owner")=Data_Plus(request("Card_Owner"))
  RS1("Owner_IDNo")=request("Owner_IDNo")
  RS1("Relation")=request("Relation")
  RS1("Authorize")=request("Authorize")
  RS1("Post_Name")=Data_Plus(request("Post_Name"))
  RS1("Post_IDNo")=request("Post_IDNo")
  RS1("Post_SavingsNo")=request("Post_SavingsNo")
  RS1("Post_AccountNo")=request("Post_AccountNo")
  RS1("Dept_Id")=request("Dept_Id")
  RS1("Invoice_Type")=request("Invoice_Type")    
  RS1("Invoice_Title")=Data_Plus(request("Invoice_Title")) 
  RS1("Accoun_Bank")=request("Accoun_Bank")
  RS1("Accounting_Title")=request("Accounting_Title")
  If request("Act_Id")<>"" Then
    RS1("Act_Id")=request("Act_Id")
  Else
    RS1("Act_Id")=null
  End If
  RS1("Status")=request("Status")
  If request("Status")="停止" Then
    If request("Break_Date")<>"" Then
      RS1("Break_Date")=request("Break_Date")
    Else
      RS1("Break_Date")=null
    End If
    RS1("Break_Reason")=request("Break_Reason")
  Else
    RS1("Break_Date")=null
    RS1("Break_Reason")=""
  End If
  '20140305 Mark by GoodTV Tanya:修正「停止原因」無法儲存的問題。多餘的程式碼清空了「停止原因」內容
'  RS1("Break_Reason")=""
  RS1("Comment")=request("Comment")
  RS1("LastUpdate_Date")=Date()
  RS1("LastUpdate_DateTime")=Now()
  RS1("LastUpdate_User")=session("user_name")
  RS1("LastUpdate_IP")=Request.ServerVariables("REMOTE_HOST")   
  RS1("P_BANK")=request("P_BANK")
  RS1("P_RCLNO")=request("P_RCLNO")
  RS1("P_PID")=request("P_PID")
  RS1.Update
  RS1.Close
  Set RS1=Nothing

  SQL="DECLARE @last_pledgedate datetime " & _
      "Select Top 1 @last_pledgedate=Donate_ToDate From PLEDGE Where Donor_Id='"&Donor_Id&"' Order By Donate_ToDate Desc " & _
      "Update DONOR Set Last_PledgeDate=@last_pledgedate Where Donor_Id='"&Donor_Id&"'"
  Set RS=Conn.Execute(SQL)
  session("errnumber")=1
  session("msg")="轉帳授權書修改成功 !"
  If request("ctype")="pledge_data" Then Response.Redirect "pledge_data.asp?donor_id="&request("donor_id")
End If

If request("action")="delete" Then
  SQL="Delete From PLEDGE Where Pledge_Id='"&request("pledge_id")&"'"
  Set RS=Conn.Execute(SQL)
  
  SQL="DECLARE @last_pledgedate datetime " & _
      "Select Top 1 @last_pledgedate=Donate_ToDate From PLEDGE Where Donor_Id='"&request("Donor_Id")&"' Order By Donate_ToDate Desc " & _
      "Update DONOR Set Last_PledgeDate=@last_pledgedate Where Donor_Id='"&request("Donor_Id")&"'"
  Set RS=Conn.Execute(SQL)
  
  session("errnumber")=1
  session("msg")="轉帳授權書刪除成功 !"
  If request("ctype")="donate_data" Then
    Response.Redirect "pledge_data.asp?donor_id="&request("donor_id")
  Else
    Response.Redirect "pledge.asp"
  End If  
End If

SQL_IS="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"' "
Set RS_IS=Server.CreateObject("ADODB.Recordset")
RS_IS.Open SQL_IS,Conn,1,1
IsMember=RS_IS("IsMember")
IsShopping=RS_IS("IsShopping")
RS_IS.Close
Set RS_IS=Nothing
CodeKind="'D'"
If IsMember="Y" Then CodeKind=CodeKind&",'M','A'"
If IsShopping="Y" Then CodeKind=CodeKind&",'S'"

SQL="Select PLEDGE.*,Donor_Name=DONOR.Donor_Name,Title=DONOR.Title,Category=DONOR.Category,Donor_Type=DONOR.Donor_Type,InvoiceType=DONOR.Invoice_Type,IDNo=DONOR.IDNo,Last_DonateDateD=DONOR.Last_DonateDateD,Remark=DONOR.Remark, " & _
    "Address2=DONOR.ZipCode+A.mValue+B.mValue+DONOR.Address,Invoice_Address2=DONOR.Invoice_ZipCode+C.mValue+D.mValue+DONOR.Invoice_Address " & _
    "From PLEDGE " & _
    "Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id " & _
    "Join Dept On PLEDGE.Dept_Id=DEPT.Dept_Id " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _
    "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
    "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
    "Where Pledge_Id='"&request("pledge_id")&"' "
Call QuerySQL(SQL,RS)
%>
<%Prog_Id="pledge"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <input type="hidden" name="Pledge_Id" value="<%=RS("Pledge_Id")%>">	
      <input type="hidden" name="DonorId" value="<%=RS("Donor_Id")%>">
      <input type="hidden" name="DonorName" value="<%=RS("Donor_Name")%>">
      <input type="hidden" name="DonorIDNo" value="<%=RS("IDNo")%>">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;"><%=Prog_Desc%>【修改】</td>
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
	          <table width="900" border="0" cellspacing="0" cellpadding="0" align="center">
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
                        <table border="0" cellpadding="2" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
						    <td align="right">授權編號:</td>                    
                            <td align="left" colspan="7">
                            <input type="text" name="Pledge_Id" size="10" class="font9t" readonly value="<%=RS("Pledge_Id")%>">							
                            </td>
                          </tr>
						  <tr>												  							
                            <td align="right">捐款人：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Donor_Name" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Donor_Name"))%>&nbsp;<%=RS("Title")%>">
                              <!--20131111Modify by GoodTV Tanya:隱藏「類別」and增加「捐款人編號」-->
                              <!--&nbsp;類&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;別：
                              <input type="text" name="Category" size="10" class="font9t" readonly value="<%=RS("Category")%>">-->
                              捐款人編號：                   	
                            	<input type="text" name="Donor_Id" size="9" class="font9t" readonly value="<%=(RS("Donor_Id"))%>">
                              &nbsp;身份別：
                              <input type="text" name="Donor_Type" size="40" class="font9t" readonly value="<%=RS("Donor_Type")%>">
                            </td>
                          </tr>
                          <tr>
                            <td align="right">收據地址：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Invoice_Address2" size="30" class="font9t" readonly value="<%=Data_Plus(RS("Invoice_Address2"))%>">
                              &nbsp;收據開立：
                              <input type="text" name="InvoiceType" size="10" class="font9t" readonly value="<%=RS("InvoiceType")%>">
                              &nbsp;身分證/統編：
                              <input type="text" name="IDNo" size="10" class="font9t" readonly value="<%=RS("IDNo")%>">
                              &nbsp;最近捐款日：
                              <input type="text" name="Last_DonateDate" size="10" class="font9t" readonly value="<%=RS("Last_DonateDateD")%>">	
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td width="10%" align="right">授權方式：</td>
                            <td width="15%">
                            <%
                              SQL1="Select Donate_Payment=CodeDesc From CASECODE Where CodeType='Payment' And CodeDesc Like '%授權書%' Order By Seq"
                              Set RS1 = Server.CreateObject("ADODB.RecordSet")
                              RS1.Open SQL1,Conn,1,1
                              Response.Write "<SELECT Name='Donate_Payment' size='1' style='font-size: 9pt; font-family: 新細明體' OnChange=""Donate_Payment_OnChange(this.value)"">"
                              Response.Write "<OPTION>" & " " & "</OPTION>"
                              While Not RS1.EOF
                                If Cstr(RS("Donate_Payment"))=Cstr(RS1("Donate_Payment")) Then
                                  Response.Write "<OPTION value='"&RS1("Donate_Payment")&"' selected >"&RS1("Donate_Payment")&"</OPTION>"  
                                Else  
                                  Response.Write "<OPTION value='"&RS1("Donate_Payment")&"'>"&RS1("Donate_Payment")&"</OPTION>"
                                End If
                                RS1.MoveNext
                              Wend
                              Response.Write "</SELECT>"
                              RS1.Close
                              Set RS1=Nothing
                            %>
                            </td>
                            <td width="10%" align="right">指定用途：</td>
                            <td width="15%">
                            <%
                              SQL="Select Donate_Purpose=CodeDesc From CASECODE Where CodeType='Purpose' And CodeKind In ("&CodeKind&") Order By Seq"
                              FName="Donate_Purpose"
                              Listfield="Donate_Purpose"
                              menusize="1"
                              BoundColumn=RS("Donate_Purpose")
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                            </td>
                            <td width="10%" align="right">每次扣款：</td>
                            <td width="14% align="left"><input type="text" name="Donate_Amt" size="12" class="font9" maxlength="7" value="<%=RS("Donate_Amt")%>"></td>
                            <!--20131105 Mark by GoodTV Tanya:目前用不到先隱藏「下次強制扣款」-->
                            <!--<td width="13%" align="right">下次強制扣款：</td>
                            <td width="13%"><input type="text" name="Donate_FirstAmt" size="12" class="font9" maxlength="7" value="<%=RS("Donate_FirstAmt")%>"></td>-->
                          </tr>
                          <tr>
                            <td align="right">授權期間：</td>
                            <td colspan="3"><%call Calendar("Donate_FromDate",RS("Donate_FromDate"))%> ~ <%call Calendar("Donate_ToDate",RS("Donate_ToDate"))%></td>
                            <td align="right">轉帳週期：</td>
                            <td>
                      	      <SELECT Name='Donate_Period' size='1' style='font-size: 9pt; font-family: 新細明體'>
                      		      <OPTION  value='單筆' <%If RS("Donate_Period")="單筆" Then Response.Write "selected" End If%> >單筆</OPTION>
                      		      <OPTION  value='月繳' <%If RS("Donate_Period")="月繳" Then Response.Write "selected" End If%> >月繳</OPTION>
                      		      <OPTION  value='季繳' <%If RS("Donate_Period")="季繳" Then Response.Write "selected" End If%> >季繳</OPTION>
                      		      <OPTION  value='半年繳' <%If RS("Donate_Period")="半年繳" Then Response.Write "selected" End If%> >半年繳</OPTION>
                      		      <OPTION  value='年繳' <%If RS("Donate_Period")="年繳" Then Response.Write "selected" End If%> >年繳</OPTION>
                      	      </SELECT>
                            </td>
                            <!--20131105 Modify by GoodTV Tanya:皆顯示為「下次扣款日」-->
                            <!--<%If RS("Transfer_Date")<>"" Then%>
                            <td align="right">下次扣款日：</td>
                            <!--<%Else%>
                            <td align="right">首次扣款日：</td>
                            <%End if%>-->
                            <td align="right">下次扣款日：</td>
                            <td><%call Calendar("Next_DonateDate",RS("Next_DonateDate"))%></td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <%
                            Check_Line=False
                            If Instr(Cstr(RS("Donate_Payment")),"信用卡")>0 And (Instr(Cstr(RS("Donate_Payment")),"扣款")>0 Or Instr(Cstr(RS("Donate_Payment")),"授權書")>0) Then
                              Check_Line=True
                              Response.Write "<tr id=""donatecard1"" style=""display:"">"
                            Else
                              Response.Write "<tr id=""donatecard1"" style=""display:none"">"
                            End If
                          %>
                            <td align="right">銀行別：</td>
                            <td align="left"><input type="text" name="Card_Bank" size="12" class="font9" maxlength="20" value="<%=RS("Card_Bank")%>"></td></td>
                            <td align="right">信用卡別：</td>
                            <td colspan="3">
                            <%
                              SQL="Select Card_Type=CodeDesc From CASECODE Where CodeType='CardType' Order By Seq"
                              FName="Card_Type"
                              Listfield="Card_Type"
                              menusize="1"
                              BoundColumn=RS("Card_Type")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              
                              Account_No1=""
                              Account_No2=""
                              Account_No3=""
                              Account_No4=""
                              If RS("Account_No")<>"" Then
                                If Len(RS("Account_No"))>=4 Then Account_No1=Left(RS("Account_No"),4)
                                If Len(RS("Account_No"))>=8 Then Account_No2=Mid(RS("Account_No"),5,4)
                                If Len(RS("Account_No"))>=12 Then Account_No3=Mid(RS("Account_No"),9,4)
                                If Len(RS("Account_No"))>=16 Then Account_No4=Right(RS("Account_No"),4)
                                '20130923 Modify by GoodTV Tanya:修正AE僅15碼後3碼未顯示的問題
                                If Len(RS("Account_No"))=15 Then Account_No4=Right(RS("Account_No"),3)
                              End If
                            %>
                              &nbsp;信用卡號:
                              <input type="text" name="Account_No1" size="3" class="font9" maxlength="4" OnKeyup="Account_No_Keyup('1');" value="<%=Account_No1%>">
                              -	
                              <input type="text" name="Account_No2" size="3" class="font9" maxlength="4" OnKeyup="Account_No_Keyup('2');" value="<%=Account_No2%>">
                              -	
                              <input type="text" name="Account_No3" size="3" class="font9" maxlength="4" OnKeyup="Account_No_Keyup('3');" value="<%=Account_No3%>">
                              -	
                              <input type="text" name="Account_No4" size="3" class="font9" maxlength="4" value="<%=Account_No4%>" onblur="<%call CheckCreditCardNumJ("Card_Type","Account_No1","Account_No2","Account_No3","Account_No4")%>">			
                            </td>
                            <td align="right">有效月年：</td>
                            <td>
                            <%
                              ValidMonth=0
                              ValidYear=0
                              If RS("Valid_Date")<>"" Then
                                If Len(RS("Valid_Date"))>=2 Then ValidMonth=Cint(Left(RS("Valid_Date"),2))
                                If Len(RS("Valid_Date"))>=4 Then ValidYear=Cint(Right(RS("Valid_Date"),2))
                              End If
                              Response.Write "<SELECT Name='Valid_Month' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                              Response.Write "<OPTION>" & " " & "</OPTION>"
                              For I= 1 To 12
                                If Len(I)=1 Then
                                  If Cint(ValidMonth)=Cint(I) Then
                                    Response.Write "<OPTION value='0"&I&"' selected >0"&I&"</OPTION>"
                                  Else
                                    Response.Write "<OPTION value='0"&I&"'>0"&I&"</OPTION>"
                                  End If
                                Else
                                  If Cint(ValidMonth)=Cint(I) Then
                                    Response.Write "<OPTION value='"&I&"' selected >"&I&"</OPTION>"
                                  Else
                                    Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
                                  End If
                                End If
                              Next
                              Response.Write "</SELECT>/"
                              Response.Write "<SELECT Name='Valid_Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                              Response.Write "<OPTION>" & " " & "</OPTION>"
                              For I= Year(Date())-10 To Year(Date())+15
                                If Cint(ValidYear)=Cint(Right(I,2)) Then 
                                  Response.Write "<OPTION value='"&Right(I,2)&"' selected >"&Right(I,2)&"</OPTION>"
                                Else
                                  Response.Write "<OPTION value='"&Right(I,2)&"'>"&Right(I,2)&"</OPTION>"
                                End If
                              Next
                              Response.Write "</SELECT>"
                            %>
                            </td>
                          </tr>
                          <%
                            If Instr(Cstr(RS("Donate_Payment")),"信用卡")>0 And (Instr(Cstr(RS("Donate_Payment")),"扣款")>0 Or Instr(Cstr(RS("Donate_Payment")),"授權書")>0) Then
                              Check_Line=True
                              Response.Write "<tr id=""donatecard2"" style=""display:"">"
                            Else
                              Response.Write "<tr id=""donatecard2"" style=""display:none"">"
                            End If
                          %>
                            <td align="right">持卡人：</td>
                            <td align="left"><input type="text" name="Card_Owner" size="12" class="font9" maxlength="40" value="<%=Data_Minus(RS("Card_Owner"))%>"></td>
                            <td colspan="6">
                            	<input type="checkbox" name="Same_Owner" value="Y" OnClick="Same_Owner_OnClick()">同捐款人
                              &nbsp;持卡人身分證：
                              <input type="text" name="Owner_IDNo" size="9" class="font9" maxlength="10" onKeyUp="UCaseOwnerIDNO();" onblur="<%call CheckIDJ("Owner_IDNo")%>" value="<%=RS("Owner_IDNo")%>">
                              &nbsp;與捐款人關係：
                              <input type="text" name="Relation" size="8" class="font9" maxlength="10">
                              &nbsp;&nbsp;&nbsp;授權碼：
                              <input type="text" name="Authorize" size="12" class="font9" maxlength="5" value="<%=RS("Authorize")%>">
                            </td>
                          </tr>
                          <%
                            If Instr(Cstr(RS("Donate_Payment")),"郵局")>0 And (Instr(Cstr(RS("Donate_Payment")),"轉帳")>0 Or Instr(Cstr(RS("Donate_Payment")),"授權書")>0) Then
                              Check_Line=True
                              Response.Write "<tr id=""post"" style=""display:"">"
                            Else
                              Response.Write "<tr id=""post"" style=""display:none"">"
                            End If
                          %>
                            <td align="right">存簿戶名：</td>
                            <td><input type="text" name="Post_Name" size="12" class="font9" maxlength="40" value="<%=Data_Minus(RS("Post_Name"))%>"></td>
                            <td colspan="6">
                              &nbsp;&nbsp;持有人身分證：
                              <input type="text" name="Post_IDNo" size="9" class="font9" maxlength="10" onKeyUp="UCasePostIDNO();" value="<%=RS("Post_IDNo")%>">
                            	<input type="checkbox" name="Same_Post" value="Y" OnClick="Same_Post_OnClick()">同捐款人
                            	&nbsp;&nbsp;存簿局號：
                            	<input type="text" name="Post_SavingsNo" size="12" class="font9" maxlength="7" value="<%=RS("Post_SavingsNo")%>">
                              &nbsp;&nbsp;存簿帳號：
                              <input type="text" name="Post_AccountNo" size="12" class="font9" maxlength="7" value="<%=RS("Post_AccountNo")%>">
                            </td>
                            <td colspan="2"> </td>
                          </tr>
                          <%
                            If Instr(Cstr(RS("Donate_Payment")),"ACH轉帳授權書")>0 Then
                              Check_Line=True
                              Response.Write "<tr id=""bank"" style=""display:"">"
                            Else
                              Response.Write "<tr id=""bank"" style=""display:none"">"
                            End If
                          %>
                            <td colspan="8">
                            	收受行代號：
                            	<input type="text" name="P_BANK" size="10" class="font9" maxlength="7" value="<%=RS("P_BANK")%>">
                              &nbsp;&nbsp;&nbsp;&nbsp;收受者帳號：
                              <input type="text" name="P_RCLNO" size="15" class="font9" maxlength="14" value="<%=RS("P_RCLNO")%>">
                              &nbsp;&nbsp;&nbsp;&nbsp;收受者身分證/統編：
                              <input type="text" name="P_PID" size="12" class="font9" maxlength="10" onKeyUp="UCaseP_PID();" value="<%=RS("P_PID")%>">
                              &nbsp;&nbsp;&nbsp;&nbsp;
                            </td>
                            <td colspan="2"> </td>
                          </tr>
                          <tr id="line" style="display:block">
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">機構名稱：</td>
                            <td>
                            <%
                              If Session("comp_label")="1" Then
                                SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' And Dept_Id In ("&Session("all_dept_type")&") Order By Comp_Label,Dept_Id"
                              Else
                                SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id In ("&Session("all_dept_type")&") Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
                              End If
                              FName="Dept_Id"
                              Listfield="Comp_ShortName"
                              menusize="1"
                              BoundColumn=RS("Dept_Id")
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                  　        </td>
                            <td align="right">收據開立：</td>
                            <td>
                            <%
                              SQL="Select Invoice_Type=CodeDesc From CASECODE Where CodeType='InvoiceType' Order By Seq"
                              FName="Invoice_Type"
                              Listfield="Invoice_Type"
                              menusize="1"
                              BoundColumn=RS("Invoice_Type")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td align="right">收據抬頭：</td><!--20131112 Modify by GoodTV Tanya:修改欄位顯示長度-->
                            <td><input type="text" name="Invoice_Title" size="25" class="font9" maxlength="40" value="<%=Data_Minus(RS("Invoice_Title"))%>"></td>
                            <td align="right">入帳銀行：</td>
                            <td>   
                            <!--20140123 Modify by GoodTV Tanya:入帳銀行改為唯讀文字欄位-->
                            <input type="text" name="Accoun_Bank" size="10" class="font9t" readonly value="<%=RS("Accoun_Bank")%>">
                            <!--<%
                              SQL="Select Accoun_Bank=CodeDesc From CASECODE Where CodeType='Bank' Order By Seq"
                              FName="Accoun_Bank"
                              Listfield="Accoun_Bank"
                              menusize="1"
                              BoundColumn=RS("Accoun_Bank")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>-->
                            </td>
                          </tr>
                          <tr>
                            <td align="right">會計科目：</td>
                            <td align="left" colspan="3">
                            <%
                              SQL="Select Accounting_Title=CodeDesc From CASECODE Where CodeType='Accoun' Order By Seq"
                              FName="Accounting_Title"
                              Listfield="Accounting_Title"
                              menusize="1"
                              BoundColumn=RS("Accounting_Title")
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                            </td>
                            <td align="right">募款活動：</td>
                            <td align="left" colspan="3">
                            <%
                              SQL="Select Act_Id,Act_ShortName From ACT Order By Act_id Desc"
                              FName="Act_Id"
                              Listfield="Act_ShortName"
                              menusize="1"
                              BoundColumn=RS("Act_Id")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">授權狀態：</td>
                            <td>
                      	      <SELECT Name='Status' size='1' style='font-size: 9pt; font-family: 新細明體'>
                      		      <OPTION  value='授權中' <%If RS("Status")="授權中" Then Response.Write "selected" End If%> >授權中</OPTION>
                      		      <OPTION  value='停止' <%If RS("Status")="停止" Then Response.Write "selected" End If%> >停止</OPTION>
                      	      </SELECT>
                  　        </td>
                            <td align="right">停止日期：</td>
                            <td><%call Calendar("Break_Date",RS("Break_Date"))%></td>
                            <td align="right">停止原因：</td><!--20131112 Modify by GoodTV Tanya:修改欄位顯示長度-->
                            <td colspan="3"><input type="text" name="Break_Reason" size="25" class="font9" maxlength="40" value="<%=RS("Break_Reason")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">備註：</td>
                            <td align="left" colspan="7"><textarea rows="4" name="Comment" cols="45" class="font9"><%=RS("Comment")%></textarea></td>
                          </tr>
                          <!--#include file="../include/calendar2.asp"-->
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="center" colspan="8">
                               <input type="button" value=" 修 改 " name="save" class="cbutton" style="cursor:hand" onClick="Update_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value=" 刪 除 " name="save" class="cbutton" style="cursor:hand" onClick="Delete_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value=" 取 消 " name="cancel" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">
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
function Window_OnLoad(){
  document.form.Donate_Payment.focus();
}	
function Donate_Payment_OnChange(DonatePayment){
  if(DonatePayment.indexOf('信用卡')>-1&&(document.form.Donate_Payment.value.indexOf('扣款')>-1||document.form.Donate_Payment.value.indexOf('轉帳')>-1||DonatePayment.indexOf('授權書')>-1)){
    donatecard1.style.display='block';
    donatecard2.style.display='block';
    post.style.display='none';
	  bank.style.display='none';
    document.form.Post_Name.value='';
    document.form.Post_IDNo.value='';
    document.form.Post_SavingsNo.value='';
    document.form.Post_AccountNo.value='';
    line.style.display='block';
    //20130807 Modify by GoodTV Tanya:授權方式連動信用卡卡別
    if(DonatePayment.indexOf('聯信')>-1){    
			document.form.Card_Type.value='AE';
    }
    else if(DonatePayment.indexOf('聯信')==-1&&document.form.Card_Type.value=='AE'){
    	document.form.Card_Type.value='';
    }    
  }else if(DonatePayment.indexOf('帳戶')>-1&&(document.form.Donate_Payment.value.indexOf('扣款')>-1||document.form.Donate_Payment.value.indexOf('轉帳')>-1||DonatePayment.indexOf('授權書')>-1)){
    donatecard1.style.display='none';
    donatecard2.style.display='none';
	bank.style.display='none';
    document.form.Card_Bank.value='';
    document.form.Card_Type.value='';
    document.form.Account_No1.value='';
    document.form.Account_No2.value='';
    document.form.Account_No3.value='';
    document.form.Account_No4.value='';
    document.form.Valid_Month.value='';
    document.form.Valid_Year.value='';
    document.form.Card_Owner.value='';
    document.form.Same_Owner.checked=false;
    document.form.Owner_IDNo.value='';
    document.form.Authorize.value='';
    document.form.Relation.value='';
    post.style.display='block';
    line.style.display='block';    
  }else if(DonatePayment.indexOf('ACH轉帳授權書')>-1){
    donatecard1.style.display='none';
    donatecard2.style.display='none';
    post.style.display='none';
    document.form.Card_Bank.value='';
    document.form.Card_Type.value='';
    document.form.Account_No1.value='';
    document.form.Account_No2.value='';
    document.form.Account_No3.value='';
    document.form.Account_No4.value='';
    document.form.Valid_Month.value='';
    document.form.Valid_Year.value='';
    document.form.Card_Owner.value='';
    document.form.Same_Owner.checked=false;
    document.form.Owner_IDNo.value='';
    document.form.Authorize.value='';
    document.form.Relation.value='';
    document.form.Post_Name.value='';
    document.form.Post_IDNo.value='';
    document.form.Post_SavingsNo.value='';
    document.form.Post_AccountNo.value='';
    bank.style.display='block';
    line.style.display='block';    
  }
  //20140123 Add by GoodTV Tanya:連動入帳銀行
  if(DonatePayment.indexOf('郵局')>-1)
  	document.form.Accoun_Bank.value='郵局';  
  else	
		document.form.Accoun_Bank.value='臺灣銀行';	
}
function Account_No_Keyup(Account_No){
  if(Account_No==1&&document.form.Account_No1.value.length==4){
    document.form.Account_No2.focus();
  }else if(Account_No==2&&document.form.Account_No2.value.length==4){
    document.form.Account_No3.focus();
  }else if(Account_No==3&&document.form.Account_No3.value.length==4){
    document.form.Account_No4.focus();    
  }  
}
function Same_Owner_OnClick(){
  if(document.form.Same_Owner.checked){
    document.form.Card_Owner.value=document.form.DonorName.value;
    document.form.Owner_IDNo.value=document.form.DonorIDNo.value;
    document.form.Relation.value='本人';
  }
}
function Same_Post_OnClick(){
  if(document.form.Same_Post.checked){
    document.form.Post_Name.value=document.form.DonorName.value;
    document.form.Post_IDNo.value=document.form.DonorIDNo.value;
  }
}
function UCaseOwnerIDNO(){
  document.form.Owner_IDNo.value=document.form.Owner_IDNo.value.toUpperCase();
}
function UCasePostIDNO(){
  document.form.Post_IDNo.value=document.form.Post_IDNo.value.toUpperCase();
}
function UCaseP_PID(){
  document.form.P_PID.value=document.form.P_PID.value.toUpperCase();
}
function Update_OnClick(){
  <%call CheckStringJ("Donate_Payment","授權方式")%>
  <%call ChecklenJ("Donate_Payment",20,"授權方式")%>
  <%call CheckStringJ("Donate_Purpose","指定用途")%>
  <%call ChecklenJ("Donate_Purpose",20,"指定用途")%>
  <%call CheckStringJ("Donate_Amt","每次扣款")%>
  <%call CheckNumberJ("Donate_Amt","每次扣款")%>
  //20131105 Mark by GoodTV Tanya:目前用不到先隱藏「下次強制扣款」
  //if(document.form.Donate_FirstAmt.value==''){
  //  document.form.Donate_FirstAmt.value='0';
  //}  
  <%'call CheckNumberJ("Donate_FirstAmt","下次強制扣款")%>  
  <%call CheckStringJ("Donate_FromDate","授權起日")%>
  <%call CheckDateJ("Donate_FromDate","授權起日")%>
  <%call CheckStringJ("Donate_ToDate","授權迄日")%>
  <%call CheckDateJ("Donate_ToDate","授權迄日")%>
  var begindate=new Date(document.form.Donate_FromDate.value)
  var enddate=new Date(document.form.Donate_ToDate.value)
  var diffdate=(Date.parse(begindate.toString())-Date.parse(enddate.toString()))/(1000*60*60*24)
  if(parseInt(diffdate)>0){
    alert('授權起日不可大於迄日！');
    document.form.Donate_FromDate.focus();
    return;
  }
  <%call CheckStringJ("Donate_Period","轉帳週期")%>
  <%call CheckStringJ("Next_DonateDate","下次扣款日")%>
  <%call CheckDateJ("Next_DonateDate","下次扣款日")%>
  if(document.form.Donate_Payment.value.indexOf('信用卡')>-1&&(document.form.Donate_Payment.value.indexOf('扣款')>-1||document.form.Donate_Payment.value.indexOf('轉帳')>-1||document.form.Donate_Payment.value.indexOf('授權書')>-1)){
    document.form.Post_Name.value='';
    document.form.Post_IDNo.value='';
    document.form.Post_SavingsNo.value='';
    document.form.Post_AccountNo.value='';
    <%call ChecklenJ("Card_Bank",20,"銀行別")%>
    <%call ChecklenJ("Card_Type",10,"信用卡別")%>
    <%call CheckStringJ("Account_No1","信用卡號")%>
    <%call CheckNumberJ("Account_No1","信用卡號")%>
    <%call ChecklenJ("Account_No1",4,"信用卡號")%>
    <%call CheckStringJ("Account_No2","信用卡號")%>
    <%call CheckNumberJ("Account_No2","信用卡號")%>
    <%call ChecklenJ("Account_No2",4,"信用卡號")%>
    <%call CheckStringJ("Account_No3","信用卡號")%>
    <%call CheckNumberJ("Account_No3","信用卡號")%>
    <%call ChecklenJ("Account_No3",4,"信用卡號")%>
    <%call CheckStringJ("Account_No4","信用卡號")%>
    <%call CheckNumberJ("Account_No4","信用卡號")%>
    <%call ChecklenJ("Account_No4",4,"信用卡號")%>
    <%call CheckStringJ("Valid_Month","有效月")%>
    <%call CheckStringJ("Valid_Year","有效年")%>
    <%call ChecklenJ("Card_Owner",40,"持卡人")%>
    <%call ChecklenJ("Owner_IDNo",10,"持卡人身分證")%>
    <%call ChecklenJ("Relation",10,"與捐款人關係")%>
  }else if(document.form.Donate_Payment.value.indexOf('帳戶')>-1&&(document.form.Donate_Payment.value.indexOf('扣款')>-1||document.form.Donate_Payment.value.indexOf('轉帳')>-1||document.form.Donate_Payment.value.indexOf('授權書')>-1)){
    document.form.Card_Bank.value='';
    document.form.Card_Type.value='';
    document.form.Account_No1.value='';
    document.form.Account_No2.value='';
    document.form.Account_No3.value='';
    document.form.Account_No4.value='';
    document.form.Valid_Month.value='';
    document.form.Valid_Year.value='';
    document.form.Card_Owner.value='';
    document.form.Owner_IDNo.value='';
    document.form.Relation.value='';
    document.form.Authorize.value='';
    <%call ChecklenJ("Post_Name",40,"存簿戶名")%>
    <%call CheckStringJ("Post_IDNo","持有人身分證")%>
    <%call ChecklenJ("Post_IDNo",10,"持有人身分證")%>
    <%call CheckStringJ("Post_SavingsNo","存簿局號")%>
    <%call ChecklenJ("Post_SavingsNo",7,"存簿局號")%>
    <%call CheckNumberJ("Post_SavingsNo","存簿局號")%>
    <%call CheckStringJ("Post_AccountNo","存簿帳號")%>
    <%call ChecklenJ("Post_AccountNo",7,"存簿帳號")%>
    <%call CheckNumberJ("Post_AccountNo","存簿帳號")%>
  }else if(document.form.Donate_Payment.value.indexOf('ACH轉帳授權書')>-1){
    document.form.Card_Bank.value='';
    document.form.Card_Type.value='';
    document.form.Account_No1.value='';
    document.form.Account_No2.value='';
    document.form.Account_No3.value='';
    document.form.Account_No4.value='';
    document.form.Valid_Month.value='';
    document.form.Valid_Year.value='';
    document.form.Card_Owner.value='';
    document.form.Owner_IDNo.value='';
    document.form.Relation.value='';
    document.form.Authorize.value='';
    document.form.Post_Name.value='';
    document.form.Post_IDNo.value='';
    document.form.Post_SavingsNo.value='';
    document.form.Post_AccountNo.value='';
    <%call CheckStringJ("P_BANK","收受行代號")%>
    <%call ChecklenJ("P_BANK",7,"收受行代號")%>
    <%call CheckNumberJ("P_BANK","收受行代號")%>
    <%call CheckStringJ("P_RCLNO","收受者帳號")%>
    <%call ChecklenJ("P_RCLNO",14,"收受者帳號")%>
    <%call CheckNumberJ("P_RCLNO","收受者帳號")%>
    <%call CheckStringJ("P_PID","收受者身分證/統編")%>
    <%call ChecklenJ("P_PID",10,"收受者身分證/統編")%>
  }  
  <%call CheckStringJ("Dept_Id","機構名稱")%>
  if(document.form.Invoice_Type.value.indexOf('不')==-1){
    <%call CheckStringJ("Invoice_Title","收據抬頭")%>
  }
  <%call ChecklenJ("Invoice_Title",40,"收據抬頭")%>
  <%call ChecklenJ("Accoun_Bank",20,"入帳銀行")%>
  <%call ChecklenJ("Accounting_Title",30,"會計科目")%>
  if(document.form.Status.value=='停止'){
    <%call CheckStringJ("Break_Date","停止日期")%>
  }
  if(document.form.Break_Date.value!=''){
    <%call CheckDateJ("Break_Date","停止日期")%>
  }
  <%call ChecklenJ("Break_Reason",40,"原因")%>
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  if(document.form.ctype.value=="pledge_data"){
  //alert (document.form.Donor_Id.value) ;
  //if (true) return;
    location.href="pledge_data.asp?donor_id="+document.form.Donor_Id.value+"";
  }else{
    location.href="pledge.asp";
  }
}
--></script>