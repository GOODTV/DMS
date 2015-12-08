<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL1="Select * From PLEDGE Where Pledge_id='"&request("pledge_id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  Donor_Id=RS1("Donor_Id")
  RS1("Donate_Payment")=request("Donate_Payment")
  RS1("Donate_Purpose")=request("Donate_Purpose")
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
  If request("Donate_FirstAmt")<>"" Then
    RS1("Donate_FirstAmt")=request("Donate_FirstAmt")
  Else
    RS1("Donate_FirstAmt")="0"
  End If
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
  RS1("Card_Owner")=request("Card_Owner")
  RS1("Owner_IDNo")=request("Owner_IDNo")
  RS1("Relation")=request("Relation")
  RS1("Authorize")=request("Authorize")
  RS1("Post_Name")=request("Post_Name")
  RS1("Post_IDNo")=request("Post_IDNo")
  RS1("Post_SavingsNo")=request("Post_SavingsNo")
  RS1("Post_AccountNo")=request("Post_AccountNo")
  RS1("Dept_Id")=request("Dept_Id")
  RS1("Invoice_Type")=request("Invoice_Type")    
  RS1("Invoice_Title")=request("Invoice_Title")    
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
  RS1("Break_Reason")=""
  RS1("Comment")=request("Comment")
  RS1("LastUpdate_Date")=Date()
  RS1("LastUpdate_DateTime")=Now()
  RS1("LastUpdate_User")=session("user_name")
  RS1("LastUpdate_IP")=Request.ServerVariables("REMOTE_HOST")   
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
  If request("ctype")="member_donate_data" Then
    Response.Redirect "member_pledge_data.asp?donor_id="&request("donor_id")
  Else
    Response.Redirect "member_pledge.asp"
  End If  
End If

SQL="Select PLEDGE.*,Donor_Name=DONOR.Donor_Name,Title=DONOR.Title,Category=DONOR.Category,Donor_Type=DONOR.Donor_Type,InvoiceType=DONOR.Invoice_Type,IDNo=DONOR.IDNo,Last_DonateDate=DONOR.Last_DonateDate,Remark=DONOR.Remark, " & _
    "Address2=A.mValue+DONOR.ZipCode+B.mValue+DONOR.Address,Invoice_Address2=C.mValue+DONOR.Invoice_ZipCode+D.mValue+DONOR.Invoice_Address " & _
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
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>固定轉帳授權書資料修改</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <input type="hidden" name="Pledge_Id" value="<%=RS("Pledge_Id")%>">	
      <input type="hidden" name="Donor_Id" value="<%=RS("Donor_Id")%>">
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
                      <td class="table62-bg">　</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">固定轉帳授權書資料修改</td>
                      <td class="table63-bg">　</td>
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
                <td class="table62-bg">　</td>
                <td valign="top">
                  <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                    <tr>
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="2" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td align="right">捐款人：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Donor_Name" size="30" class="font9t" readonly value="<%=RS("Donor_Name")%>&nbsp;<%=RS("Title")%>">
                              &nbsp;類&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;別：
                              <input type="text" name="Category" size="10" class="font9t" readonly value="<%=RS("Category")%>">
                              &nbsp;身份別：
                              <input type="text" name="Donor_Type" size="40" class="font9t" readonly value="<%=RS("Donor_Type")%>">
                            </td>
                          </tr>
                          <tr>
                            <td align="right">收據地址：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Invoice_Address2" size="30" class="font9t" readonly value="<%=RS("Invoice_Address2")%>">
                              &nbsp;收據開立：
                              <input type="text" name="InvoiceType" size="10" class="font9t" readonly value="<%=RS("InvoiceType")%>">
                              &nbsp;身分證/統編：
                              <input type="text" name="IDNo" size="10" class="font9t" readonly value="<%=RS("IDNo")%>">
                              &nbsp;最近捐款日：
                              <input type="text" name="Last_DonateDate" size="10" class="font9t" readonly value="<%=RS("Last_DonateDate")%>">	
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
                              SQL="Select Donate_Purpose=CodeDesc From CASECODE Where CodeType='Purpose' And CodeKind='D' Order By Seq"
                              FName="Donate_Purpose"
                              Listfield="Donate_Purpose"
                              menusize="1"
                              BoundColumn=RS("Donate_Purpose")
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                            </td>
                            <td width="10%" align="right">每次扣款：</td>
                            <td width="14% align="left"><input type="text" name="Donate_Amt" size="12" class="font9" maxlength="7" value="<%=RS("Donate_Amt")%>"></td>
                            <td width="13%" align="right">下次強制扣款：</td>
                            <td width="13%"><input type="text" name="Donate_FirstAmt" size="12" class="font9" maxlength="7" value="<%=RS("Donate_FirstAmt")%>"></td>
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
                            <%If RS("Transfer_Date")<>"" Then%>
                            <td align="right">下次扣款日：</td>
                            <%Else%>
                            <td align="right">首次扣款日：</td>
                            <%End if%>
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
                              End If
                            %>
                              &nbsp;信用卡號:
                              <input type="text" name="Account_No1" size="3" class="font9" maxlength="4" OnKeyup="Account_No_Keyup('1');" value="<%=Account_No1%>">
                              -	
                              <input type="text" name="Account_No2" size="3" class="font9" maxlength="4" OnKeyup="Account_No_Keyup('2');" value="<%=Account_No2%>">
                              -	
                              <input type="text" name="Account_No3" size="3" class="font9" maxlength="4" OnKeyup="Account_No_Keyup('3');" value="<%=Account_No3%>">
                              -	
                              <input type="text" name="Account_No4" size="3" class="font9" maxlength="4" value="<%=Account_No4%>">			
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
                              For I= Year(Date())-10 To Year(Date())+10
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
                            <td align="left"><input type="text" name="Card_Owner" size="12" class="font9" maxlength="40" value="<%=RS("Card_Owner")%>"></td>
                            <td colspan="6">
                            	<input type="checkbox" name="Same_Owner" value="Y" OnClick="Same_Owner_OnClick()">同捐款人
                              &nbsp;持卡人身分證：
                              <input type="text" name="Owner_IDNo" size="9" class="font9" maxlength="10" onKeyUp="UCaseOwnerIDNO();" value="<%=RS("Owner_IDNo")%>">
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
                            <td><input type="text" name="Post_Name" size="12" class="font9" maxlength="40" value="<%=RS("Post_Name")%>"></td>
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
                          <tr id="line" style="display:none">
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
                            <td align="right">收據抬頭：</td>
                            <td><input type="text" name="Invoice_Title" size="12" class="font9" maxlength="40" value="<%=RS("Invoice_Title")%>"></td>
                            <td align="right">入帳銀行：</td>
                            <td>
                            <%
                              SQL="Select Accoun_Bank=CodeDesc From CASECODE Where CodeType='Bank' Order By Seq"
                              FName="Accoun_Bank"
                              Listfield="Accoun_Bank"
                              menusize="1"
                              BoundColumn=RS("Accoun_Bank")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
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
                            <td align="right">停止原因：</td>
                            <td colspan="3"><input type="text" name="Break_Reason" size="12" class="font9" maxlength="40" value="<%=RS("Break_Reason")%>"></td>
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
                <td class="table63-bg">　</td>
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
function Donate_Payment_OnChange(DonatePayment){
  if(DonatePayment.indexOf('信用卡')>-1&&(document.form.Donate_Payment.value.indexOf('扣款')>-1||document.form.Donate_Payment.value.indexOf('轉帳')>-1||DonatePayment.indexOf('授權書')>-1)){
    donatecard1.style.display='block';
    donatecard2.style.display='block';
    post.style.display='none';
    document.form.Post_Name.value='';
    document.form.Post_IDNo.value='';
    document.form.Post_SavingsNo.value='';
    document.form.Post_AccountNo.value='';
    line.style.display='block';
  }else if(DonatePayment.indexOf('帳戶')>-1&&(document.form.Donate_Payment.value.indexOf('扣款')>-1||document.form.Donate_Payment.value.indexOf('轉帳')>-1||DonatePayment.indexOf('授權書')>-1)){
    donatecard1.style.display='none';
    donatecard2.style.display='none';
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
  }
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
function Update_OnClick(){
  <%call CheckStringJ("Donate_Payment","授權方式")%>
  <%call ChecklenJ("Donate_Payment",20,"授權方式")%>
  <%call CheckStringJ("Donate_Purpose","指定用途")%>
  <%call ChecklenJ("Donate_Purpose",20,"指定用途")%>
  <%call CheckStringJ("Donate_Amt","每次扣款")%>
  <%call CheckNumberJ("Donate_Amt","每次扣款")%>
  if(document.form.Donate_FirstAmt.value==''){
    document.form.Donate_FirstAmt.value='0';
  }
  <%call CheckNumberJ("Donate_FirstAmt","下次強制扣款")%>
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
  if(document.form.ctype.value=='member_pledge_data'){
    location.href='member_pledge_data.asp?donor_id='+document.form.Donor_Id.value+'';
  }else{
    location.href='member_pledge.asp';
  }
}
--></script>