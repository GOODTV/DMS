<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  Check_Close=Get_Close("1",Cstr(request("Dept_Id")),Cstr(request("Donate_Date")),Cstr(Session("user_id")))
  If Check_Close Then
    CodeKind="M"
    SQL1="Select CodeKind From CASECODE Where CodeType='Purpose' And CodeDesc='"&request("Donate_Purpose")&"' "
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    CodeKind=RS1("CodeKind")
    RS1.Close
    Set RS1=Nothing
        
    SQL1="Select * From DONATE Where Donate_Id='"&request("Donate_Id")&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,3
    RS1("Donate_Date")=request("Donate_Date")
    If request("Donate_Year")<>"" Then
      RS1("Donate_Year")=request("Donate_Year")
    Else
      RS1("Donate_Year")=null
    End If    
    RS1("Donate_Payment")=request("Donate_Payment")
    RS1("Donate_Purpose")=request("Donate_Purpose")
    RS1("Donate_Purpose_Type")=CodeKind
    RS1("Donate_Type")=request("Donate_Type")
    If request("Donate_Amt")<>"" Then
      RS1("Donate_Amt")=request("Donate_Amt")
      RS1("Donate_Fee")=request("Donate_Fee")
      RS1("Donate_Accou")=request("Donate_Accou")
      RS1("Donate_Amt"&CodeKind&"")=request("Donate_Amt")
      RS1("Donate_Fee"&CodeKind&"")=request("Donate_Fee")
      RS1("Donate_Accou"&CodeKind&"")=request("Donate_Accou")
    Else
      RS1("Donate_Amt")="0"
      RS1("Donate_Accou")="0"
      RS1("Donate_Fee")="0"
      RS1("Donate_Amt"&CodeKind&"")="0"
      RS1("Donate_Accou"&CodeKind&"")="0"
      RS1("Donate_Fee"&CodeKind&"")="0"
    End If
    If CodeKind="D" Then
      RS1("Donate_AmtM")="0"
      RS1("Donate_FeeM")="0"
      RS1("Donate_AccouM")="0"
      RS1("Donate_AmtA")="0"
      RS1("Donate_FeeA")="0"
      RS1("Donate_AccouA")="0"
      RS1("Donate_AmtS")="0"
      RS1("Donate_RateS")="0"
      RS1("Donate_FeeS")="0"
      RS1("Donate_AccouS")="0"
    ElseIf CodeKind="M" Then
      RS1("Donate_AmtD")="0"
      RS1("Donate_FeeD")="0"
      RS1("Donate_AccouD")="0"
      RS1("Donate_AmtA")="0"
      RS1("Donate_FeeA")="0"
      RS1("Donate_AccouA")="0"
      RS1("Donate_AmtS")="0"
      RS1("Donate_RateS")="0"
      RS1("Donate_FeeS")="0"
      RS1("Donate_AccouS")="0"
    ElseIf CodeKind="A" Then
      RS1("Donate_AmtD")="0"
      RS1("Donate_FeeD")="0"
      RS1("Donate_AccouD")="0"
      RS1("Donate_AmtM")="0"
      RS1("Donate_FeeM")="0"
      RS1("Donate_AccouM")="0"
      RS1("Donate_AmtS")="0"
      RS1("Donate_RateS")="0"
      RS1("Donate_FeeS")="0"
      RS1("Donate_AccouS")="0"
    ElseIf CodeKind="S" Then
      RS1("Donate_AmtD")="0"
      RS1("Donate_FeeD")="0"
      RS1("Donate_AccouD")="0"
      RS1("Donate_AmtM")="0"
      RS1("Donate_FeeM")="0"
      RS1("Donate_AccouM")="0"
      RS1("Donate_AmtA")="0"
      RS1("Donate_FeeA")="0"
      RS1("Donate_AccouA")="0"
      RS1("Donate_RateS")="0"
    End If
    RS1("Donate_Forign")=request("Donate_Forign")
    RS1("Card_Bank")=request("Card_Bank")
    RS1("Card_Type")=request("Card_Type")
    RS1("Account_No")=request("Account_No1")&request("Account_No2")&request("Account_No3")&request("Account_No4")
    RS1("Valid_Date")=request("Valid_Month")&request("Valid_Year")
    RS1("Card_Owner")=request("Card_Owner")
    RS1("Owner_IDNo")=request("Owner_IDNo")
    RS1("Relation")=request("Relation")
    RS1("Authorize")=request("Authorize")
    RS1("Check_No")=request("Check_No")
    If request("Check_ExpireDate")<>"" Then
      RS1("Check_ExpireDate")=request("Check_ExpireDate")
    Else
      RS1("Check_ExpireDate")=null
    End If
    RS1("Post_Name")=request("Post_Name")
    RS1("Post_IDNo")=request("Post_IDNo")
    RS1("Post_SavingsNo")=request("Post_SavingsNo")
    RS1("Post_AccountNo")=request("Post_AccountNo")
    'RS1("Dept_Id")=request("Dept_Id")
    RS1("Invoice_Title")=request("Invoice_Title")
    If Trim(request("Invoice_No"))<>"" Then RS1("Invoice_No")=Trim(request("Invoice_No"))
    If request("Request_Date")<>"" Then
      RS1("Request_Date")=request("Request_Date")
    Else
      RS1("Request_Date")=null
    End If
    RS1("Accoun_Bank")=request("Accoun_Bank")
    If request("Accoun_Date")<>"" Then
      RS1("Accoun_Date")=request("Accoun_Date")
    Else
      RS1("Accoun_Date")=null
    End If
    RS1("Invoice_Type")=request("Invoice_Type")
    If request("InvoceSend_Date")<>"" Then
      RS1("InvoceSend_Date")=request("InvoceSend_Date")
    Else
      RS1("InvoceSend_Date")=null
    End If
    RS1("Accounting_Title")=request("Accounting_Title")
    RS1("Donation_NumberNo")=request("Donation_NumberNo")
    RS1("Donation_SubPoenaNo")=request("Donation_SubPoenaNo")
    If request("Act_id")<>"" Then
      RS1("Act_id")=request("Act_id")
    Else
      RS1("Act_id")=null
    End If
    RS1("Comment")=request("Comment")
    RS1("Invoice_PrintComment")=request("Invoice_PrintComment")
    If request("Issue_Type")<>"" Then
      RS1("Issue_Type")=request("Issue_Type")
      RS1("Issue_Type_Keep")=request("Issue_Type")
    Else
      RS1("Issue_Type")=""
      RS1("Issue_Type_Keep")=""
    End If
    RS1("LastUpdate_Date")=Date()
    RS1("LastUpdate_DateTime")=Now()
    RS1("LastUpdate_User")=session("user_name")
    RS1("LastUpdate_IP")=Request.ServerVariables("REMOTE_HOST")
    RS1.Update
    RS1.Close
    Set RS1=Nothing

    '修改捐款人捐款紀錄
    call Declare_DonorId (request("donor_id"))

    session("errnumber")=1
    session("msg")="捐款資料修改成功 ！"
    If request("ctype")="member_donate_data" Then
      Response.Redirect "member_donate_data.asp?donor_id="&request("donor_id")
    Else
      Response.Redirect "member_donate_detail.asp?donate_id="&request("donate_id")
    End If
  Else
    session("errnumber")=1
    session("msg")="您輸入的繳費日期『 "&Cstr(request("Donate_Date"))&"』 已關帳無法修改 ！"
  End If
End If

If request("action")="delete" Then
  Check_Close=Get_Close("1",Cstr(request("Dept_Id")),Cstr(request("Donate_Date")),Cstr(Session("user_id")))
  If Check_Close Then
    SQL="Delete From DONATE Where Donate_Id='"&request("Donate_Id")&"'"
    Set RS=Conn.Execute(SQL)
    
    '修改捐款人捐款紀錄
    call Declare_DonorId (request("donor_id"))
    
    session("errnumber")=1
    session("msg")="捐款資料刪除成功 ！"
    If request("ctype")="member_donate_data" Then
      Response.Redirect "member_donate_data.asp?donor_id="&request("donor_id")
    Else
      Response.Redirect "member_donate.asp"
    End If
  Else
    session("errnumber")=1
    session("msg")="您輸入的繳費日期『 "&Cstr(request("Donate_Date"))&"』 已關帳無法刪除 ！"
  End If  
End If

SQL="Select DONATE.*,Comp_ShortName,Donor_Name=DONOR.Donor_Name,Title=DONOR.Title,Category=DONOR.Category,Donor_Type=DONOR.Donor_Type,InvoiceType=DONOR.Invoice_Type,IDNo=DONOR.IDNo,Member_Type,Member_Act_ShortName, " & _
    "Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+DONOR.Address Else A.mValue+DONOR.ZipCode+DONOR.Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When C.mValue<>D.mValue Then C.mValue+DONOR.Invoice_ZipCode+D.mValue+Invoice_Address Else C.mValue+DONOR.Invoice_ZipCode+DONOR.Invoice_Address End End) " & _
    "From DONATE " & _
    "Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id " & _
    "Join Dept On DONATE.Dept_Id=DEPT.Dept_Id " & _
    "Left Join MEMBER_ACT On DONATE.Member_Act_Id=MEMBER_ACT.Member_Act_Id " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _
    "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
    "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
    "Where Donate_Id='"&request("donate_id")&"' "
Call QuerySQL(SQL,RS)

Donate_Purpose_Member="常年會費"
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>會員繳費資料修改</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="Donate_Purpose_Member" value="<%=Donate_Purpose_Member%>">		
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <input type="hidden" name="Dept_Id" value="<%=RS("Dept_Id")%>">	
      <input type="hidden" name="Donate_Id" value="<%=RS("Donate_Id")%>">	
      <input type="hidden" name="Donor_Id" value="<%=RS("donor_id")%>">
      <input type="hidden" name="DonorName" value="<%=RS("Donor_Name")%>">
      <input type="hidden" name="DonorIDNo" value="<%=RS("IDNo")%>">
      <input type="hidden" name="InvoiceNo" value="<%=RS("Invoice_Pre")&RS("Invoice_No")%>">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">會員繳費資料修改</td>
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
                        <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td align="right">會員姓名：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Donor_Name" size="30" class="font9t" readonly value="<%=RS("Donor_Name")%>&nbsp;<%=RS("Title")%>">
                              &nbsp;類&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;別：
                              <input type="text" name="Category" size="10" class="font9t" readonly value="<%=RS("Category")%>">
                              &nbsp;會員別：
                              <input type="text" name="Member_Type" size="40" class="font9t" readonly value="<%=RS("Member_Type")%>">
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
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td width="10%" align="right">繳費日期：</td>
                            <td width="15% align="left"><%call Calendar("Donate_Date",RS("Donate_Date"))%></td>
                            <td width="10%" align="right">繳費方式：</td>
                            <td width="17%">
                            <%
                              SQL1="Select Donate_Payment=CodeDesc From CASECODE Where CodeType='Payment' Order By Seq"
                              Set RS1 = Server.CreateObject("ADODB.RecordSet")
                              RS1.Open SQL1,Conn,1,1
                              Response.Write "<SELECT Name='Donate_Payment' size='1' style='font-size: 9pt; font-family: 新細明體' OnChange=""Donate_Payment_OnChange(this.value)"">"
                              Response.Write "<OPTION>" & " " & "</OPTION>"
                              While Not RS1.EOF
                                If Cstr(RS1("Donate_Payment"))=Cstr(RS("Donate_Payment")) Then
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
                            <td width="10%" align="right">繳費用途：</td>
                            <td width="14%">
                             <%
                              SQL1="Select Donate_Purpose=CodeDesc From CASECODE Where CodeType='Purpose' And CodeKind='M' Order By Seq"
                              Set RS1 = Server.CreateObject("ADODB.RecordSet")
                              RS1.Open SQL1,Conn,1,1
                              Response.Write "<SELECT Name='Donate_Purpose' size='1' style='font-size: 9pt; font-family: 新細明體' OnChange=""Donate_Purpose_OnChange(this.value)"">"
                              Response.Write "<OPTION>" & " " & "</OPTION>"
                              While Not RS1.EOF
                                If RS("Donate_Purpose")=RS1("Donate_Purpose") Then
                                  Response.Write "<OPTION value='"&RS1("Donate_Purpose")&"' selected >"&RS1("Donate_Purpose")&"</OPTION>"
                                Else
                                  Response.Write "<OPTION value='"&RS1("Donate_Purpose")&"'>"&RS1("Donate_Purpose")&"</OPTION>"
                                End If
                                RS1.MoveNext
                              Wend
                              Response.Write "</SELECT>"
                              RS1.Close
                              Set RS1=Nothing
                            %>
                            </td>
                            <td width="10%" align="right">繳費年度：</td>
                            <%If Cstr(RS("Donate_Purpose"))=Donate_Purpose_Member Then%>
                            <td width="14%"><input type="text" name="Donate_Year" size="5" class="font9" maxlength="4" value="<%=RS("Donate_Year")%>">(西元)</td>
                            <%Else%>
                            <td width="14%"><input type="text" name="Donate_Year" size="5" class="font9t" maxlength="4" value="">(西元)</td>
                            <%End If%>
                          </tr>
                          <tr>
                            <td align="right">繳費金額：</td>
                            <td><input type="text" name="Donate_Amt" size="12" class="font9" maxlength="10" value="<%=RS("Donate_Amt")%>" onKeyUp="javascript:SumAccou();"></td>
                            <td align="right">手續費：</td>
                            <td><input type="text" name="Donate_Fee" size="12" class="font9" maxlength="10" value="<%=RS("Donate_Fee")%>" onKeyUp="javascript:SumAccou();"></td>
                            <td align="right">實收金額：</td>
                            <td><input type="text" name="Donate_Accou" size="12" class="font9" maxlength="10" value="<%=RS("Donate_Accou")%>"></td>
                            <td align="right">外幣：</td>
                            <td><input type="text" name="Donate_Forign" size="12" class="font9" maxlength="20" value="<%=RS("Donate_Forign")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">收據開立：</td>
                            <td colspan="7">
                            <%
                              SQL="Select Invoice_Type=CodeDesc From CASECODE Where CodeType='InvoiceType' Order By Seq"
                              FName="Invoice_Type"
                              Listfield="Invoice_Type"
                              menusize="1"
                              BoundColumn=RS("Invoice_Type")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
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
                                If Len(RS("Account_No"))>=4 Then Account_No1=Cint(Left(RS("Account_No"),4))
                                If Len(RS("Account_No"))>=8 Then Account_No2=Cint(Mid(RS("Account_No"),5,4))
                                If Len(RS("Account_No"))>=12 Then Account_No3=Cint(Mid(RS("Account_No"),9,4))
                                If Len(RS("Account_No"))>=16 Then Account_No4=Cint(Right(RS("Account_No"),4))
                              End If
                            %>
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;信用卡號:
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
                              Response.Write "<tr id=""donatecard2"" style=""display:block"">"
                            Else
                              Response.Write "<tr id=""donatecard2"" style=""display:none"">"
                            End If
                          %>
                            <td align="right">持卡人：</td>
                            <td align="left"><input type="text" name="Card_Owner" size="12" class="font9" maxlength="40" value="<%=RS("Card_Owner")%>"></td>
                            <td colspan="4">
                            	<input type="checkbox" name="Same_Owner" value="Y" OnClick="Same_Owner_OnClick()">同捐款人
                              &nbsp;&nbsp;&nbsp;持卡人身分證：
                              <input type="text" name="Owner_IDNo" size="10" class="font9" maxlength="10" onKeyUp="UCaseOwnerIDNO();" value="<%=RS("Owner_IDNo")%>">
                              &nbsp;&nbsp;&nbsp;與會員關係：
                              <input type="text" name="Relation" size="4" class="font9" maxlength="10" value="<%=RS("Relation")%>">
                            </td> 
                            <td align="right">授權碼：</td>
                            <td align="left"><input type="text" name="Authorize" size="12" class="font9" maxlength="3" value="<%=RS("Authorize")%>"></td>
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
                            <td><input type="text" name="Check_No" size="12" class="font9" maxlength="20" value="<%=RS("Check_No")%>"></td>
                            <td align="right">到期日：</td>
                            <td colspan="5"><%call Calendar("Check_ExpireDate",RS("Check_ExpireDate"))%></td>
                          </tr>
                          <%
                            If Instr(Cstr(RS("Donate_Payment")),"帳戶")>0 And (Instr(Cstr(RS("Donate_Payment")),"轉帳")>0 Or Instr(Cstr(RS("Donate_Payment")),"授權書")>0) Then
                              Check_Line=True
                              Response.Write "<tr id=""post"" style=""display:block"">"
                            Else
                              Response.Write "<tr id=""post"" style=""display:none"">"
                            End If
                          %>
                            <td align="right">存簿戶名：</td>
                            <td><input type="text" name="Post_Name" size="12" class="font9" maxlength="40" value="<%=RS("Post_Name")%>"></td>
                            <td colspan="6">
                            	&nbsp;&nbsp;持有人身分證：
                              <input type="text" name="Post_IDNo" size="9" class="font9" maxlength="10" onKeyUp="UCasePostIDNO();"  value="<%=RS("Post_IDNo")%>">
                            	<input type="checkbox" name="Same_Post" value="Y" OnClick="Same_Post_OnClick()">同捐款人
                            	&nbsp;&nbsp;&nbsp;存簿局號：
                            	<input type="text" name="Post_SavingsNo" size="12" class="font9" maxlength="10" value="<%=RS("Post_SavingsNo")%>">
                              &nbsp;&nbsp;&nbsp;存簿帳號：
                              <input type="text" name="Post_AccountNo" size="12" class="font9" maxlength="10" value="<%=RS("Post_AccountNo")%>">
                            </td>
                            <td colspan="2"> </td>
                          </tr>
                          <%If Check_Line Then%>
                          <tr id="line" style="display:block">
                          <%Else%>
                          <tr id="line" style="display:none">
                          <%End If%>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">機構名稱：</td>
                            <td><input type="text" name="Comp_ShortName" size="12" class="font9t" readonly value="<%=RS("Comp_ShortName")%>"></td>
                            <td align="right">收據抬頭：</td>
                            <td colspan="2"><input type="text" name="Invoice_Title" size="30" class="font9" maxlength="80" value="<%=RS("Invoice_Title")%>"></td>
                            <td><input type="checkbox" name="Issue_Type" value="M" OnClick="Issue_Type_OnClick()" <%If RS("Issue_Type")="M" Then Response.Write "checked" End If%> >手開收據</td>
                            <td align="right">收據編號：</td>
                            <td><input type="text" name="Invoice_No" size="12" <%If RS("Issue_Type")="M" Then Response.Write "class=""font9""" Else Response.Write "class=""font9t"" readonly " End If%> maxlength="20" value="<%=RS("Invoice_Pre")&RS("Invoice_No")%>"></td>
                          </tr> 
                          <tr>
                            <td align="right">請款日期：</td>
                            <td align="left"><%call Calendar("Request_Date",RS("Request_Date"))%></td>
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
                            <td align="right">沖帳日期：</td>
                            <td><%call Calendar("Accoun_Date",RS("Accoun_Date"))%></td>
                            <td align="right">繳費類別：</td>
                            <td>
                              <Select size="1" name="Donate_Type" class="font9">
                                <option value="單次捐款" <%If RS("Donate_Type")="單次捐款" Then Response.Write "selected" End If%> >單次捐款</option>
                                <option value="長期捐款" <%If RS("Donate_Type")="長期捐款" Then Response.Write "selected" End If%> >長期捐款</option>
                              </Select>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">傳票號碼：</td>
                            <td align="left"><input type="text" name="Donation_NumberNo" size="14" class="font9" maxlength="20" onKeyUp="javascript:UCaseNumberNo();" value="<%=RS("Donation_NumberNo")%>"></td>
                            <td align="left" colspan="2">&nbsp;&nbsp;劃撥&nbsp;/&nbsp;匯款單號：<input type="text" name="Donation_SubPoenaNo" size="14" class="font9" maxlength="20" onKeyUp="javascript:UCaseSubPoenaNo();" value="<%=RS("Donation_SubPoenaNo")%>"></td>
                            <td align="right">會計科目：</td>
                            <td align="left">
                            <%
                              SQL="Select Accounting_Title=CodeDesc From CASECODE Where CodeType='Accoun' And CodeKind In ('M','A') Order By Seq"
                              FName="Accounting_Title"
                              Listfield="Accounting_Title"
                              menusize="1"
                              BoundColumn=RS("Accounting_Title")
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                            </td>
                            <td align="right">收據寄送：</td>
                            <td align="left"><%call Calendar("InvoceSend_Date",RS("InvoceSend_Date"))%></td>
                          </tr>
                          <tr>
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
                            <td align="right">活動報名：</td>
                            <td align="left" colspan="3"><input type="text" name="Member_Act_ShortName" size="42" class="font9t" readonly value="<%=RS("Member_Act_ShortName")%>"></td>
                          </tr>
                          
                          <tr>
                            <td align="right">繳費備註：</td>
                            <td align="left" colspan="3"><textarea rows="3" name="Comment" cols="45" class="font9"><%=RS("Comment")%></textarea></td>
                            <td align="right">收據備註：<br />(列印用)&nbsp;&nbsp;&nbsp;&nbsp;</td>
                            <td align="left" colspan="3"><textarea rows="3" name="Invoice_PrintComment" cols="40" class="font9"><%=RS("Invoice_PrintComment")%></textarea></td>
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
    checkno.style.display='none';
    document.form.Check_No.value='';
    document.form.Check_ExpireDate.value='';
    post.style.display='none';
    document.form.Check_No.value='';
    document.form.Post_Name.value='';
    document.form.Post_IDNo.value='';
    document.form.Post_SavingsNo.value='';
    document.form.Post_AccountNo.value='';
    line.style.display='block';
  }else if(DonatePayment.indexOf('支票')>-1){
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
    checkno.style.display='block';
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
    checkno.style.display='none';
    document.form.Check_No.value='';
    document.form.Check_ExpireDate.value='';
    post.style.display='block';
    line.style.display='block';
  }else{
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
    document.form.Relation.value='';
    document.form.Authorize.value='';
    checkno.style.display='none';
    document.form.Check_No.value='';
    document.form.Check_ExpireDate.value='';
    post.style.display='none';
    document.form.Post_Name.value='';
    document.form.Post_IDNo.value='';
    document.form.Post_SavingsNo.value='';
    document.form.Post_AccountNo.value='';
    line.style.display='none';
  }
}
function Donate_Purpose_OnChange(Donate_Purpose){
  if(Donate_Purpose==document.form.Donate_Purpose_Member.value){
    document.form.Donate_Year.style.backgroundColor='#ffffff';
    document.form.Donate_Year.readOnly=false;
    document.form.Donate_Year.value=new Date().getYear();
    document.form.Donate_Year.focus();
  }else{
    document.form.Donate_Year.style.backgroundColor='#ffffcc';
    document.form.Donate_Year.readOnly=true;
    document.form.Donate_Year.value='';
  }
}
function SumAccou(){
  if(isNaN(Number(document.form.Donate_Amt.value))||isNaN(Number(document.form.Donate_Fee.value))){
    document.form.Donate_Accou.value='0';
  }else{
    document.form.Donate_Accou.value=Number(document.form.Donate_Amt.value)-Number(document.form.Donate_Fee.value);
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
function Issue_Type_OnClick(){
  if(document.form.Issue_Type.checked){
    document.form.Invoice_No.style.backgroundColor='#ffffff';
    document.form.Invoice_No.readOnly=false;
    document.form.Invoice_No.value='';
    document.form.Invoice_No.focus();
  }else{
    document.form.Invoice_No.style.backgroundColor='#ffffcc';
    document.form.Invoice_No.readOnly=true;
    document.form.Invoice_No.value=document.form.InvoiceNo.value;
  }
}
function UCaseNumberNo(){
  document.form.Donation_NumberNo.value=document.form.Donation_NumberNo.value.toUpperCase();
}
function UCaseSubPoenaNo(){
  document.form.Donation_SubPoenaNo.value=document.form.Donation_SubPoenaNo.value.toUpperCase();
}
function Update_OnClick(){
  <%call CheckStringJ("Donate_Date","繳費日期")%>
  <%call CheckDateJ("Donate_Date","繳費日期")%>
  <%call CheckStringJ("Donate_Payment","繳費方式")%>
  <%call ChecklenJ("Donate_Payment",20,"繳費方式")%>
  <%call CheckStringJ("Donate_Purpose","繳費用途")%>
  <%call ChecklenJ("Donate_Purpose",20,"繳費用途")%>
  if(document.form.Donate_Purpose.value==document.form.Donate_Purpose_Member.value){
    <%call CheckStringJ("Donate_Year","繳費年度")%>
    <%call CheckNumberJ("Donate_Year","繳費年度")%>
    if(document.form.Donate_Year.value.length!=4){
      alert('繳費年度必須為西元年');
      document.form.Donate_Year.focus();
      return;
    }
  }  
  <%call CheckStringJ("Invoice_Type","收據開立")%>
  <%call ChecklenJ("Invoice_Type",20,"收據開立")%>  
  if(document.form.Donate_Payment.value.indexOf('信用卡')>-1&&(document.form.Donate_Payment.value.indexOf('扣款')>-1||document.form.Donate_Payment.value.indexOf('轉帳')>-1||document.form.Donate_Payment.value.indexOf('授權書')>-1)){
    document.form.Check_No.value='';
    document.form.Check_ExpireDate.value='';
    document.form.Post_Name.value='';
    document.form.Post_IDNo.value='';
    document.form.Post_SavingsNo.value='';
    document.form.Post_AccountNo.value='';
    <%call CheckStringJ("Donate_Amt","繳費金額")%>
    <%call CheckNumberJ("Donate_Amt","繳費金額")%>
    <%call CheckNumberJ("Donate_Fee","手續費")%>
    if(document.form.Donate_Fee.value==''){
      document.form.Donate_Fee.value='0';
    }
    <%call CheckNumberJ("Donate_Accou","實收金額")%>
    <%call ChecklenJ("Donate_Forign",20,"外幣")%>
  }else if(document.form.Donate_Payment.value.indexOf('支票')>-1){
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
    <%call CheckStringJ("Donate_Amt","繳費金額")%>
    <%call CheckNumberJ("Donate_Amt","繳費金額")%>
    <%call CheckNumberJ("Donate_Fee","手續費")%>
    if(document.form.Donate_Fee.value==''){
      document.form.Donate_Fee.value='0';
    }
    <%call CheckNumberJ("Donate_Accou","實收金額")%>
    <%call ChecklenJ("Donate_Forign",20,"外幣")%>
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
    document.form.Check_No.value='';
    document.form.Check_ExpireDate.value='';
    <%call CheckStringJ("Donate_Amt","繳費金額")%>
    <%call CheckNumberJ("Donate_Amt","繳費金額")%>
    <%call CheckNumberJ("Donate_Fee","手續費")%>
    if(document.form.Donate_Fee.value==''){
      document.form.Donate_Fee.value='0';
    }
    <%call CheckNumberJ("Donate_Accou","實收金額")%>
    <%call ChecklenJ("Donate_Forign",20,"外幣")%>
  }else{
    <%call CheckStringJ("Donate_Amt","繳費金額")%>
    <%call CheckNumberJ("Donate_Amt","繳費金額")%>
    <%call CheckNumberJ("Donate_Fee","手續費")%>
    if(document.form.Donate_Fee.value==''){
      document.form.Donate_Fee.value='0';
    }
    <%call CheckNumberJ("Donate_Accou","實收金額")%>
    <%call ChecklenJ("Donate_Forign",20,"外幣")%>
  }
  if(Number(document.form.Donate_Accou.value)>Number(document.form.Donate_Amt.value)){
    alert('實收金額不可大於繳費金額');
    document.form.Donate_Accou.focus();
    return;
  }  
  <%call ChecklenJ("Card_Bank",20,"銀行別")%>
  <%call ChecklenJ("Card_Type",10,"信用卡別")%>
  <%call CheckNumberJ("Account_No1","信用卡號")%>
  <%call ChecklenJ("Account_No1",4,"信用卡號")%>
  <%call CheckNumberJ("Account_No2","信用卡號")%>
  <%call ChecklenJ("Account_No2",4,"信用卡號")%>
  <%call CheckNumberJ("Account_No3","信用卡號")%>
  <%call ChecklenJ("Account_No3",4,"信用卡號")%>
  <%call CheckNumberJ("Account_No4","信用卡號")%>
  <%call ChecklenJ("Account_No4",4,"信用卡號")%>
  <%call ChecklenJ("Card_Owner",40,"持卡人")%>
  <%call ChecklenJ("Owner_IDNo",10,"持卡人身分證")%>
  <%call ChecklenJ("Relation",10,"與捐款人關係")%>
  <%call ChecklenJ("Authorize",10,"授權碼")%>
  <%call ChecklenJ("Check_No",20,"支票號碼")%>
  <%call ChecklenJ("Post_Name",40,"存簿戶名")%>
  <%call ChecklenJ("Post_IDNo",10,"持有人身分證")%>
  <%call ChecklenJ("Post_SavingsNo",10,"存簿局號")%>
  <%call CheckNumberJ("Post_SavingsNo","存簿局號")%>
  <%call ChecklenJ("Post_AccountNo",10,"存簿帳號")%>
  <%call CheckNumberJ("Post_AccountNo","存簿帳號")%>
  <%call CheckStringJ("Dept_Id","機構名稱")%>
  <%call ChecklenJ("Invoice_Title",100,"收據抬頭")%>
  if(document.form.Issue_Type.checked){
    <%call CheckStringJ("Invoice_No","收據編號")%>
    <%call ChecklenJ("Invoice_No",20,"收據編號")%>
  }else{
    document.form.Invoice_No.value='';
  }
  if(document.form.Request_Date.value!=''){
    <%call CheckDateJ("Request_Date","請款日期")%>
  }  
  <%call ChecklenJ("Accoun_Bank",20,"入帳銀行")%>
  if(document.form.Accoun_Date.value!=''){
    <%call CheckDateJ("Accoun_Date","沖帳日期")%>
  }  
  <%call ChecklenJ("Donate_Type",10,"繳費類別")%>
  <%call ChecklenJ("Accounting_Title",30,"會計科目")%>
  <%call ChecklenJ("Donation_NumberNo",20,"傳票號碼")%>
  <%call ChecklenJ("Donation_SubPoenaNo",20,"劃撥 / 匯款單號")%>
  if(document.form.InvoceSend_Date.value!=''){
    <%call CheckDateJ("InvoceSend_Date","收據寄送日期")%>
  }  
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  location.href='member_donate_detail.asp?ctype='+document.form.ctype.value+'&donate_id='+document.form.Donate_Id.value+'';
}
--></script>