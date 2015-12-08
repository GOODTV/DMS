<% Response.Buffer = True %>
<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Session("Transfer_AspName")=request("Transfer_AspName")
  Session("Donate_Date")=request("Donate_Date")  
  Server.Transfer "creditcard/"&request("Transfer_AspName")
End If

If request("action")="input" Then
  
  If request("pledge_id") <>"" then
	Session("pledge_id")=request("pledge_id")
	Session("Donate_Date")=request("Donate_Date")
	Session("Transfer_Date")=request("Transfer_Date")
	Response.Redirect "pledge_inputcheck.asp"
  End If
  
End If


SQL1="Select * From DEPT Where Dept_Id='"&session("dept_id")&"'"
Call QuerySQL(SQL1,RS1)
Transfer_Date=RS1("Transfer_Date")
LastTransfer_Date=RS1("LastTransfer_Date")
RS1.Close
Set RS1=Nothing
If request("Donate_Date")="" Then
  If LastTransfer_Date<>"" Then
    DonateDate=DateAdd("M",1,LastTransfer_Date)  
  Else
    DonateDate=Year(Date())&"/"&Month(Date())&"/"&Transfer_Date
  End If
Else
  DonateDate=request("Donate_Date")
End If

LastTransfer_Check="0"
If LastTransfer_Date<>"" Then
  If Cstr(Year(LastTransfer_Date))=Cstr(Year(Date())) And Cstr(Month(LastTransfer_Date))=Cstr(Month(Date())) Then LastTransfer_Check="1"
End If
If request("action")<>"" Then 
'response.write request("action")
	'撈出Temp Table:PLEDGE_SEND_RETURN裡的Pledge_Id串成陣列，丟給hidden欄位：PledgeId
	SQL_ChkPledge="Select Pledge_Id From dbo.PLEDGE_SEND_RETURN Where Return_Status='Y'"
	Set RS_Chk = Server.CreateObject("ADODB.RecordSet")
	RS_Chk.Open SQL_ChkPledge,Conn,1,1
	While Not RS_Chk.EOF
		Arry_PledgeId = RS_Chk("Pledge_Id")&","&Arry_PledgeId
		RS_Chk.MoveNext
	Wend
End If
if request("PledgeId_check") <> "" then
	PledgeId_check=request("PledgeId_check")
else
	PledgeId_check="N"	
end if
'20140120 Add by GoodTV Tanya:紀錄是否已點選過「查詢」按鈕
if request("Query_Flag") <> "" then
	Query_Flag=request("Query_Flag")
else
	Query_Flag="N"	
end if		
'Response.write "query_flag1="&Query_Flag
%>
<%Prog_Id="pledge_transfer"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
	  <input type="hidden" name="PledgeId" value="<%=Arry_PledgeId%>">
	  <input type="hidden" id="PledgeId_check" name="PledgeId_check" value="<%=PledgeId_check%>">
      <input type="hidden" name="Transfer_Date" value="<%=Transfer_Date%>">	
      <input type="hidden" name="LastTransfer_Check" value="<%=LastTransfer_Check%>">	
      <input type="hidden" id="Query_Flag" name="Query_Flag" value="<%=Query_Flag%>">		  
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="830" border="0" cellspacing="0" cellpadding="0"  align="left" style="background-color:#EEEEE3">
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
	          <table width="830" border="0" cellspacing="0" cellpadding="0" align="left">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">&nbsp;</td>
                <td valign="top">
                  <table width="100%" border="0" cellpadding="1" cellspacing="1">
                    <tr>
                      <td class="td02-c" width="65" align="right">機構：</td>
                      <td class="td02-c" width="95">
                      <%
                        If Session("comp_label")="1" Then
                          SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' And Dept_Id In ("&Session("all_dept_type")&") Order By Comp_Label,Dept_Id"
                        Else
                          SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id In ("&Session("all_dept_type")&") Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
                        End If
                        FName="Dept_Id"
                        Listfield="Comp_ShortName"
                        menusize="1"
                        BoundColumn=request("Dept_Id")
                        call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                      %>
                      </td>
                      <td class="td02-c" width="65" align="right">捐款人：</td>
                      <td class="td02-c" width="70"><input type="text" name="Donor_Name" size="8" class="font9" value="<%=request("Donor_Name")%>"></td>
                      <td class="td02-c" width="65" align="right">TXT格式：</td>
                      <td class="td02-c" width="110">
                      <%
                        SQL="Select Transfer_AspName,Transfer_BankName From DONATE_TRANSFER Where Transfer_StopFlg='N' Order By Transfer_Seq,Ser_No Desc"
                        FName="Transfer_AspName"
                        Listfield="Transfer_BankName"
                        menusize="1"
                        BoundColumn=request("Transfer_AspName")
                        call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                      %>
                      </td>
                      <td class="td02-c" width="325" align="right">
                        <input type="button" value=" 查詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">
												<input type="button" value="1.TXT匯出" name="export" class="addbutton" style="cursor:hand" onClick="Export_OnClick()">
												<input type="button" value="2.授權轉捐款" name="input" class="addbutton" style="cursor:hand" onClick="Input_OnClick()">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">扣款日期：</td>
                      <td class="td02-c"><%call Calendar("Donate_Date",DonateDate)%></td>
                      <td class="td02-c" align="right">轉帳週期：</td>
                      <td class="td02-c">
                        <SELECT Name='Donate_Period' size='1' style='font-size: 9pt; font-family: 新細明體'>
                      	  <OPTION  value=''> </OPTION>
                      		<OPTION  value='單筆' <%If request("Donate_Period")="單筆" Then Response.Write "selected" End If%> >單筆</OPTION>
                      		<OPTION  value='月繳' <%If request("Donate_Period")="月繳" Then Response.Write "selected" End If%> >月繳</OPTION>
                      		<OPTION  value='季繳' <%If request("Donate_Period")="季繳" Then Response.Write "selected" End If%> >季繳</OPTION>
                      		<OPTION  value='半年繳' <%If request("Donate_Period")="半年繳" Then Response.Write "selected" End If%> >半年繳</OPTION>
                      		<OPTION  value='年繳' <%If request("Donate_Period")="年繳" Then Response.Write "selected" End If%> >年繳</OPTION>
                      	</SELECT>
                      </td>
                      <td class="td02-c" align="right">授權方式：</td>
                      <td class="td02-c">
                      <%
                        SQL="Select Donate_Payment=CodeDesc From CASECODE Where CodeType='Payment' And CodeDesc Like '%授權書%' Order By Seq"
                        FName="Donate_Payment"
                        Listfield="Donate_Payment"
                        menusize="1"
                        BoundColumn=request("Donate_Payment")
                        call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                      %>
                      </td>
                      <td class="td02-c">
					  						<input type="button" value="2.1匯入台銀回覆檔" name="import" class="delbutton" style="cursor:hand" onClick="Import_OnClick()">
					  
					  						<input type="button" value="2.2查詢台銀回覆檔" name="import_check" class="delbutton" style="cursor:hand" onClick="Import_Check()" >
					  
					  					</td>	
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <%
                    	'20131111Modify by GoodTV Tanya:page_load不帶全部資料
                    	If request("action")<>"" Then                    		                    	
	                      SQL_P="Select Donate_No=Count(*),Donate_Amt=Sum(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End) From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Status='授權中' And Post_SavingsNo<>'' And Post_AccountNo<>'' "
	                      SQL_C="Select Donate_No=Count(*),Donate_Amt=Sum(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End) From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Status='授權中' And Account_No<>'' And (Convert(int,RIGHT(Valid_Date,2))>"&Right(Year(DonateDate),2)&" Or (Convert(int,RIGHT(Valid_Date,2))="&Right(Year(DonateDate),2)&" And Convert(int,Left(Valid_Date,2))>="&Month(DonateDate)&")) "
	                      SQL1="Select Pledge_Id,授權編號=Pledge_Id,捐款人=DONOR.Donor_Name,授權方式=Donate_Payment,扣款金額=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),授權起日=CONVERT(VarChar,Donate_FromDate,111),授權迄日=CONVERT(VarChar,Donate_ToDate,111),轉帳週期=Donate_Period,下次扣款日=Next_DonateDate,有效月年=(Case When Valid_Date<>'' Then Left(Valid_Date,2)+'/'+Right(Valid_Date,2) Else '' End),最後扣款日=Transfer_Date " & _
	                           "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Status='授權中' And ((Post_SavingsNo<>'' And Post_AccountNo<>'') Or Account_No<>'' Or P_RCLNO<>'') "
	                      If request("Dept_Id")<>"" Then 
	                        WhereSQL=WhereSQL & "And PLEDGE.Dept_Id = '"&request("Dept_Id")&"' "
	                      Else
	                        If request("action")="" Then WhereSQL=WhereSQL & "And PLEDGE.Dept_Id In ("&Session("all_dept_type")&") "
	                      End If
	                      If request("Donor_Name")<>"" Then WhereSQL=WhereSQL & "And (DONOR.Donor_Name Like '%"&request("Donor_Name")&"%' Or DONOR.NickName Like '%"&request("Donor_Name")&"%' Or DONOR.Contactor Like '%"&request("Donor_Name")&"%' Or DONOR.Invoice_Title Like '%"&request("Donor_Name")&"%') "
	                      If request("Donate_Date")<>"" Then 
	                        WhereSQL=WhereSQL & "And Donate_FromDate <= '"&request("Donate_Date")&"' And Donate_ToDate>= '"&request("Donate_Date")&"' And ((Year(Next_DonateDate)<'"&Year(request("Donate_Date"))&"') Or (Year(Next_DonateDate)='"&Year(request("Donate_Date"))&"' And Month(Next_DonateDate)<='"&Month(request("Donate_Date"))&"')) "
	                        WhereSQL_PC=WhereSQL_PC & "And Donate_FromDate <= '"&request("Donate_Date")&"' And Donate_ToDate>= '"&request("Donate_Date")&"' And ((Year(Next_DonateDate)<'"&Year(request("Donate_Date"))&"') Or (Year(Next_DonateDate)='"&Year(request("Donate_Date"))&"' And Month(Next_DonateDate)<='"&Month(request("Donate_Date"))&"')) "
	                      Else
	                        WhereSQL=WhereSQL & "And Donate_FromDate <= '"&DonateDate&"' And Donate_ToDate>= '"&DonateDate&"' And ((Year(Next_DonateDate)<'"&Year(DonateDate)&"') Or (Year(Next_DonateDate)='"&Year(DonateDate)&"' And Month(Next_DonateDate)<='"&Month(DonateDate)&"')) "
	                        WhereSQL_PC=WhereSQL_PC & "And Donate_FromDate <= '"&DonateDate&"' And Donate_ToDate>= '"&DonateDate&"' And ((Year(Next_DonateDate)<'"&Year(DonateDate)&"') Or (Year(Next_DonateDate)='"&Year(DonateDate)&"' And Month(Next_DonateDate)<='"&Month(DonateDate)&"')) "
	                      End If
	                      If request("Donate_Payment")<>"" Then WhereSQL=WhereSQL & "And Donate_Payment = '"&request("Donate_Payment")&"' "
	                      If request("Donate_Period")<>"" Then WhereSQL=WhereSQL & "And Donate_Period = '"&request("Donate_Period")&"' "
                      
	                      '帳戶轉帳授權筆數/金額
	                      Donate_No_P=0
	                      Donate_Amt_P=0
	                      SQL_P=SQL_P&WhereSQL&WhereSQL_PC
	                      Set RS_P = Server.CreateObject("ADODB.RecordSet")
	                      RS_P.Open SQL_P,Conn,1,1
	                      If Not RS_P.EOF Then
	                        If RS_P("Donate_No")<>"" Then Donate_No_P=Cdbl(RS_P("Donate_No"))
	                        If RS_P("Donate_Amt")<>"" Then Donate_Amt_P=Cdbl(RS_P("Donate_Amt"))
	                      End If
	                      RS_P.Close
	                      Set RS_P=Nothing
                      
	                      '信用卡扣款授權筆數/金額
	                      Donate_No_C=0
	                      Donate_Amt_C=0  
	                      SQL_C=SQL_C&WhereSQL&WhereSQL_PC
	                      Set RS_C = Server.CreateObject("ADODB.RecordSet")
	                      RS_C.Open SQL_C,Conn,1,1
	                      If Not RS_C.EOF Then
	                        If RS_C("Donate_No")<>"" Then Donate_No_C=Cdbl(RS_C("Donate_No"))
	                        If RS_C("Donate_Amt")<>"" Then Donate_Amt_C=Cdbl(RS_C("Donate_Amt"))
	                      End If
	                      RS_C.Close
	                      Set RS_C=Nothing
	                  	End If
                    %>
                    <tr>
                      <td class="td02-c" colspan="7">
                         <table width="100%" border="0" cellpadding="3" cellspacing="3">
                           <tr>
                    	       <td width="100">信用卡有效月年：</td>
                    	       <td width="15" height="15" bgcolor='#66FF99'> </td>
                    	       <td width="30">2個月</td>
                    	       <td width="15" height="15" bgcolor='#FFFF99'> </td>
                    	       <td width="30">1個月</td>
                    	       <td width="15" height="15" bgcolor='#FFCCFF'> </td>
                    	       <td width="25">本月</td>
                    	       <td width="15" height="15" bgcolor='#FF6666'> </td>
                    	       <td width="25">逾期</td>
                    	       <td align="right">本月應執行&nbsp;帳戶轉帳<%=FormatNumber(Donate_No_P,0)%>筆&nbsp;&nbsp;金額<%=FormatNumber(Donate_Amt_P,0)%>元&nbsp;╱&nbsp;信用卡扣款<%=FormatNumber(Donate_No_C,0)%>筆&nbsp;&nbsp;金額<%=FormatNumber(Donate_Amt_C,0)%>&nbsp;元</td>
                           </tr>
                         </table>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" colspan="7" width="100%" height="400" valign="top">
                      <%		           
                      	'20131111Modify by GoodTV Tanya:page_load不帶全部資料
                      	If request("action")="" Then 		
                      		Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
												  Response.Write "  <tr>"
												  Response.Write "    <td width='100%' align='center' style='color:#ff0000'>** 請先輸入查詢條件 **</td>"	  
												  Response.Write "  </tr>"
												  Response.Write "</table>"  
                      	Else
                      		'20140103Modify by GoodTV Tanya:排序僅by授權編號
	                        SQL1=SQL1&WhereSQL&"Order By Pledge_Id "
	                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
	                        RS1.Open SQL1,Conn,1,1
							'response.write SQL1
	                        FieldsCount = RS1.Fields.Count-1
	                        Dim I, J
	                        Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;background-color: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
	                        Response.Write "<tr>"
	                        If RS1.Recordcount>0 Then
	                          Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'><input type='checkbox' name='pledge_id' id='pledge_id_0' value='0' checked OnClick='PledgeId_OnClick()'></span></font></td>"
	                        Else
	                          Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>選擇</span></font></td>"
	                        End If
	                        For I = 1 To FieldsCount
		                        Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
	                        Next
	                        Response.Write "</tr>"
	                        Row=0
	                        While Not RS1.EOF
	                          Row=Row+1
	                          StrChecked="checked"
	                          bgColor="#FFFFFF"
	                          If RS1(9)<>"" Then                            
	                            Valid_Year_N=Right(Year(Date()),2)
	                            Valid_Month_N=Month(Date())
	                            If Len(Valid_Month_N)=1 Then Valid_Month_N="0"&Valid_Month_N
	                            Valid_Date_N=Cint(Cstr(Valid_Year_N&Valid_Month_N))
	                                                        
	                            Valid_Year_2=Right(Year(DateAdd("M",2,"20"&Valid_Year_N&"/"&Valid_Month_N&"/1")),2)
	                            Valid_Month_2=Month(DateAdd("M",2,"20"&Valid_Year_N&"/"&Valid_Month_N&"/1"))
	                            If Len(Valid_Month_2)=1 Then Valid_Month_2="0"&Valid_Month_2
	                            Valid_Date_2=Cint(Cstr(Valid_Year_2&Valid_Month_2))
	                            
	                            Valid_Year_1=Right(Year(DateAdd("M",1,"20"&Valid_Year_N&"/"&Valid_Month_N&"/1")),2)
	                            Valid_Month_1=Month(DateAdd("M",1,"20"&Valid_Year_N&"/"&Valid_Month_N&"/1"))
	                            If Len(Valid_Month_1)=1 Then Valid_Month_1="0"&Valid_Month_1
	                            Valid_Date_1=Cint(Cstr(Valid_Year_1&Valid_Month_1))
	                            
	                            Valid_Year=Right(RS1(9),2)
	                            Valid_Month=Left(RS1(9),2)
	                            Valid_Date=Cint(Cstr(Valid_Year&Valid_Month))
	                            
	                            If Valid_Date=Valid_Date_2 Then
	                              bgColor="#66FF99"
	                            ElseIf Valid_Date=Valid_Date_1 Then
	                              bgColor="#FFFF99"
	                            ElseIf Valid_Date=Valid_Date_N Then
	                              bgColor="#FFCCFF"
	                            ElseIf Valid_Date<Valid_Date_N Then
	                              StrChecked=""
	                              bgColor="#FF6666"
	                            End If
	                          End If
	                          
	                          'If request("Donate_Date")<>"" Then
	                          '  If Cint(Year(RS1(8)))<>Cint(Year(request("Donate_Date"))) Or Cint(Month(RS1(8)))<>Cint(Month(request("Donate_Date"))) Then StrChecked=""
	                          'ElseIf DonateDate<>"" Then
	                          '  If Cint(Year(RS1(8)))<>Cint(Year(DonateDate)) Or Cint(Month(RS1(8)))<>Cint(Month(DonateDate)) Then StrChecked=""
	                          'End If
	                          
	                          Response.Write "<tr "&showhand&" bgcolor='"&bgColor&"' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor="""&bgColor&"""'>"
	                          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'><input type='checkbox' name='pledge_id' id='pledge_id_"&Row&"' value='"&RS1(0)&"' "&StrChecked&" >" & "</span></td>"
	                          For J = 1 To FieldsCount
	                            If J=4 Then
	                              Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(J),0) & "</span></td>"
	                            Else
	                              Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
	                            End If
	                          Next
	                          Response.Write "</tr>"
	                          Response.Flush
	                          Response.Clear
	                          RS1.MoveNext
	                        Wend  
	                        Response.Write "</table>"
	                        
							if Row>0 then
								Response.Write "<input type='hidden' id='flag' name='flag' value='True'>"
							else
								Response.Write "<input type='hidden' id='flag' name='flag' value='False'>"
							End if													
													
	                        RS1.Close
	                        Set RS1=Nothing
	                    	End If
                      %>
					  
                      <input type="hidden" name="Total_Row" value="<%=Row%>">		
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
           <td valign="top">
           		<table width="332" border="0" cellspacing="0" cellpadding="0" align="left">
           			<tr height="35px">
    							<td style="background-color:#EEEEE3">
    								<b><font size="4px" color="brown">&nbsp;::信用卡批次授權(一般)說明-台銀::</font></b>
    							</td>
    						</tr>
    						<tr>
    							<td style="background-color:#EEEEE3">
    								<font size="3px" color="darkmagenta">
    									&nbsp;&nbsp;&nbsp;1.1查詢預進行批次授權資料</font><br/>
    									<font size="3px">
    										&nbsp;&nbsp;&nbsp;&nbsp;☆扣款日期：預計執行授權之日期<br/>
    										&nbsp;&nbsp;&nbsp;&nbsp;☆授權方式：信用卡授權書(一般)<br/>
    										&nbsp;&nbsp;&nbsp;&nbsp;☆點選按鈕</font><font size="3px" color="blue">【查詢】</font>
    									<p/> 								      								
    							</td>
    						</tr>
    						<tr>
    							<td style="background-color:#EEEEE3">
    								<font size="3px" color="darkmagenta">
    									&nbsp;&nbsp;&nbsp;1.2匯出批次授權文字檔</font><br>
    									<font size="3px">
    										&nbsp;&nbsp;&nbsp;&nbsp;☆勾選預計進行批次授權資料<br/>
    										&nbsp;&nbsp;&nbsp;&nbsp;☆TXT格式：台灣銀行<br/>
    										&nbsp;&nbsp;&nbsp;&nbsp;☆點選按鈕</font><font size="3px" color="green">【1.TXT匯出】</font>
    								<p/>
    							</td>
    						</tr>
    						<tr>
    							<td style="background-color:#EEEEE3">
    								<font size="3px" color="darkmagenta">
    								&nbsp;&nbsp;&nbsp;1.3將匯出文字檔壓縮加密後，寄送給台銀</font><br>
    								<font size="3px" color="red">
    								&nbsp;&nbsp;&nbsp;※文字檔附檔名不得為TXT</font>
    								<p />
    							</td>
    						</tr>
    						<tr align="center">
    							<td style="background-color:#EEEEE3" >
    								<font size="2px" color="chocolate">
    								=============待收到台銀回覆檔後=============<p /></font>
    							</td>
    					  </tr>
    						<tr>
    							<td style="background-color:#EEEEE3">
    									<font size="3px" color="darkmagenta">
    									&nbsp;&nbsp;&nbsp;2.1重新查詢同(1.1)條件資料</font><br/>
    									<font size="3px">
    										&nbsp;&nbsp;&nbsp;&nbsp;☆輸入同(1.1)查詢條件資料<br/>
    										&nbsp;&nbsp;&nbsp;&nbsp;☆點選按鈕</font><font size="3px" color="blue">【查詢】</font>
    									<p/>
    							</td>
    						</tr>
    						<tr>
    							<td style="background-color:#EEEEE3">
    								<font size="3px" color="darkmagenta">
    								&nbsp;&nbsp;&nbsp;2.2匯入「回覆檔」</font><br/>
    								<font size="3px">
    									&nbsp;&nbsp;&nbsp;&nbsp;☆點選按鈕</font><font size="3px" color="red">【2.1匯入台銀回覆檔】</font><font size="3px">進行匯入</font>    									
    								<p />
    							</td>
    						</tr>
    						<tr>
    							<td style="background-color:#EEEEE3">
    								<font size="3px" color="darkmagenta">
    								&nbsp;&nbsp;&nbsp;2.3查詢已上傳「回覆檔」之資料</font><br/>
    								<font size="3px">
    									&nbsp;&nbsp;&nbsp;&nbsp;☆點選按鈕</font><font size="3px" color="red">【2.2查詢台銀回覆檔】</font><br />
    									<font size="3px">&nbsp;&nbsp;&nbsp;&nbsp;☆系統自動勾選已授權成功資料</font>
    								<p />
    							</td>
    						</tr>
    						<tr>
    							<td style="background-color:#EEEEE3">
    							<font size="3px" color="darkmagenta">
    								&nbsp;&nbsp;&nbsp;2.4進行授權轉捐款作業</font><br/>
    								<font size="3px">
    									&nbsp;&nbsp;&nbsp;&nbsp;☆點選按鈕</font><font size="3px" color="green">【2.授權轉捐款】</font>
    							</td>
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
  document.form.Donor_Name.focus();
  var Query_Flag = document.getElementById('Query_Flag').value;
  //alert(Query_Flag);
  if (Query_Flag=='N')
  	document.getElementById('import_check').disabled=true;
  else{
  	var val= document.getElementById('flag').value;
  	
  	if (val=='True')
			document.getElementById('import_check').disabled=false;
    else
			document.getElementById('import_check').disabled=true;
  }
   
}

function Export_OnClick(){
  <%call CheckStringJ("Donate_Date","扣款日期")%>
  <%call CheckDateJ("Donate_Date","扣款日期")%>
  <%call CheckStringJ("Transfer_AspName","TXT格式")%>
  if(confirm('您是否確定要將TXT授權資料匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
}
function Input_OnClick(){
  <%call CheckStringJ("Donate_Date","扣款日期")%>
  <%call CheckDateJ("Donate_Date","扣款日期")%>
  var choose = "";
  var j = 0;
  for(var i=1;i<=Number(document.form.Total_Row.value);i++){
	if(document.form.pledge_id[i].checked == true){
		var temp = document.form.pledge_id[i].value;
		if (j==0)
		{
			choose = temp;
			j++;
		}
		else
		{
			choose = choose + "," + temp;
		}
	}
  }
  var count = choose.split(",")
  //alert(count.length);//return;
  if(document.form.Total_Row.value>'0'){
    if(document.form.LastTransfer_Check.value=='1'){
      if(confirm('本月份已執行過授權資料轉入捐款資料，\n\n您是否確定要『再次』執行？')){
        document.form.action.value='input';
		/*location.href= "pledge_inputcheck.asp?pledge_id="+choose;
		var win = window.open("pledge_inputcheck.asp?pledge_id="+choose,"blank","status=no,toolbar=no,location=no,menubar=no,scrollbars=yes,width=650,height=400");
        var timer = setInterval(function() {   
			if(win.closed) {  
				clearInterval(timer);  
				alert('closed');  
				document.form.submit();
			}  
		}, 100);*/
		document.form.submit(); 
      }  
    }else{
       if(confirm('您是否確定要將授權資料轉入捐款資料？')){
        document.form.action.value='input';
		document.form.submit(); 
      }  
    }
  }else{
    alert('查無相關授權資料無法轉入！');
    return;
  }  
}
function Query_OnClick(){
  document.form.action.value="query";
  document.getElementById("Query_Flag").value="Y";
  document.form.submit();
}	
function Import_OnClick(){
  var winvar = window.open("pledge_import.asp","NAME","status=no,toolbar=no,location=no,menubar=no,width=450,height=200");
  var timer = setInterval(function() {   
			if(winvar.closed) {  
				clearInterval(timer);  
				//alert('closed');
				document.form.action.value="import";
				document.form.submit();
			}  
		}, 100);
}
function Import_Check(){
	//alert("PledgeId:"+document.form.PledgeId.value);
	//alert("PledgeId_check:"+document.form.PledgeId_check.value);
	if(document.form.PledgeId.value=="" && document.form.PledgeId_check.value=="N"){
		alert("請先匯入台銀回覆檔！");
	}else{
		
		for(var i=0;i<=Number(document.form.Total_Row.value);i++){
		  document.form.pledge_id[i].checked=false;
		}

		var str = document.form.PledgeId.value;
		var numberList = str.split(",")
		var flag = true;
		//alert (numberList.length-1);
		for(var i = 0; i<numberList.length;i++)
		{
			for(var j=1;j<=Number(document.form.Total_Row.value);j++)
			{
				if(document.form.pledge_id[j].value == numberList[i])
				{
					document.form.pledge_id[j].checked=true;
					flag = false;
				}
			}
		}

		if(flag == true)
		{
			alert("授權清單與回覆檔資料不符！");
		}else{
			alert("查詢完成，授權成功資料已自動勾選！");
		}
		
	}
}
function PledgeId_OnClick(){
  if(document.form.pledge_id[0].checked){
    for(var i=1;i<=Number(document.form.Total_Row.value);i++){
      document.form.pledge_id[i].checked=true;
    }
  }else{
    for(var i=1;i<=Number(document.form.Total_Row.value);i++){
      document.form.pledge_id[i].checked=false;
    }
  }
}
--></script>	