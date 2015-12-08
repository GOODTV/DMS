<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function DonateGridList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  If Not RS1.EOF Then 
    FieldsCount = RS1.Fields.Count-1
    totRec=RS1.Recordcount
    If totRec>0 Then 
      RS1.PageSize=PageSize
      If nowPage="" or nowPage=0 Then
        nowPage=1
      ElseIf cint(nowPage) > RS1.PageCount Then 
        nowPage=RS1.PageCount 
      End If
      session("nowPage")=nowPage
      RS1.AbsolutePage=nowPage
      totPage=RS1.PageCount
      SQL=server.URLEncode(SQL)
    End If
    Response.Write "<Form action='' id=form1 name=form1>" 
    Response.Write "<table border=0 cellspacing='0' cellpadding='1' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='30%'></td>"
    Response.Write "<td width='55%'><span style='font-size: 9pt; font-family: 新細明體'>第" & nowPage & "/" & totPage & "頁&nbsp;&nbsp;</span>"
    Response.Write "<span style='font-size: 9pt; font-family: 新細明體'>共" & FormatNumber(totRec,0) & "筆&nbsp;&nbsp;</span>"
    If cint(nowPage) <>1 Then
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage-1)&"&SQL="&SQL&"'>上一頁</a>" 
    End If
    If cint(nowPage)<>RS1.PageCount and cint(nowPage)<RS1.PageCount Then 
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage+1)&"&SQL="&SQL&"'>下一頁</a>" 
    End If
    Response.Write " |&nbsp;<span style='font-size: 9pt; font-family: 新細明體'> 跳至第<select name=GoPage size='1' style='font-size: 9pt; font-family: 新細明體' onchange='GoPage_OnChange(this.value)'>"
    For iPage=1 to totPage
      If iPage=cint(nowPage) Then
        strSelected = "selected"
      Else
	      strSelected = ""
      End If
      Response.Write "<option value='"&iPage&"'" & strSelected & ">" & iPage & "</option>"
    Next
    Response.Write "</select>頁</span></td>" 
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>回查詢</a></span></td>"
    End If   
    Response.Write "</tr></table>"
    Dim I
    Dim J
    Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    For J = 1 To FieldsCount
      Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J).Name & "</span></font></td>"
    Next
    Response.Write "</tr>"
    I = 1
    While Not RS1.EOF And I <= RS1.PageSize
      If Hlink<>"" Then
        Response.Write "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>"
        showhand="style='cursor:hand'"
      End If
      Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
      For J = 1 To FieldsCount
        If J=3 Then
          If RS1(J)<>"" Then
            Response.Write "<td bgcolor='#FFFFFF' align='right'><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(J),0) & "</span></td>"
          Else
            Response.Write "<td bgcolor='#FFFFFF' align='right'><span style='font-size: 9pt; font-family: 新細明體'>0</span></td>"
          End If
        Else
          Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
	      End If
      Next
      I = I + 1
      RS1.MoveNext
      Response.Write "</tr>"
      Response.Write "</a>"
    Wend
    Response.Write "</table>"
    Response.Write "</form>"
  Else
    Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 沒有符合條件的資料 **</td>"
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    End If
    Response.Write "</table>"
  End If
  RS1.Close
  Set RS1=Nothing
End Function
%>
<%Prog_Id="donate_name_list"%>
<!--#include file="../include/head.asp"-->
<body>
  <p><div align="center"><center>
    <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
      <tr>
        <td>
          <table width="800" border="0" cellspacing="0" cellpadding="0" align="center" style="background-color:#EEEEE3">
            <tr>
              <td width="5%"></td>
              <td width="95%">
  		          <table width="40%" border="0" cellspacing="0" cellpadding="0">
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
	        <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
              <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
              <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
            </tr>
            <tr>
              <td class="table62-bg">&nbsp;</td>
              <td valign="top">
                <table width="100%"  border="0" cellspacing="0" style="border-collapse: collapse">
                  <tr>
                    <td class="td02-c" width="100%">
                      <table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                        <tr>
                          <td>
                  	      <%
                            If request("SQL")="" Then
                              '搜尋條件一
                              Donate_Where=""
                              If Request("Dept_Id")<>"" Then
                                Donate_Where=Donate_Where&"And Donate.Dept_Id = '"&Request("Dept_Id")&"' "
                              Else
                                Donate_Where=Donate_Where&"And Donate.Dept_Id In ("&Session("all_dept_type")&") "
                              End If
                              If Request("Donate_Date_Begin")<>"" Then Donate_Where=Donate_Where&"And Donate_Date>='"&Request("Donate_Date_Begin")&"' "
                              If Request("Donate_Date_End")<>"" Then Donate_Where=Donate_Where&"And Donate_Date<='"&Request("Donate_Date_End")&"' "
                              If Request("Donate_Amt_Begin")<>"" Then Donate_Where=Donate_Where&"And CONVERT(numeric,Donate_Amt)>='"&Request("Donate_Amt_Begin")&"' " 
                              If Request("Donate_Amt_End")<>"" Then Donate_Where=Donate_Where&"And CONVERT(numeric,Donate_Amt)<='"&Request("Donate_Amt_End")&"' " 
                              If Request("Act_Id")<>"" Then Donate_Where=Donate_Where&"And DONATE.Act_Id='"&Request("Act_Id")&"' "
                              If Request("Invoice_No_Begin")<>"" Then Donate_Where=Donate_Where&"And Invoice_No>='"&Request("Invoice_No_Begin")&"' "
                              If Request("Invoice_No_End")<>"" Then Donate_Where=Donate_Where&"And Invoice_No<='"&Request("Invoice_No_End")&"' "
                              If Request("Donate_Payment")<>"" Then
                                donatepayment=""
                                Ary_Donate_Payment=Split(Request("Donate_Payment"),",")
                                If UBound(Ary_Donate_Payment)=0 Then
                                  donatepayment=donatepayment&"('"&Trim(Ary_Donate_Payment(0))&"')"
                                Else
                                  For I = 0 To UBound(Ary_Donate_Payment)
                                    If I=0 Then
                                      donatepayment=donatepayment&"('"&Trim(Ary_Donate_Payment(I))&"',"
                                    ElseIf I=UBound(Ary_Donate_Payment) Then 
                                      donatepayment=donatepayment&"'"&Trim(Ary_Donate_Payment(I))&"')" 
                                    Else
                                      donatepayment=donatepayment&"'"&Trim(Ary_Donate_Payment(I))&"',"
                                    End If
                                  Next
                                End If
                                If donatepayment<>"" Then Donate_Where=Donate_Where&"And Donate_Payment In "&donatepayment&" "
                              End If
                              If Request("Donate_Purpose")<>"" Then
                                donatepurpose=""
                                Ary_Donate_Purpose=Split(Request("Donate_Purpose"),",")
                                If UBound(Ary_Donate_Purpose)=0 Then
                                  donatepurpose=donatepurpose&"('"&Trim(Ary_Donate_Purpose(0))&"')"
                                Else
                                  For I = 0 To UBound(Ary_Donate_Purpose)
                                    If I=0 Then
                                      donatepurpose=donatepurpose&"('"&Trim(Ary_Donate_Purpose(I))&"',"
                                    ElseIf I=UBound(Ary_Donate_Purpose) Then 
                                      donatepurpose=donatepurpose&"'"&Trim(Ary_Donate_Purpose(I))&"')" 
                                    Else
                                      donatepurpose=donatepurpose&"'"&Trim(Ary_Donate_Purpose(I))&"',"
                                    End If
                                  Next
                                End If
                                If donatepurpose<>"" Then Donate_Where=Donate_Where&"And Donate_Purpose In "&donatepurpose&" "
                              End If
                              If Request("Donate_Type")<>"" Then
                                donatetype=""
                                Ary_Donate_Type=Split(Request("Donate_Type"),",")
                                If UBound(Ary_Donate_Type)=0 Then
                                  donatetype=donatetype&"('"&Trim(Ary_Donate_Type(0))&"')"
                                Else
                                  For I = 0 To UBound(Ary_Donate_Type)
                                    If I=0 Then
                                      donatetype=donatetype&"('"&Trim(Ary_Donate_Type(I))&"',"
                                    ElseIf I=UBound(Ary_Donate_Type) Then 
                                      donatetype=donatetype&"'"&Trim(Ary_Donate_Type(I))&"')" 
                                    Else
                                      donatetype=donatetype&"'"&Trim(Ary_Donate_Type(I))&"',"
                                    End If
                                  Next
                                End If
                                If donatetype<>"" Then Donate_Where=Donate_Where&"And Donate_Type In "&donatetype&" "
                              End If
                              If Request("Accoun_Bank")<>"" Then
                                accounbank=""
                                Ary_Accoun_Bank=Split(Request("Accoun_Bank"),",")
                                If UBound(Ary_Accoun_Bank)=0 Then
                                  accounbank=accounbank&"('"&Trim(Ary_Accoun_Bank(0))&"')"
                                Else
                                  For I = 0 To UBound(Ary_Accoun_Bank)
                                    If I=0 Then
                                      accounbank=accounbank&"('"&Trim(Ary_Accoun_Bank(I))&"',"
                                    ElseIf I=UBound(Ary_Accoun_Bank) Then 
                                      accounbank=accounbank&"'"&Trim(Ary_Accoun_Bank(I))&"')" 
                                    Else
                                      accounbank=accounbank&"'"&Trim(Ary_Accoun_Bank(I))&"',"
                                    End If
                                  Next
                                End If
                                If accounbank<>"" Then Donate_Where=Donate_Where&"And Accoun_Bank In "&accounbank&" "
                              End If
                              If Request("Accoun_Date_Begin")<>"" Then Donate_Where=Donate_Where&"And Accoun_Date>='"&Request("Accoun_Date_Begin")&"' "
                              If Request("Accoun_Date_End")<>"" Then Donate_Where=Donate_Where&"And Accoun_Date<='"&Request("Accoun_Date_End")&"' "
                              If Request("Accounting_Title")<>"" Then
                                accountingtitle=""
                                Ary_Accounting_Title=Split(Request("Accounting_Title"),",")
                                If UBound(Ary_Accounting_Title)=0 Then
                                  accountingtitle=accountingtitle&"('"&Trim(Ary_Accounting_Title(0))&"')"
                                Else
                                  For I = 0 To UBound(Ary_Accounting_Title)
                                    If I=0 Then
                                      accountingtitle=accountingtitle&"('"&Trim(Ary_Accounting_Title(I))&"',"
                                    ElseIf I=UBound(Ary_Accounting_Title) Then 
                                      accountingtitle=accountingtitle&"'"&Trim(Ary_Accounting_Title(I))&"')" 
                                    Else
                                      accountingtitle=accountingtitle&"'"&Trim(Ary_Accounting_Title(I))&"',"
                                    End If
                                  Next
                                End If
                                If accountingtitle<>"" Then Donate_Where=Donate_Where&"And Accounting_Title In "&accountingtitle&" "
                              End If
                              If request("Issue_TypeM")<>"" Then Donate_Where=Donate_Where&"And Issue_Type = '"&request("Issue_TypeM")&"' "
                              If request("Issue_TypeD")<>"" Then Donate_Where=Donate_Where&"And Issue_Type = '"&request("Issue_TypeD")&"' "
                              Donate_Purpose_Type=""
                              Ary_Purpose_Type=Split(request("Donate_Purpose_Type"),",")
                              For I = 0 To UBound(Ary_Purpose_Type)
                                If Donate_Purpose_Type="" Then
                                   Donate_Purpose_Type="'"&Trim(Ary_Purpose_Type(I))&"'"
                                Else
                                  Donate_Purpose_Type=Donate_Purpose_Type&",'"&Trim(Ary_Purpose_Type(I))&"'"
                                End If
                              Next                              
                              If Donate_Purpose_Type<>"" Then Donate_Where=Donate_Where&"And Donate_Purpose_Type In ("&Donate_Purpose_Type&") "

                              '搜尋條件二
                              Donor_Where=""
                              If Request("Donor_Name")<>"" Then Donor_Where=Donor_Where&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Donor.Invoice_Title Like '%"&request("Donor_Name")&"%') " 
                              If Request("Category")<>"" Then
                                donorcategory=""
                                Ary_Category=Split(Request("Category"),",")
                                If UBound(Ary_Category)=0 Then
                                  donorcategory=donorcategory&"('"&Trim(Ary_Category(0))&"')"
                                Else
                                  For I = 0 To UBound(Ary_Category)
                                    If I=0 Then
                                      donorcategory=donorcategory&"('"&Trim(Ary_Category(I))&"',"
                                    ElseIf I=UBound(Ary_Category) Then 
                                      donorcategory=donorcategory&"'"&Trim(Ary_Category(I))&"')" 
                                    Else
                                      donorcategory=donorcategory&"'"&Trim(Ary_Category(I))&"',"
                                    End If
                                  Next
                                End If
                                If donorcategory<>"" Then Donor_Where=Donor_Where&"And Category In "&donorcategory&" "
                              End If
                              If Request("Donor_Type")<>"" Then
                                donortype=""
                                Ary_Donor_Type=Split(Request("Donor_Type"),",")
                                If UBound(Ary_Donor_Type)=0 Then
                                  donortype=donortype&"And Donor_Type Like '%"&Trim(Ary_Donor_Type(0))&"%' "
                                Else
                                  For I = 0 To UBound(Ary_Donor_Type)
                                    If I=0 Then
                                      donortype=donortype&"And (Donor_Type Like '%"&Trim(Ary_Donor_Type(I))&"%' "
                                    ElseIf I=UBound(Ary_Donor_Type) Then 
                                      donortype=donortype&"Or Donor_Type Like '%"&Trim(Ary_Donor_Type(I))&"%') "
                                    Else
                                      donortype=donortype&"Or Donor_Type Like '%"&Trim(Ary_Donor_Type(I))&"%' "
                                    End If
                                  Next
                                End If
                                If donortype<>"" Then Donor_Where=Donor_Where&donortype
                              End If
                              If Request("City")<>"" Then Donor_Where=Donor_Where&"And City='"&Request("City")&"' "
                              If Request("Area")<>"" Then Donor_Where=Donor_Where&"And Area='"&Request("Area")&"' "
                              If Request("Invoice_City")<>"" Then Donor_Where=Donor_Where&"And City='"&Request("Invoice_City")&"' "
                              If Request("Invoice_Area")<>"" Then Donor_Where=Donor_Where&"And Invoice_Area='"&Request("Invoice_Area")&"' "                              
                            
                              '排序方式
                              OrderBy_Where="Order By Donate_Date Desc,Donate_Id Desc" 
   						                
   						                SQL="Select Donate_Id,捐款人=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),捐款日期=CONVERT(VarChar,Donate_Date,111),捐款金額=Isnull(Donate_Amt,0),捐款方式=Donate_Payment,捐款用途=Donate_Purpose,收據開立=DONATE.Invoice_Type,收據編號=Invoice_Pre+Invoice_No,沖帳日期=CONVERT(VarChar,Accoun_Date,111),捐款類別=Donate_Type,入帳銀行=Accoun_Bank,會計科目=Accounting_Title,活動名稱=Act_ShortName,狀態=(Case When Issue_Type='M' Then '手開' Else Case When Issue_Type='D' Then '作廢' Else '' End End) " & _
                                  "From DONATE Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id " & _
                                  "Left Join ACT On DONATE.Act_Id=ACT.Act_Id " & _
                                  "Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
                                  "Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode " & _
                                  "Where 1=1 "&Donate_Where&Donor_Where&OrderBy_Where&" "
                            Else
                              SQL=request("SQL")
                            End If
                            Session("SQL")=SQL
                            Session("Donate_Date_Begin")=Request("Donate_Date_Begin")
                            Session("Donate_Date_End")=Request("Donate_Date_End")
                            Session("Act_Id")=Request("Act_Id")
                            PageSize=20
                            If request("nowPage")="" Then
                              nowPage=1
                            Else
                              nowPage=request("nowPage")
                            End If
                            ProgID="donate_name_list_qry"
                            HLink=""
                            LinkParam="Donate_Id"
                            LinkTarget="main"
                            AddLink="donate_name_list.asp"
                            Session("SQL")=SQL
                            If request("action")="report" Then
					                    Response.Redirect "../include/movebar.asp?URL=../donation/donate_name_list_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "donate_name_list_rpt.asp?action=export"
					                  End If 
                            call DonateGridList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
                  	      %>
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
  </center></div></p>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->