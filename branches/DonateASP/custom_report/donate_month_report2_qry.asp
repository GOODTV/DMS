<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="report_list"%>
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
                            If request("SQL1")="" Then
							SQL1="SELECT Case When C.Donate_FromDate > A.Donate_FromDate AND C.Donate_ToDate > A.Donate_ToDate " & _
										"Then 'V' Else '' END '新增' " & _
								    ",Case When C.Donate_FromDate > A.Donate_FromDate AND C.Donate_ToDate > A.Donate_ToDate " & _
										"Then C.Pledge_Id Else NULL END '新授權書號' " & _
								    ",A.Donate_ToDate AS '授權迄日' " & _
									",Case When C.Donate_FromDate > A.Donate_FromDate AND C.Donate_ToDate > A.Donate_ToDate " & _
										"Then C.Donate_Amt Else NULL END '新授權額' ,A.Donate_Payment AS '授權方式' " & _
									",A.Card_Bank AS '銀行名稱' ,A.Card_Type AS '卡別' " & _
									",A.Pledge_Id AS '授權書號' ,A.Donor_Id AS '編號' ,D.Donor_Name AS '天使姓名' " & _
									",Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then " & _
												"(Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office " & _
												"Else Tel_Office + ' Ext.' + Tel_Office_Ext End) " & _
										  "Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office " & _
												  "Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話' " & _
								    ",IsNull(Cellular_Phone,'') AS '手機',IsNull(D.ZipCode,'') AS '郵遞區號' " & _
									",Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') " & _
										  "Else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址' " & _
									"FROM dbo.PLEDGE A " & _
									"LEFT JOIN (SELECT Donor_Id,Donate_Payment,MAX(Pledge_Id) AS Pledge_Id " & _
												"FROM dbo.PLEDGE WHERE Donate_Type='長期捐款' " & _
												"GROUP BY Donor_Id,Donate_Payment) B " & _
									"ON A.Donor_Id = B.Donor_Id AND A.Donate_Payment = B.Donate_Payment " & _
									"LEFT JOIN (SELECT * FROM dbo.PLEDGE WHERE Donate_Type='長期捐款') C " & _
									"ON B.Donor_Id=C.Donor_Id AND B.Donate_Payment = C.Donate_Payment AND B.Pledge_Id = C.Pledge_Id " & _
									"LEFT JOIN dbo.DONOR D " & _
									"ON A.Donor_Id = D.Donor_Id " & _
									"LEFT JOIN dbo.CODECITYNew C1 " & _
									"ON D.City = C1.ZipCode " & _
									"LEFT JOIN dbo.CODECITYNew C2 " & _
									"ON D.Area = C2.ZipCode " & _
									"WHERE A.Donate_Type='長期捐款' "
									
                              '搜尋條件
                              'Donate_Where=""

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
                                If donatepayment<>"" Then SQL1=SQL1&"And A.Donate_Payment In "&donatepayment&" "
                              Else
                                SQL1=SQL1&"And A.Donate_Payment <> '' "
                              End If
							  
                              If Request("Year")<>""  AND Request("Month")<>"" Then 
								Donate_Date_Begin=DateAdd("M",0,Request("Year")&"/"&Request("Month")&"/1")
								Donate_Date_End=DateAdd("D",-1,DateAdd("M",1,Request("Year")&"/"&Request("Month")&"/1"))
								'response.write Donate_Date_Begin&" "&Donate_Date_End
								'response.end()
								SQL1=SQL1&"And A.Donate_ToDate >= '"&Donate_Date_Begin&"' And A.Donate_ToDate <= '"&Donate_Date_End&"' "
							  Else
								SQL1=SQL1
							  End If
							  If Request("Is_Abroad")<>"" Then SQL1=SQL1&"And D.IsAbroad='N' "
							  If Request("Is_ErrAddress")<>"" Then SQL1=SQL1&"And D.IsErrAddress='N' "
							  If Request("Sex")<>"" Then 
								SQL1=SQL1&"And IsNull(D.Sex,'') <> '歿' Order by A.Donor_Id "
							  Else 
								SQL1=SQL1&"Order by A.Donor_Id "	
							  End If                      

							'response.write SQL1
							'response.end()            

                            Else
                              SQL1=request("SQL1")
                            End If
                            Session("SQL1")=SQL1
                            Session("DeptId")=Request("Dept_Id")
                            Session("Year")=Request("Year")
                            Session("Month")=Request("Month")
                            Session("Donate_Payment")=Request("Donate_Payment")
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
                            AddLink="donate_month_list.asp"
                            Session("SQL1")=SQL1
                            If request("action")="report" Then
					                    Response.Redirect "../include/movebar.asp?URL=../custom_report/donate_month_report2_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "donate_month_report2_rpt.asp?action=export"
					                  End If 
                            call GridList_Q (SQL1,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
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