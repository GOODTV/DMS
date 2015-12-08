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
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">查詢單筆捐款金額</td>
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
							SQL1="SELECT D.Donor_Id AS '編號',D.Donate_Date AS '捐款日期'  " & _
                                    ",M.Donor_Name AS '天使姓名' ,M.Title AS '稱謂' " & _
									",D.Donate_Amt AS '捐款金額' ,Case When D2.Times > D1.Times Then '舊' Else '新' END '舊/新大額' " & _
								    ",Case When M.Donor_Name like '%主知名' then D.Donate_Date ELSE Begin_DonateDate END AS '首捐日期' " & _
									",Case When M.Donor_Name like '%主知名' then '' ELSE D3.Sum_Amt END AS '累計金額',D.Donate_Purpose AS '捐款用途' " & _
									",IsNull(M.ZipCode,'') AS '郵遞區號' " & _
								    ",Case When M.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') " & _
											"Else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址' " & _
									",Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = ''  " & _
											"Then (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office Else Tel_Office + ' Ext.' + Tel_Office_Ext End) " & _
									"Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  " & _
											"Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話' ,IsNull(Cellular_Phone,'') AS '手機' " & _
									"FROM dbo.DONOR M " & _
									"inner JOIN (SELECT Donor_Id ,Donate_Id,Donate_Date,Donate_Amt,Donate_Purpose,Donate_Payment " & _
									"FROM DONATE "
									
                              '搜尋條件
                              Donate_Where=""
							  Donate_Where1=""
							  Donate_Where2=""
							  
                              If Request("Donate_Amt")<>"" Then 
								Donate_Where=Donate_Where&"WHERE Donate_Amt >= "&Request("Donate_Amt")&" "
							  Else
								Donate_Where=Donate_Where&"WHERE Donate_Amt > 0 "
							  End IF
							  
							  If Request("Donate_Date_Begin")<>"" Then Donate_Where1=Donate_Where1&"And Donate_Date>='"&Request("Donate_Date_Begin")&"' "

                              If Request("Donate_Date_End")<>"" Then Donate_Where1=Donate_Where1&"And Donate_Date<='"&Request("Donate_Date_End")&"' " 
							  
							  If Request("Accumulate_Date_Begin")<>"" Then Donate_Where2=Donate_Where2&"And Donate_Date>='"&Request("Accumulate_Date_Begin")&"' "
							  If Request("Accumulate_Date_End")<>"" Then Donate_Where2=Donate_Where2&"And Donate_Date<='"&Request("Accumulate_Date_End")&"' "
							  
							  SQL1=SQL1&Donate_Where&Donate_Where1&") D ON M.Donor_Id = D.Donor_Id " & _
									"left JOIN (SELECT Donor_Id ,Count(Donor_Id) AS Times " & _
									"FROM dbo.DONATE "
							  SQL1=SQL1&Donate_Where&Donate_Where1&"group by Donor_Id ) D1 ON M.Donor_Id = D1.Donor_Id " & _
									"left JOIN (SELECT Donor_Id ,Count(Donor_Id) AS Times " & _
									"FROM dbo.DONATE "
							  SQL1=SQL1&Donate_Where&Donate_Where2&"group by Donor_Id ) D2 ON M.Donor_Id = D2.Donor_Id " & _
									"left JOIN (SELECT Donor_Id , Sum(Donate_Amt) AS Sum_Amt " & _
									"FROM dbo.DONATE "
							  SQL1=SQL1&"WHERE Donate_Amt > 0 "&Donate_Where2&"group by Donor_Id ) D3 ON M.Donor_Id = D3.Donor_Id " & _
									"LEFT JOIN dbo.CODECITYNew C1 " & _
									"ON M.City = C1.ZipCode " & _
									"LEFT JOIN dbo.CODECITYNew C2 " & _
									"ON M.Area = C2.ZipCode "
							  If Request("Is_Abroad")<>"" Then SQL1=SQL1&"And IsAbroad='N' "
							  If Request("Is_ErrAddress")<>"" Then SQL1=SQL1&"And IsErrAddress='N' "
							  If Request("Sex")<>"" Then 
								SQL1=SQL1&"And IsNull(Sex,'') <> '歿' ORDER BY D.Donate_Purpose,D.Donate_Amt,D.Donate_Date,Last_DonateDate "
							  Else 
								SQL1=SQL1&"ORDER BY D.Donate_Purpose,D.Donate_Amt,D.Donate_Date,Last_DonateDate "	
							  End If				

							'response.write SQL1
							'response.end()
                            Else
                              SQL1=request("SQL1")
                            End If
                            Session("SQL1")=SQL1
                            Session("Donate_Date_Begin")=Request("Donate_Date_Begin")
                            Session("Donate_Date_End")=Request("Donate_Date_End")
							Session("Donate_Amt")=Request("Donate_Amt")

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
					                    Response.Redirect "../include/movebar.asp?URL=../custom_report/donate_week_report2_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "donate_week_report2_rpt.asp?action=export"
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