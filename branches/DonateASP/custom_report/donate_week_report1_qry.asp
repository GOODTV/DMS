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
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">建台/非建台奉獻級距表</td>
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
							  '搜尋條件
                              Donate_Where=""
                              If Request("Donate_Date_Begin")<>"" and Request("Donate_Date_End")<>"" Then 
								Donate_Where=Donate_Where&"And Donate_Date>='"&Request("Donate_Date_Begin")&"' And Donate_Date <='"&Request("Donate_Date_End")&"' "
							  Else
								Donate_Where=Donate_Where
							  End If
                              
								If Request("Donate_Purpose")="建台" Then 
									SQL1="SELECT '1萬(不含)以下' as '級距',COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '小計'  " & _
											"FROM dbo.DONATE  " & _
											"WHERE Donate_Purpose = '建台' "&Donate_Where& _
												"AND Donate_Amt Between 1 and 9999 " & _
											"UNION " & _
											"SELECT '1萬(含)以上' as '級距', COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '小計' " & _
											"FROM dbo.DONATE " & _
											"WHERE Donate_Purpose = '建台' "&Donate_Where& _
												"AND Donate_Amt Between 10000 and 99999 " & _
											"UNION " & _
											"SELECT '10萬(含)以上' as '級距', COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '小計' " & _
											"FROM dbo.DONATE " & _
											"WHERE Donate_Purpose = '建台' "&Donate_Where& _
												"AND Donate_Amt Between 100000 and 999999 " & _
											"UNION " & _
											"SELECT '100萬(含)以上' as '級距', COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '小計' " & _
											"FROM dbo.DONATE " & _
											"WHERE Donate_Purpose = '建台' "&Donate_Where& _
												"AND Donate_Amt Between 1000000 and 9999999 " & _
											"UNION " & _
											"SELECT '1000萬(含)以上' as '級距', COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '小計' " & _
											"FROM dbo.DONATE " & _
											"WHERE Donate_Purpose = '建台' "&Donate_Where& _
												"AND Donate_Amt >= 10000000 " & _
											"UNION " & _
											"SELECT '總計' as '級距', COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '總金額' " & _
											"FROM dbo.DONATE " & _
											"WHERE Donate_Purpose = '建台' "&Donate_Where& _
												"AND Donate_Amt > 0 "
								Else
									SQL1="SELECT '1萬(不含)以下' as '級距',COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '小計'  " & _
											"FROM dbo.DONATE  " & _
											"WHERE Donate_Purpose <> '建台' "&Donate_Where& _
												"AND Donate_Amt Between 1 and 9999 " & _
											"UNION " & _
											"SELECT '1萬(含)以上' as '級距', COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '小計' " & _
											"FROM dbo.DONATE " & _
											"WHERE Donate_Purpose <> '建台' "&Donate_Where& _
												"AND Donate_Amt Between 10000 and 99999 " & _
											"UNION " & _
											"SELECT '10萬(含)以上' as '級距', COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '小計' " & _
											"FROM dbo.DONATE " & _
											"WHERE Donate_Purpose <> '建台' "&Donate_Where& _
												"AND Donate_Amt Between 100000 and 999999 " & _
											"UNION " & _
											"SELECT '100萬(含)以上' as '級距', COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '小計' " & _
											"FROM dbo.DONATE " & _
											"WHERE Donate_Purpose <> '建台' "&Donate_Where& _
												"AND Donate_Amt Between 1000000 and 9999999 " & _
											"UNION " & _
											"SELECT '1000萬(含)以上' as '級距', COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '小計' " & _
											"FROM dbo.DONATE " & _
											"WHERE Donate_Purpose <> '建台' "&Donate_Where& _
												"AND Donate_Amt >= 10000000 " & _
											"UNION " & _
											"SELECT '總計' as '級距', COUNT(*) as '筆數' ,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0) as '總金額' " & _
											"FROM dbo.DONATE " & _
											"WHERE Donate_Purpose <> '建台' "&Donate_Where& _
												"AND Donate_Amt > 0 "
								End IF
							
							'response.write SQL1
							'response.end()
                            Else
                              SQL1=request("SQL1")
                            End If
                            Session("SQL1")=SQL1
                            Session("Donate_Date_Begin")=Request("Donate_Date_Begin")
                            Session("Donate_Date_End")=Request("Donate_Date_End")
                            Session("Donate_Purpose")=Request("Donate_Purpose")
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
					                    Response.Redirect "../include/movebar.asp?URL=../custom_report/donate_week_report1_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "donate_week_report1_rpt.asp?action=export"
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