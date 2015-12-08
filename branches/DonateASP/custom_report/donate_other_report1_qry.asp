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
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">人數統計查詢報表</td>
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
							SQL2=""
							SQL3=""
							SQL4=""
							SQL5=""
							SQL6=""
                            If request("SQL1")="" Then

								SQL1="select COUNT(DISTINCT D.Donor_Id) AS '所有奉獻天使人數' ,COUNT(D1.Donor_Id) AS '累計奉獻10萬以上總人數' " & _
									 ",COUNT(D2.Donor_Id) AS '累計奉獻30萬以上總人數' ,COUNT(D3.Donor_Id) AS '累計奉獻50萬以上總人數' " & _
									 ",COUNT(D4.Donor_Id) AS '累計奉獻100萬以上總人數' , COUNT(M1.Donor_Type) AS '友好教會及機構數量' " & _
									 "from Donor M " & _
									 "LEFT JOIN " & _
									 "(SELECT Donor_Id, COUNT(DISTINCT Donor_Id) AS '奉獻人數' ,COUNT(Donate_Id) AS '奉獻筆數' ,SUM(Donate_Amt) AS '奉獻總金額' " & _
									 "FROM dbo.DONATE " & _
									 "WHERE Donate_Purpose <> '' "
								SQL2="GROUP BY Donor_Id )D " & _
									 "ON M.Donor_Id = D.Donor_Id " & _
									 "LEFT JOIN (SELECT Donor_Id  ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times " & _
									 "FROM dbo.DONATE " & _
									 "WHERE Donate_Amt >0 "
								SQL3="GROUP BY Donor_Id Having Sum(Donate_Amt) >= 100000 ) D1 " & _
									 "ON M.Donor_Id = D1.Donor_Id " & _
									 "LEFT JOIN (SELECT Donor_Id  ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times " & _
									 "FROM dbo.DONATE " & _
									 "WHERE Donate_Amt >0 "
								SQL4="GROUP BY Donor_Id Having Sum(Donate_Amt) >= 300000 ) D2 " & _
									 "ON M.Donor_Id = D2.Donor_Id " & _
									 "LEFT JOIN (SELECT Donor_Id  ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times " & _
									 "FROM dbo.DONATE " & _
									 "WHERE Donate_Amt >0 "
								SQL5="GROUP BY Donor_Id Having Sum(Donate_Amt) >= 500000 ) D3 " & _
									 "ON M.Donor_Id = D3.Donor_Id " & _
									 "LEFT JOIN (SELECT Donor_Id  ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times " & _
									 "FROM dbo.DONATE " & _
									 "WHERE Donate_Amt >0 "
								SQL6="GROUP BY Donor_Id Having Sum(Donate_Amt) >= 1000000 ) D4 " & _
									 "ON M.Donor_Id = D4.Donor_Id " & _
									 "LEFT JOIN (SELECT Donor_Id  ,Donor_Type FROM dbo.Donor " & _
									 "where Donor_type IN ('教會','福音機構') ) M1 " & _
									 "ON M.Donor_Id = M1.Donor_Id "
	
								'搜尋條件
								Donate_Where=""
								If Request("Donate_Date_Begin")<>"" AND Request("Donate_Date_End")<>"" Then 
									Donate_Where="And Donate_Date >= '"&Request("Donate_Date_Begin")&"' And Donate_Date <= '"&Request("Donate_Date_End")&"' "
									SQL1=SQL1&Donate_Where&SQL2&Donate_Where&SQL3&Donate_Where&SQL4&Donate_Where&SQL5&Donate_Where&SQL6
								Else
									SQL1=SQL1&SQL2&SQL3&SQL4&SQL5&SQL6
								End If

								'response.write SQL1
								'response.end()
                            Else
                              SQL1=request("SQL1")
                            End If
                            Session("SQL1")=SQL1
                            Session("Donate_Date_Begin")=Request("Donate_Date_Begin")
                            Session("Donate_Date_End")=Request("Donate_Date_End")

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
					                    Response.Redirect "../include/movebar.asp?URL=../custom_report/donate_other_report1_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "donate_other_report1_rpt.asp?action=export"
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