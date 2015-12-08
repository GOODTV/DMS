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
							SQL1="Select  convert(nvarchar(10),A.Donate_Date,111) AS 'Donate_Date' " & _
								 ",Sum(Payment01) AS '現金',Sum(Payment02) AS '劃撥',Sum(Payment03) AS '信用卡授權書(一般)' " & _
								 ",Sum(Payment04) AS '郵局轉帳',Sum(Payment05) AS '匯款',Sum(Payment06) AS '支票' " & _
								 ",Sum(Payment07) AS '實物奉獻',Sum(Payment08) AS 'ATM',Sum(Payment09) AS '網路信用卡' " & _
								 ",Sum(Payment10) AS 'ACH',Sum(Payment11) AS '美國運通' " & _
								 ",Sum(Payment01)+Sum(Payment02)+Sum(Payment03)+Sum(Payment04)+Sum(Payment05)+Sum(Payment06)+Sum(Payment07) " & _
								 "+Sum(Payment08)+Sum(Payment09)+Sum(Payment10)+Sum(Payment11) AS '小計' " & _
								 "FROM " & _
								 "(Select Donate_Date " & _
								 ", Case When Donate_Payment='現金' Then Donate_Amt Else 0 End 'Payment01' " & _
								 ", Case When Donate_Payment='劃撥' Then Donate_Amt Else 0 End 'Payment02' " & _
								 ", Case When Donate_Payment='信用卡授權書(一般)' Then Donate_Amt Else 0 End 'Payment03' " & _
								 ", Case When Donate_Payment='郵局帳戶轉帳授權書' Then Donate_Amt Else 0 End 'Payment04' " & _
								 ", Case When Donate_Payment='匯款' Then Donate_Amt Else 0 End 'Payment05' " & _
								 ", Case When Donate_Payment='支票' Then Donate_Amt Else 0 End 'Payment06' " & _
								 ", Case When Donate_Payment='實物奉獻' Then Donate_Amt Else 0 End 'Payment07' " & _
								 ", Case When Donate_Payment='ATM' Then Donate_Amt Else 0 End 'Payment08' " & _
								 ", Case When Donate_Payment='網路信用卡' Then Donate_Amt Else 0 End 'Payment09' " & _
								 ", Case When Donate_Payment like '%ACH%' Then Donate_Amt Else 0 End 'Payment10' " & _
								 ", Case When Donate_Payment='信用卡授權書(聯信)' Then Donate_Amt Else 0 End 'Payment11',Issue_type " & _
								 "From dbo.DONATE ) A "

								 
								If Request("Year")<>""  AND Request("Month")<>"" Then
									SQL1=SQL1&"where DATEPART(YEAR,Donate_Date) = '"&Request("Year")&"' And DATEPART(MONTH,Donate_Date) = '"&Request("Month")&"' and ISNULL(A.Issue_Type,'') <> 'D' GROUP BY convert(nvarchar(10),Donate_Date,111)  ORDER BY convert(nvarchar(10),Donate_Date,111) "
								Else
									SQL1=SQL1&"GROUP BY convert(nvarchar(10),Donate_Date,111)  ORDER BY convert(nvarchar(10),Donate_Date,111) "
								End If
								
								'response.write SQL1
								'response.end()
								SQL2="Select ISNULL([現金],0) As '現金', ISNULL([劃撥],0) As '劃撥', ISNULL([信用卡授權書(一般)],0) As '信用卡授權書(一般)', " & _
									 "ISNULL([郵局帳戶轉帳授權書],0) As '郵局轉帳', ISNULL([匯款],0) As '匯款', ISNULL([支票],0) As '支票', " & _
									 "ISNULL([實物奉獻],0) As '實物奉獻', ISNULL([ATM],0) As 'ATM', ISNULL([網路信用卡],0) As '網路信用卡', " & _
									 "ISNULL([ACH轉帳授權書],0) As 'ACH', ISNULL([信用卡授權書(聯信)],0) As '美國運通', " & _
									 "ISNULL([現金],0)+ISNULL([劃撥],0)+ISNULL([信用卡授權書(一般)],0)+ISNULL([郵局帳戶轉帳授權書],0)+ISNULL([匯款],0)+ISNULL([支票],0)+ISNULL([實物奉獻],0)+ISNULL([ATM],0)+ISNULL([網路信用卡],0)+ISNULL([ACH轉帳授權書],0)+ISNULL([信用卡授權書(聯信)],0) AS '總計' " & _
									 "FROM ( select Donate_Payment,Count(Donate_Id) AS 'Cnt' from donate " & _
									 "where DATEPART(YEAR,Donate_Date) = '"&Request("Year")&"' And DATEPART(MONTH,Donate_Date) = '"&Request("Month")&"' and ISNULL(Issue_Type,'') <> 'D' " & _
									 "group by Donate_Payment ) As S " & _
									 "PIVOT ( Sum(Cnt) FOR " & _
									 "Donate_Payment IN ([現金], [劃撥], [信用卡授權書(一般)], [郵局帳戶轉帳授權書], [匯款], [支票], [實物奉獻], [ATM], [網路信用卡], [ACH轉帳授權書], [信用卡授權書(聯信)]) " & _
									 ") As P "
								
                            Else
                              SQL1=request("SQL1")
                            End If
                            Session("SQL1")=SQL1
							Session("SQL2")=SQL2
                            Session("DeptId")=Request("Dept_Id")
                            Session("Year")=Request("Year")
                            Session("Month")=Request("Month")

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

                            If request("action")="report" Then
								Response.Redirect "../include/movebar.asp?URL=../custom_report/donate_month_report5_rpt.asp"
							ElseIf request("action")="export" then
								Response.Redirect "donate_month_report5_rpt.asp?action=export"
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