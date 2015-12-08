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
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">郵寄月刊明細表</td>
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
							
							SQL1="SELECT MAIN.Donor_Id AS '編號' ,Donor_Name + ' ' + ISNULL(Title,'') AS '天使姓名'  " & _
                                    ",MAIN.ZipCode AS '郵遞區號' " & _
								    ",Case When MAIN.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') " & _
											"Else IsNull(C1.Name,'') + IsNull(C2.Name,'') + " & _
                                "(case when ISNULL([Attn],'') = '' then ISNULL([Address],'') " & _
                                " when Charindex([Attn],[Address]) <> 0 then ISNULL([Address],'') " & _
                                " else ISNULL([Address],'')+[Attn] end) END '地址' " & _
								    ",IsNull(Donor_Type,'') AS '身份別' ,IsNull(IsSendNewsNum,'') AS '月刊份數' " & _
									",Case when IsNull(IsContact,'') = 'N' Then '是' Else '' End '不主動聯絡' " & _
									"FROM DONOR MAIN " & _
								"LEFT JOIN dbo.CODECITYNew C1 " & _
									"ON MAIN.City = C1.ZipCode " & _
								"LEFT JOIN dbo.CODECITYNew C2 " & _
									"ON MAIN.Area = C2.ZipCode " & _
									"WHERE IsSendNews='Y' "
									
                              '搜尋條件
                              'Donate_Where=""
							  
							  If Request("IsMember") ="Y" Then
                                SQL1=SQL1&"And IsMember = 'Y' "
                              Else
                                SQL1=SQL1&"And IsMember = 'N' "
                              End If
							  If Request("Is_Abroad")<>"" Then SQL1=SQL1&"And IsAbroad='N' "
							  If Request("Is_ErrAddress")<>"" Then SQL1=SQL1&"And IsErrAddress='N' "
							  If Request("Sex")<>"" Then 
								SQL1=SQL1&"And IsNull(Sex,'') <> '歿' Order by MAIN.Donor_Id "
							  Else 
								SQL1=SQL1&"Order by MAIN.Donor_Id "	
							  End If

									
							'response.write SQL1
							'response.end()
							
                            Else
                              SQL1=request("SQL1")
							
                            End If
                            Session("SQL1")=SQL1
                            Session("DeptId")=Request("Dept_Id")
							Session("IsMember")=Request("IsMember")
							
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
					                    Response.Redirect "../include/movebar.asp?URL=../custom_report/donate_month_report3_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "donate_month_report3_rpt.asp?action=export"
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