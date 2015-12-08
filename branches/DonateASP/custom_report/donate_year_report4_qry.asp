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
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">國內地區捐款總額明細表</td>
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
							SQL1="SELECT D.Donor_Id AS '編號'  " & _
                                    ",M.Donor_Name AS '天使姓名' ,M.Title AS '稱謂' " & _
								    ",D.Sum_Amt AS '累計奉獻金額' ,D.Times AS '奉獻次數' " & _
								    ",IsNull(M.ZipCode,'') AS '郵遞區號' " & _
									",IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address AS '地址' " & _
								    "FROM DONOR M " & _
								    "INNER JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times " & _
									"FROM DONATE "
									
                              '搜尋條件
                              Donate_Where=""
							  If Request("Donate_Amt")<>"" Then 
								SQL1=SQL1&"WHERE Donate_Amt > '"&Request("Donate_Amt")&"' "
							  Else
								SQL1=SQL1&"WHERE Donate_Amt > 0 "
							  End IF
							  
							  If Request("Donate_Date_Begin")<>"" Then SQL1=SQL1&"And Donate_Date>='"&Request("Donate_Date_Begin")&"' "
                              If Request("Donate_Date_End")<>"" Then SQL1=SQL1&"And Donate_Date<='"&Request("Donate_Date_End")&"' "
							  
                              If Request("Donate_Total_Amt")<>"" Then 
								SQL1=SQL1&"GROUP BY Donor_Id HAVING SUM(Donate_Amt) > '"&Request("Donate_Total_Amt")&"' ) D "
							  Else
								SQL1=SQL1&"GROUP BY Donor_Id ) D "
							  End IF
							  
							  SQL1=SQL1&"ON M.Donor_Id = D.Donor_Id " & _
										"LEFT JOIN dbo.CODECITYNew C1 " & _
										"ON M.City = C1.ZipCode " & _
										"LEFT JOIN dbo.CODECITYNew C2 " & _
										"ON M.Area = C2.ZipCode " & _
										"WHERE M.IsAbroad = 'N' " 
										
                              If Request("Zip_Code")<>"" Then
                                zipcode=""
                                Ary_Zip_Code=Split(Request("Zip_Code"),",")
                                If UBound(Ary_Zip_Code)=0 Then
                                  zipcode=zipcode&"('"&Trim(Ary_Zip_Code(0))&"')"
                                Else
                                  For I = 0 To UBound(Ary_Zip_Code)
                                    If I=0 Then
                                      zipcode=zipcode&"('"&Trim(Ary_Zip_Code(I))&"',"
                                    ElseIf I=UBound(Ary_Zip_Code) Then 
                                      zipcode=zipcode&"'"&Trim(Ary_Zip_Code(I))&"')" 
                                    Else
                                      zipcode=zipcode&"'"&Trim(Ary_Zip_Code(I))&"',"
                                    End If
                                  Next
                                End If
                                If zipcode<>"" Then SQL1=SQL1&"AND C1.Name In "&zipcode&" "
                              Else
                                SQL1=SQL1
                              End If
							  If Request("Is_ErrAddress")<>"" Then SQL1=SQL1&"And IsErrAddress='N' "
							  If Request("Sex")<>"" Then 
								SQL1=SQL1&"And IsNull(Sex,'') <> '歿' Order by D.Donor_Id "
							  Else 
								SQL1=SQL1&"Order by D.Donor_Id "	
							  End If

							'response.write SQL1
							'response.end()
                            Else
                              SQL1=request("SQL1")
                            End If
                            Session("SQL1")=SQL1
                            Session("Donate_Date_Begin")=Request("Donate_Date_Begin")
                            Session("Donate_Date_End")=Request("Donate_Date_End")
							Session("Donate_Total_Begin")=Request("Donate_Total_Begin")
                            Session("Donate_Total_End")=Request("Donate_Total_End")
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
					                    Response.Redirect "../include/movebar.asp?URL=../custom_report/donate_year_report4_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "donate_year_report4_rpt.asp?action=export"
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