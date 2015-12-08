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
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">雷同資料查詢</td>
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
							SQL1="Select Donor_Id AS '捐款人編號',Donor_Name AS '捐款人姓名',ISNULL(Sex,'') AS '性別'  " & _
									",Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = ''  " & _
											"Then (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office Else Tel_Office + ' Ext.' + Tel_Office_Ext End) " & _
									"Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  " & _
											"Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話' " & _
									",IsNull(Cellular_Phone,'') AS '手機' " & _
									",Case When M.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') " & _
											"Else IsNull(M.ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '通訊地址' " & _
									",Case When IsAbroad = 'Y' Then IsNull(Invoice_OverseasCountry,'') + ' ' + IsNull(Invoice_OverseasAddress,'') " & _
											"Else IsNull(M.Invoice_ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Invoice_Address END ' 收據地址' " & _
									"FROM dbo.DONOR M " & _
									"LEFT JOIN dbo.CODECITYNew C1 " & _
									"ON M.City = C1.ZipCode " & _
									"LEFT JOIN dbo.CODECITYNew C2 " & _
									"ON M.Area = C2.ZipCode " & _
									"Where Donor_Id In " & _
									 "(Select D1.Donor_Id From dbo.DONOR D1 " & _
									  "Left Join dbo.DONOR D2 "
									
                              '搜尋條件
                              'Donate_Where=""
							  
                              If Request("Is_SameName")<>"" and Request("Is_SameAddress")="" Then '只選同姓名
								SQL1=SQL1&"ON D1.Donor_Name = D2.Donor_Name " & _
										  "Where D1.Donor_Id <> D2.Donor_Id " & _
									      "Group by D1.Donor_Id) " & _
									      "Order by Donor_Name, Donor_Id, IsNull(M.ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address "
							  ElseIf Request("Is_SameName")="" and Request("Is_SameAddress")<>"" Then '只選同地址
								If Request("Address")="1" Then '通訊地址
									SQL1=SQL1&"On IsNull(D1.City,'') = IsNull(D2.City,'') And IsNull(D1.Area,'') = IsNull(D2.Area,'') And IsNull(D1.Address,'') = IsNull(D2.Address,'') " & _
											  "Where D1.Donor_Id <> D2.Donor_Id " & _
											  "And (IsNull(D1.Address,'') <> '' Or IsNull(D1.OverseasAddress,'') <> '') " & _
											  "Group by D1.Donor_Id) " & _
											  "Order by IsNull(M.ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address, Donor_Id, Donor_Name "	
								Else '收據地址
									SQL1=SQL1&"On IsNull(D1.Invoice_City,'') = IsNull(D2.Invoice_City,'') And IsNull(D1.Invoice_Area,'') = IsNull(D2.Invoice_Area,'') And IsNull(D1.Invoice_Address,'') = IsNull(D2.Invoice_Address,'') " & _
											  "Where D1.Donor_Id <> D2.Donor_Id " & _
											  "And (IsNull(D1.Invoice_Address,'') <> '' Or IsNull(D1.Invoice_OverseasAddress,'') <> '') " & _
											  "Group by D1.Donor_Id) " & _
											  "Order by IsNull(M.Invoice_ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Invoice_Address, Donor_Id, Donor_Name "
								End If
							  Else '同時選同姓名、同地址
								SQL1=SQL1&"ON D1.Donor_Name = D2.Donor_Name "
								If Request("Address")="1" Then '通訊地址
									SQL1=SQL1&"And IsNull(D1.Address,'') = IsNull(D2.Address,'') " & _
											  "Where D1.Donor_Id <> D2.Donor_Id " & _
											  "And (IsNull(D1.Address,'') <> '' Or IsNull(D1.OverseasAddress,'') <> '') " & _
											  "Group by D1.Donor_Id) " & _
											  "Order by IsNull(M.ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address, Donor_Name, Donor_Id "		  
								Else '收據地址
									SQL1=SQL1&"And IsNull(D1.Invoice_Address,'') = IsNull(D2.Invoice_Address,'') " & _
											  "Where D1.Donor_Id <> D2.Donor_Id " & _
											  "And (IsNull(D1.Invoice_Address,'') <> '' Or IsNull(D1.Invoice_OverseasAddress,'') <> '') " & _
											  "Group by D1.Donor_Id) " & _
											  "Order by IsNull(M.Invoice_ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Invoice_Address, Donor_Name, Donor_Id "
								End If
							  
							  End If			

							'response.write SQL1
							'response.end()
                            Else
                              SQL1=request("SQL1")
                            End If
                            Session("SQL1")=SQL1
                            Session("Is_SameName")=Request("Is_SameName")
							Session("Is_SameAddress")=Request("Is_SameAddress")
                            Session("Address")=Request("Address")

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
					                    Response.Redirect "../include/movebar.asp?URL=../custom_report/donate_other_report3_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "donate_other_report3_rpt.asp?action=export"
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