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
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">DVD贈品索取人報表</td>
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
							SQL1="SELECT distinct M.Donor_Id AS '編號' ,Donor_Name AS '天使姓名' ,ISNULL(Title,'') AS '稱謂'  " & _
									",IsNull(M.ZipCode,'') AS '郵遞區號' " & _
								    ",Case When M.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') " & _
											"Else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址' " & _
									",Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = ''  " & _
											"Then (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office Else Tel_Office + ' Ext.' + Tel_Office_Ext End) " & _
									"Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  " & _
											"Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話' " & _
									",G2.Goods_Qty AS '份數' ,CONVERT(varchar(12) , G2.Create_Date, 111 ) AS '資料輸入日期' " & _
									"FROM dbo.DONOR M " & _
									"INNER JOIN dbo.GIFT G ON M.Donor_Id=G.Donor_Id " & _
									"INNER JOIN dbo.GIFTData G2 ON M.Donor_Id=G2.Donor_Id  and G.Gift_Id=G2.Gift_Id " & _
									"INNER JOIN dbo.LINKED2 L2 ON L2.Linked_Id=G2.Goods_Id " & _
									"INNER JOIN dbo.LINKED L ON L.Ser_No=L2.Linked_Id " & _
									"LEFT JOIN dbo.CODECITYNew C1 " & _
									"ON M.City = C1.ZipCode " & _
									"LEFT JOIN dbo.CODECITYNew C2 " & _
									"ON M.Area = C2.ZipCode "
									
                              '搜尋條件
                              'Donate_Where=""
							  
                              If Request("Gift_Date_Begin")<>"" and Request("Gift_Date_End")<>"" Then SQL1=SQL1&"WHERE Gift_Date >= '"&Request("Gift_Date_Begin")&"' And Gift_Date <= '"&Request("Gift_Date_End")&"' "

							  If Request("Linked2_Name")<>"" Then SQL1=SQL1&"AND Goods_Name = '"&Request("Linked2_Name")&"' " 
							  

							  If Request("Is_Abroad")<>"" Then SQL1=SQL1&"And IsAbroad='N' "
							  If Request("Is_ErrAddress")<>"" Then SQL1=SQL1&"And IsErrAddress='N' "
							  If Request("Sex")<>"" Then 
								SQL1=SQL1&"And IsNull(Sex,'') <> '歿' ORDER BY M.Donor_Id "
							  Else 
								SQL1=SQL1&"ORDER BY M.Donor_Id "	
							  End If				

							'response.write SQL1
							'response.end()
                            Else
                              SQL1=request("SQL1")
                            End If
                            Session("SQL1")=SQL1
                            Session("Gift_Date_Begin")=Request("Gift_Date_Begin")
							Session("Gift_Date_End")=Request("Gift_Date_End")
                            Session("Linked2_Name")=Request("Linked2_Name")

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
					                    Response.Redirect "../include/movebar.asp?URL=../custom_report/donate_other_report2_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "donate_other_report2_rpt.asp?action=export"
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