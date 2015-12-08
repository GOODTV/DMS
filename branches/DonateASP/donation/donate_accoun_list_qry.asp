<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="donate_accoun_list"%>
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
                              '搜尋條件
                              Donate_Where=""
                              If Request("Dept_Id")<>"" Then
                                Donate_Where=Donate_Where&"And Donate.Dept_Id = '"&Request("Dept_Id")&"' "
                              Else
                                Donate_Where=Donate_Where&"And Donate.Dept_Id In ("&Session("all_dept_type")&") "
                              End If
                              If Request("Donate_Date_Begin")<>"" Then Donate_Where=Donate_Where&"And Donate_Date>='"&Request("Donate_Date_Begin")&"' "
                              If Request("Donate_Date_End")<>"" Then Donate_Where=Donate_Where&"And Donate_Date<='"&Request("Donate_Date_End")&"' "
                              If Request("Act_Id")<>"" Then Donate_Where=Donate_Where&"And DONATE.Act_Id='"&Request("Act_Id")&"' "
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
                              Else
                                Donate_Where=Donate_Where&"And Donate_Payment <> '"&Request("DonateDesc")&"' "
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
                            Else
                              SQL1=request("SQL1")
                            End If
                            Session("Donate_Where")=Donate_Where
                            Session("DeptId")=Request("Dept_Id")
                            Session("Donate_Date_Begin")=Request("Donate_Date_Begin")
                            Session("Donate_Date_End")=Request("Donate_Date_End")
                            Session("Act_Id")=Request("Act_Id")
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
                            AddLink="donate_accoun_list.asp"
                            Session("SQL1")=SQL1
                            If request("action")="report" Then
					                    Response.Redirect "../include/movebar.asp?URL=../donation/donate_accoun_list_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "donate_accoun_list_rpt.asp?action=export"
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