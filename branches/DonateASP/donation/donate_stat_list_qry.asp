<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="stat"%>
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
                    <td class="table62-bg">¡@</td>
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">²έp¤*R</td>
                    <td class="table63-bg">¡@</td>
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
              <td class="table62-bg">¡@</td>
              <td valign="top">
                <table width="100%"  border="0" cellspacing="0" style="border-collapse: collapse">
                  <tr>
                    <td class="td02-c" width="100%">
                      <table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                        <tr>
                          <td>
                  	      <%
                            If request("Dept_Id")<>"" And request("Donate_Amt_Kind")<>"" And request("Stat_Kind")="Donate_Amt" Then
                              SQL="Update DEPT Set Donate_Amt_Kind='"&request("Donate_Amt_Kind")&"' Where Dept_Id='"&request("Dept_Id")&"'"
                              Set RS=Conn.Execute(SQL)
                            End If
                            Session("DeptId")=request("Dept_Id")
                            Session("Donate_Date_Begin")=request("Donate_Date_Begin")
                            Session("Donate_Date_End")=request("Donate_Date_End")
                            Session("Act_Id")=request("Act_Id")
                            Session("Stat_Kind")=request("Stat_Kind")
                            Session("Donate_Amt_Kind")=request("Donate_Amt_Kind")
                            Session("Donate_Purpose_Type")=request("Donate_Purpose_Type")
                            If request("action")="report" Then
					                    Response.Redirect "../include/movebar.asp?URL=../donation/donate_stat_list_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "donate_stat_list_rpt.asp?action=export"
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
              <td class="table63-bg">¡@</td>
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