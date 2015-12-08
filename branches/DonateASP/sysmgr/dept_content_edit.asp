<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From DEPT Where Dept_Id='"&request("dept_id")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("Dept_Id")=request("Dept_Id")
  RS("Comp_Name")=request("Comp_Name")
  RS("Comp_ShortName")=request("Comp_ShortName")
  RS("Dept_Desc")=request("Dept_Desc")
  RS("DEPT_LOGO")=""
  RS("Server_Url")=request("Server_Url")
  RS("Sys_Name")=request("Sys_Name")
  RS("Mail_Url")=request("Mail_Url")
  RS("Mail_SendType")=request("Mail_SendType")
  RS("ContaCtor")=request("ContaCtor")
  RS("Tel")=request("Tel")
  RS("Fax")=request("Fax")
  RS("EMail")=request("EMail")
  RS("Zip_Code")=request("Zip_Code")
  RS("City_Code")=request("City_Code")
  RS("Area_Code")=request("Area_Code")
  RS("Address")=request("Address")
  RS("Uniform_No")=request("Uniform_No")
  RS("Account")=request("Account")
  If request("CredteDate")<>"" Then
    RS("CredteDate")=request("CredteDate")
  Else
    RS("CredteDate")=null
  End If
  RS("Licence")=request("Licence")
  RS("Password_Day")=request("Password_Day")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 !"
  response.redirect "dept.asp"  
End If

If request("action")="delete" Then
  SQL="Delete From DEPT Where DEPT_ID='"&request("dept_id")&"' "
  Set RS=Conn.Execute(SQL)
  session("errnumber")=1
  session("msg")="資料刪除成功 !"
  Response.Redirect "dept.asp"
End If

SQL="Select * From DEPT Where DEPT_ID='"&request("dept_id")&"'"
call QuerySQL(SQL,RS) 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>機構部門管理</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
  <p><div align="center"><center>
    <form name="form" action="" method="post">
      <input type="hidden" name="action">
      <div align="center"><center>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
          <tr>
            <td>
              <table width="780"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
                <tr>
                  <td width="5%"></td>
                  <td width="95%">
  		              <table width="40%"  border="0" cellspacing="0" cellpadding="0">
		                  <tr>
                        <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                        <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                        <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                      </tr>
                      <tr>
                        <td class="table62-bg"></td>
                        <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">
						機構參數設定</td>
                        <td class="table63-bg"></td>
                      </tr>  
    	            </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td> 
	            <table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
                <tr>
                  <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                  <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                  <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                </tr>
                <tr>
                  <td height="445" valign="top" class="table62-bg">&nbsp;</td>
                  <td height="445" valign="top">              
                    <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                      <tr>
                        <td class="td02-c" width="100%" valign="top" colspan="3">
                          <table border="1" cellpadding="2" style="border-collapse: collapse" width="100%" height="25" cellspacing="1">
                            <tr>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">機構資訊設定</a></td>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_system_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">網站參數設定</a></td>
                              <%If Session("user_id")="npois" And (Request.ServerVariables("REMOTE_HOST")="60.250.147.33" Or Request.ServerVariables("REMOTE_HOST")="127.0.0.1") Then%>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_donate_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">捐款參數設定</a></td>
                              <td class="button2-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_content_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool"><font color="#800000"><b>捐款機制說明</b></font></a></td>
                              <%Else%>
                              <td class="button2-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_content_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool"><font color="#800000"><b>捐款機制說明</b></font></a></td>
                              <%End If%>
                            </tr>
                          </table>
                        </td>
                      </tr>
                      <tr>
                        <td valign="top">&nbsp;</td>
                      </tr>                  
                      <tr>
                        <td class="td02-c" width="100%">
                          <table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">                        
            
                          </table>
                        </td>
                      </tr>
                      <tr>
                        <td class="td02-c" width="100%" align="center"><button id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Update_OnClick()'> <img src='../images/update.gIf' width='20' height='20' align='absmiddle'> 
						修改</button>&nbsp;<button id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Cancel_OnClick()'> <img src='../images/icon6.gIf' width='20' height='15' align='absmiddle'> 
						離開</button>&nbsp;</td>
                      </tr>
                    </table>
                  </td>
                  <td height="445" class="table63-bg">&nbsp;</td>
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
      </center></div>   
    </form>
  </center></div>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Update_OnClick(){
  <%call SubmitJ("update")%>
}

function Cancel_OnClick(){
  location.href="dept.asp"
}
--></Script>