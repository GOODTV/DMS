<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From DEPT Where Dept_Id='"&request("dept_id")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  If request("Donate_Close")<>"" Then
    RS("Donate_Close")=request("Donate_Close")
  Else
    RS("Donate_Close")=null
  End If
  If request("Donate_Open_User")<>"" Then
    RS("Donate_Open_User")=request("Donate_Open_User")
    RS("Donate_Open_LastDate")=Date()
  Else
    RS("Donate_Open_User")=""
    RS("Donate_Open_LastDate")=null
  End If
  If request("IsContribute")="Y" Then  
    If request("Contribute_Close")<>"" Then
      RS("Contribute_Close")=request("Contribute_Close")
    Else
      RS("Contribute_Close")=null
    End If
    If request("Contribute_Open_User")<>"" Then
      RS("Contribute_Open_User")=request("Donate_Open_User")
      RS("Contribute_Open_LastDate")=Date()
    Else
      RS("Contribute_Open_User")=""
      RS("Contribute_Open_LastDate")=null
    End If     
  End If
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 !"
End If

SQL_IS="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"' "
Set RS_IS=Server.CreateObject("ADODB.Recordset")
RS_IS.Open SQL_IS,Conn,1,1
IsContribute=RS_IS("IsContribute")
RS_IS.Close
Set RS_IS=Nothing
                    
SQL="Select * From DEPT Where DEPT_ID='"&request("dept_id")&"'"
call QuerySQL(SQL,RS)
%>
<%Prog_Id="donate_close"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" action="" method="post">
    	<input type="hidden" name="Dept_Id" value="<%=request("dept_id")%>">
    	<input type="hidden" name="IsContribute" value="<%=IsContribute%>">
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
                        <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;"><%=Prog_Desc%>【修改】</td>
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
                        <td valign="top">&nbsp;</td>
                      </tr>                  
                      <tr>
                        <td class="td02-c" width="100%">
                          <table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">                        
                            <tr>
                              <td width="14%" align="right"><font color="#000080">機構簡稱<span lang="en-us">:</span></font></td>
                              <td width="86%"><input type="text" name="Comp_ShortName" size="21" class="font9t" readonly maxlength="30" value="<%=RS("Comp_ShortName")%>"></td>
                            </tr>
                            <tr>
                              <td width="14%" align="right"><font color="#000080">捐款關帳日<span lang="en-us">:</span></font></td>
                              <td width="86%">
                              	<%call Calendar("Donate_Close",RS("Donate_Close"))%>
                                &nbsp;&nbsp;&nbsp;<font color="#000080">開放權限修改人員(&nbsp;<font color="#ff0000">限<%=Date()%></font>&nbsp;)：</font>
                                <%
                                  If Session("user_id")<>"npois" Then
                                    SQL="Select Donate_Open_User=user_id,user_name From userfile Where user_id<>'npois' Order By dept_id,user_group"
                                  Else
                                    SQL="Select Donate_Open_User=user_id,user_name From userfile Order By dept_id,user_group"
                                  End If
                                  FName="Donate_Open_User"
                                  Listfield="user_name"
                                  menusize="1"
                                  If RS("Donate_Open_LastDate")<>"" Then
                                    If CDate(RS("Donate_Open_LastDate"))=CDate(Date()) Then
                                      BoundColumn=RS("Donate_Open_User")
                                    Else
                                      BoundColumn=""
                                    End If
                                  Else
                                    BoundColumn=""
                                  End If
                                  call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                                %>
                              </td>
                            </tr>
                            <%If IsContribute="Y" Then%>
                            <tr>
                              <td width="14%" align="right"><font color="#000080">捐物資關帳日<span lang="en-us">:</span></font></td>
                              <td width="86%">
                              	<%call Calendar("Contribute_Close",RS("Contribute_Close"))%>
                                &nbsp;&nbsp;&nbsp;<font color="#000080">開放權限修改人員(&nbsp;<font color="#ff0000">限<%=Date()%></font>&nbsp;)：</font>
                                <%
                                  If Session("user_id")<>"npois" Then
                                    SQL="Select Contribute_Open_User=user_id,user_name From userfile Where user_id<>'npois' Order By dept_id,user_group"
                                  Else
                                    SQL="Select Contribute_Open_User=user_id,user_name From userfile Order By dept_id,user_group"
                                  End If

                                  FName="Contribute_Open_User"
                                  Listfield="user_name"
                                  menusize="1"
                                  If RS("Contribute_Open_LastDate")<>"" Then
                                    If CDate(RS("Contribute_Open_LastDate"))=CDate(Date()) Then
                                      BoundColumn=RS("Contribute_Open_User")
                                    Else
                                      BoundColumn=""
                                    End If
                                  Else
                                    BoundColumn=""
                                  End If
                                  call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                                %>
                              </td>
                            </tr>
                            <%Else%>
                            <input type="hidden" name="Contribute_Close" value="">
                            <%End If%>
                            <!--#include file="../include/calendar2.asp"-->                                                                                            
                          </table>
                        </td>
                      </tr>
                      <tr>
                        <td class="td02-c" width="100%" align="center"><button id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Update_OnClick()'> <img src='../images/update.gIf' width='20' height='20' align='absmiddle'> 修改</button>&nbsp;<button id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Cancel_OnClick()'> <img src='../images/icon6.gIf' width='20' height='15' align='absmiddle'> 離開</button>&nbsp;</td>
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
function Window_OnLoad(){
  document.form.Donate_Close.focus();
}	
function Update_OnClick(){          
  if(document.form.Donate_Close.value!=''){
    <%call CheckDateJ("Donate_Close","捐款關帳日")%>
  }
  if(document.form.Contribute_Close.value!=''){
    <%call CheckDateJ("Contribute_Close","捐物資關帳日")%>
  }
  <%call SubmitJ("update")%>
}

function Cancel_OnClick(){
  location.href="donate_close.asp"
}
--></Script>