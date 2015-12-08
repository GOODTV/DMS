<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL1="Select * From DONATE_LICENCE Where Ser_No='"&request("ser_no")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  RS1("Dept_Id")=request("Dept_Id")
  RS1("Licence_No")=request("Licence_No")
  If request("Licence_BeginDate")<>"" Then
    RS1("Licence_BeginDate")=request("Licence_BeginDate")
  Else
    RS1("Licence_BeginDate")=null
  End If
  If request("Licence_EndDate")<>"" Then
    RS1("Licence_EndDate")=request("Licence_EndDate")
  Else
    RS1("Licence_EndDate")=null
  End If
  If request("Licence_Default")<>"" Then
    RS1("Licence_Default")=request("Licence_Default")
  Else
    RS1("Licence_Default")="N"
  End If
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  
  If request("Licence_Default")<>"" Then
    SQL1="Update DONATE_LICENCE Set Licence_Default='N' Where Licence_Default='Y' And Dept_Id='"&request("Dept_Id")&"' And Ser_No<>'"&request("ser_no")&"'"
    Set RS1=Conn.Execute(SQL1)
  End If
  
  session("errnumber")=1
  session("msg")="資料修改成功 ！"
End If

If request("action")="delete" Then
  SQL="Delete From DONATE_LICENCE Where Ser_No='"&request("ser_no")&"' "
  Set RS=Conn.Execute(SQL)
  session("errnumber")=1
  session("msg")="資料刪除成功 ！"
  Response.Redirect "licence.asp" 
End If

SQL="Select * From DONATE_LICENCE Where Ser_No='"&request("ser_no")&"'"
call QuerySQL(SQL,RS)
%>
<%Prog_Id="licence"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ser_no" value=<%=request("ser_no")%>>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3" size="20">
        <tr>
          <td>
            <table width="780" border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
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
                      <td class="table62-bg">&nbsp;</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;"><%=Prog_Desc%>【修改】</td>
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
	          <table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">&nbsp;</td>
                <td valign="top">
                  <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                    <tr>
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td align="right" width="12%">機構名稱：</td>
                            <td width="88%">
                            <%
                              If Session("comp_label")="1" Then
                                SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' And Dept_Id In ("&Session("all_dept_type")&") Order By Comp_Label,Dept_Id"
                              Else
                                SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id In ("&Session("all_dept_type")&") Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
                              End If
                              FName="Dept_Id"
                              Listfield="Comp_ShortName"
                              menusize="1"
                              BoundColumn=RS("Dept_Id")
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                            &nbsp;&nbsp;&nbsp;<input type="checkbox" name='Licence_Default' id="Licence_Default" value="Y" <%If RS("Licence_Default")="Y" Then%>checked<%End if%> >設定為預設勸募許可文號
                  　        </td>
                          </tr>
                          <tr>
                            <td align="right">勸募許可文號：</td>
                            <td align="left"><input type="text" name="Licence_No" size="38" class="font9" maxlength="80" value="<%=RS("Licence_No")%>" ></td>
                          </tr>
                          <tr>
                            <td align="right">使用期間：</td>
                            <td align="left"><%call Calendar("Licence_BeginDate",RS("Licence_BeginDate"))%> ~ <%call Calendar("Licence_EndDate",RS("Licence_EndDate"))%></td>
                          </tr>
                          <!--#include file="../include/calendar2.asp"-->
                          <tr>
                            <td align="center" colspan="8">
                               <input type="button" value=" 修 改 " name="save" class="cbutton" style="cursor:hand" onClick="Update_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value=" 刪 除 " name="save" class="cbutton" style="cursor:hand" onClick="Delete_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value=" 取 消 " name="cancel" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">
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
    </form>
  </center></div></p>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Window_OnLoad(){
  document.form.Licence_No.focus();
}	
function Update_OnClick(){
  <%call CheckStringJ("Licence_No","勸募許可文號")%>
  <%call ChecklenJ("Licence_No",80,"勸募許可文號")%>
  <%call CheckStringJ("Licence_BeginDate","使用起日")%>
  if(document.form.Licence_EndDate.value==''){
    document.form.Licence_EndDate.value='2099/12/31';
  }
  <%call DiffDateJ("Licence_BeginDate","Licence_EndDate")%>
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  location.href='licence.asp';
}
--></script>