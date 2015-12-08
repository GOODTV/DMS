<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL1="Select * From DONATE_SALE Where ser_no='"&request("ser_no")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  RS1("Dept_Id")=request("Dept_Id")
  RS1("Sale_Subject")=request("Sale_Subject")
  RS1("Sale_BeginDate")=request("Sale_BeginDate")
  RS1("Sale_EndDate")=request("Sale_EndDate")
  RS1("Sale_Content")=request("Sale_Content")
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 ！"
End If

If request("action")="delete" Then
  SQL="Delete From DONATE_SALE Where Ser_No='"&request("ser_no")&"'" 
  Set RS=Conn.Execute(SQL)    
  session("errnumber")=1
  session("msg")="資料刪除成功 !"
  Response.Redirect "donate_sale.asp"
End If

SQL="Select * From DONATE_SALE Where Ser_No='"&request("Ser_No")&"'"
Call QuerySQL(SQL,RS)
%>
<%Prog_Id="donate_sale"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
    	<input type="hidden" name="action">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="780" border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"> </td>
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
                        <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td align="right" width="65">機構：</td>
                            <td align="left" width="715">
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
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                            %>
                            &nbsp;&nbsp;&nbsp;文宣標題：
                            <input class=font9 size=30 name="Sale_Subject" maxlength="50" value="<%=RS("Sale_Subject")%>">
                            &nbsp;&nbsp;&nbsp;列印日期：
                            <%call Calendar("Sale_BeginDate",RS("Sale_BeginDate"))%> ~ <%call Calendar("Sale_EndDate",RS("Sale_EndDate"))%>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">文宣內容：</td>
                            <td align="left"><textarea rows="6" name="Sale_Content" cols="93" class="font9"><%=RS("Sale_Content")%></textarea></td>
                          </tr>
                          <!--#include file="../include/calendar2.asp"-->
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="2"> </td>
                          </tr>
                          <tr>
                            <td align="center" colspan="2">
                              <input type="button" value=" 存 檔 " name="save" class="cbutton" style="cursor:hand" onClick="Update_OnClick()">
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
  document.form.Sale_Subject.focus();
}			
function Update_OnClick(){
  <%call CheckStringJ("Dept_Id","機構")%>
  <%call CheckStringJ("Sale_Subject","文宣標題")%>
  <%call ChecklenJ("Sale_Subject",50,"文宣標題")%>
  <%call CheckStringJ("Sale_BeginDate","列印起日")%>
  <%call CheckDateJ("Sale_BeginDate","列印起日")%>
  if(document.form.Sale_EndDate.value==''){
    document.form.Sale_EndDate.value='2099/12/31';
  }
  <%call DiffDateJ("Sale_BeginDate","Sale_EndDate")%>
  <%call CheckDateJ("Sale_EndDate","列印迄日")%>
  <%call CheckStringJ("Sale_Content","文宣內容")%>
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  location.href='donate_sale.asp';
}
--></script>