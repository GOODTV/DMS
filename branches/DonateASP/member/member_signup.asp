<!--#include file="../include/dbfunctionJ.asp"-->
<%
  SQL1="Select Member_Act_Name,Member_Act_IsPrice,Member_Act_IsFood From MEMBER_ACT Where Member_Act_Id='"&request("member_act_id")&"'"
  call QuerySQL(SQL1,RS1)
  Member_Act_Name=RS1("Member_Act_Name")
  Member_Act_IsPrice=RS1("Member_Act_IsPrice")
  Member_Act_IsFood=RS1("Member_Act_IsFood")
  RS1.Close
  Set RS1=Nothing
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>活動報名清單</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <form name="form" method="post" action="member_signup_list.asp" target="detail">
      <input type="hidden" name="action">
      <input type="hidden" name="member_act_id" value="<%=request("member_act_id")%>">
      <input type="hidden" name="Member_Act_IsPrice" value="<%=Member_Act_IsPrice%>">
      <input type="hidden" name="Member_Act_IsFood" value="<%=Member_Act_IsFood%>">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="780"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
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
                      <td class="table62-bg">　</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">活動報名清單</td>
                      <td class="table63-bg">　</td>
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
                <td class="table62-bg">　</td>
                <td valign="top">
                  <table width="100%" border="0" cellpadding="1" cellspacing="1">
                    <tr>
                      <td class="td02-c" align="right">活動名稱：</td>
                      <td class="td02-c"><input type="text" name="Member_Act_Name" size="94" class="font9t" readonly value="<%=Member_Act_Name%>" ></td>
                      <td class="td02-c">
                        <input type="button" value="回列表" name="exit" class="cbutton" style="cursor:hand" onClick="Exit_OnClick()">	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" width="70" align="right">報名日期：</td>
                      <td class="td02-c" width="650">
                        <%call Calendar("Member_Signup_Date_B","")%> ~ <%call Calendar("Member_Signup_Date_E","")%>
                      	&nbsp;
                      	報名狀態：
                        <%
                          If Member_Act_IsPrice="N" Then
                            SQL="Select Member_Signup_Status=CodeDesc From CASECODE Where codetype='SignupStatus' And CodeDesc Not In ('已繳費','已退費') Order By Seq"
                          Else
                            SQL="Select Member_Signup_Status=CodeDesc From CASECODE Where codetype='SignupStatus' Order By Seq"
                          End If
                          FName="Member_Signup_Status"
                          Listfield="Member_Signup_Status"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                      	&nbsp;
                      	姓名：
                      	<input type="text" name="KeyWords" size="30" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                      </td>
                      <td class="td02-c" width="80">
                        <input type="button" value=" 查  詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right"> </td>
                      <td class="td02-c"> </td>
                      <td class="td02-c">
                        <input type="button" value=" 匯  出 " name="export" class="cbutton" style="cursor:hand" onClick="Export_OnClick()">	
                      </td>
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td class="td02-c" colspan="9" width="100%"><iframe name="detail" src="member_signup_list.asp?member_act_id=<%=request("member_act_id")%>" height="380" width="100%" frameborder="0" scrolling="auto"></iframe></td>
                    </tr>
                  </table>
                </td>
                <td class="table63-bg">　</td>
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
function Exit_OnClick(){
	location.href='member_act.asp';
}
function Query_OnClick(){ 
  document.form.target="detail"
  document.form.action.value="query"
  document.form.submit();
}
function Export_OnClick(){
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
}
--></script>