<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL_IS="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"' "
Set RS_IS=Server.CreateObject("ADODB.Recordset")
RS_IS.Open SQL_IS,Conn,1,1
IsMember=RS_IS("IsMember")
IsShopping=RS_IS("IsShopping")
RS_IS.Close
Set RS_IS=Nothing
%>
<%Prog_Id="iepay"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="iepay_list.asp" target="detail">
      <input type="hidden" name="action">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="900"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
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
	          <table width="900" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">&nbsp;</td>
                <td valign="top">
                  <table width="100%" border="0" cellpadding="1" cellspacing="1">
                    <tr>
					  <td>
                        授權狀態：
                        <select size="1" name="Status" class="font9">
                          <option value="">  </option>
						  <option value="0">授權成功</option>
                          <option value="1">授權失敗</option>
                          <option value="2">未完成刷卡流程</option>
                        </select>
						&nbsp;
					  </td>
					  <td>
						捐款日期：
						<%call Calendar("Donate_Date_B",Now())%> ~ <%call Calendar("Donate_Date_E",Now())%>
					  </td>
					  <td class="td02-c">
                        <input type="button" value=" 查 詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">	
                      </td>
					  <td class="td02-c">
						<input type="button" value=" 列 表 " name="report" class="cbutton" style="cursor:hand" onClick="Report_OnClick()">
					  </td>
                      <td class="td02-c">
                        <input type="button" value=" 匯 出 " name="report" class="cbutton" style="cursor:hand" onClick="Export_OnClick()">
                      </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <%
                        Donate_Date_B=Year(Date())&"/1/1"
                        CreateUser=""
                        If Session("user_group")<>"系統管理" And Session("user_group")<>"捐款管理" Then CreateUser=Session("user_name")
                      %>
                      <td class="td02-c" colspan="9" width="100%"><iframe name="detail" src="iepay_list.asp?Dept_Id=<%=Session("dept_id")%>&Donate_Date_B=<%=Donate_Date_B%>&Create_User=<%=CreateUser%>&Donate_Purpose_Type=D" height="430" width="100%" frameborder="0" scrolling="auto"></iframe></td>
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
  document.form.Status.focus();
}	
function Query_OnClick(){
  if(document.form.Donate_Date_B.value!=''){
    <%call CheckDateJ("Donate_Date_B","捐款日期起")%>
  }
  if(document.form.Donate_Date_E.value!=''){
    <%call CheckDateJ("Donate_Date_E","捐款日期迄")%>
  }
  document.form.target="detail"
  document.form.action.value="query"
  document.form.submit();
}
function Report_OnClick(){
  if(document.form.Donate_Date_B.value!=''){
    <%call CheckDateJ("Donate_Date_B","捐款日期起")%>
  }
  if(document.form.Donate_Date_E.value!=''){
    <%call CheckDateJ("Donate_Date_E","捐款日期迄")%>
  }
  if(confirm('您是否確定要將查詢結果列表？')){
    document.form.target="report"
    document.form.action.value="report"
    window.open('','report','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=800,height=600')
    document.form.submit();
  }
}

function Export_OnClick(){
  if(document.form.Donate_Date_B.value!=''){
    <%call CheckDateJ("Donate_Date_B","捐款日期起")%>
  }
  if(document.form.Donate_Date_E.value!=''){
    <%call CheckDateJ("Donate_Date_E","捐款日期迄")%>
  }
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
}
--></script>