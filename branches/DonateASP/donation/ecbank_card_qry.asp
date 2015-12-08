<!--#include file="../include/dbfunctionJ.asp"-->
<%
  If Request("Donate_CreateDate_B")="" Then
    If Month(Date())=1 Then
      Donate_CreateDate_B=Cstr(Year(Date())-1)&"/12/1"
    Else
      Donate_CreateDate_B=Cstr(Year(Date()))&"/"&Cstr(Month(Date())-1)&"/1"
    End If
  Else
    Donate_CreateDate_B=Request("Donate_CreateDate_B")
  End If
%>
<%Prog_Id="ecbank_card_qry"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="ecbank_card_qry_list.asp" target="detail">
      <input type="hidden" name="action">
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
	          <table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
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
                      <td class="td02-c" width="70" align="right">機構：</td>
                      <td class="td02-c" width="650">
                        <%
                          If Session("comp_label")="1" Then
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' Order By Comp_Label,Dept_Id"
                          Else
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id='"&Session("dept_id")&"' Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
                          End If
                          FName="Dept_Id"
                          Listfield="Comp_ShortName"
                          menusize="1"
                          BoundColumn=Session("dept_id")
                          call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                        %>
                      	&nbsp;
                      	捐款人：
                      	<input type="text" name="Donate_DonorName" size="13" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                        &nbsp;
                        捐款用途：
                        <%
                          SQL="Select Donate_Purpose=CodeDesc From CASECODE Where codetype='Purpose' Order By Seq"
                          FName="Donate_Purpose"
                          Listfield="Donate_Purpose"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        &nbsp;
                        收據開立：
                        <%
                          SQL="Select Donate_Invoice_Type=CodeDesc From CASECODE Where codetype='InvoiceType' Order By Seq"
                          FName="Donate_Invoice_Type"
                          Listfield="Donate_Invoice_Type"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                      </td>
                      <td class="td02-c" width="80">
                        <input type="button" value=" 查 詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">交易日期：</td>
                      <td class="td02-c">
                      <%call Calendar("Donate_CreateDate_B",Donate_CreateDate_B)%> ~ <%call Calendar("Donate_CreateDate_E","")%>
                      &nbsp;
                      交易狀態：
                      <SELECT Name='Close_Type' size='1' style='font-size: 9pt; font-family: 新細明體'>
                      	<OPTION value=''> </OPTION>
                      	<OPTION value='L'>失敗</OPTION>
                      	<OPTION value='S'>成功</OPTION>
                      	<OPTION value='P'>已請款</OPTION>
                      </SELECT>
                      </td>
                      <td class="td02-c">
                        <input type="button" value=" 匯 出 " name="report" class="cbutton" style="cursor:hand" onClick="Export_OnClick()">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">是否匯出：</td>
                      <td class="td02-c">
                        <SELECT Name='Donate_Export' size='1' style='font-size: 9pt; font-family: 新細明體'>
                      	  <OPTION value=''> </OPTION>
                      	  <OPTION value='N' selected >未匯出</OPTION>
                      	  <OPTION value='Y'>已匯出</OPTION>
                        </SELECT>
                        &nbsp;
                        <font color="#000080">(以下交易資料如有錯誤，以綠界科技發布紀錄為準。)</font>
                      </td>
                      <td class="td02-c"> </td>
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td class="td02-c" colspan="9" width="100%"><iframe name="detail" src="ecbank_card_qry_list.asp?Donate_Export=N&Donate_CreateDate_B=<%=Donate_CreateDate_B%>" height="380" width="100%" frameborder="0" scrolling="auto"></iframe></td>
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
  document.form.Donate_DonorName.focus();
}	
function Query_OnClick(){
  document.form.target="detail"
  document.form.action.value="query"
  document.form.submit();
}
function Report_OnClick(){
  if(confirm('您是否確定要將查詢結果列表？')){
    document.form.target="report"
    document.form.action.value="report"
    window.open('','report','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=800,height=600')
    document.form.submit();
  }
}

function Export_OnClick(){
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
}
--></script>