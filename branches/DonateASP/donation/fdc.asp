<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Session("DeptId")=request("Dept_Id")
  Session("Donate_Year")=request("Donate_Year")
  Session("Act_Id")=request("Act_Id")
  Session("Uniform_No")=request("Uniform_No")
  Session("Licence")=request("Licence")
  Server.Transfer "fdc_export.asp"
End If

SQL1="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"'"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then
  Licence=RS1("Licence")
  Uniform_No=RS1("Uniform_No")
  Rept_Licence=RS1("Rept_Licence")
End If
RS1.Close
Set RS1=Nothing
%>
<%Prog_Id="fdc"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <table width="700" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="100%"  border="0" cellspacing="0" cellpadding="0" align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"></td>
                <td width="95%">
  		            <table width="60%"  border="0" cellspacing="0" cellpadding="0">
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
	          <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">&nbsp;</td>
                <td valign="top">
                  <table width="100%" border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111">
                    <tr>
                      <td class="td02-c" align="right" width="80">機構：</td>
                      <td class="td02-c" width="120">
                        <%
                          If Session("comp_label")="1" Then
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' And Dept_Id In ("&Session("all_dept_type")&") Order By Comp_Label,Dept_Id"
                          Else
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id In ("&Session("all_dept_type")&") Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
                          End If
                          FName="Dept_Id"
                          Listfield="Comp_ShortName"
                          menusize="1"
                          BoundColumn=""
                          call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                        %>
                      </td>
                      <td class="td02-c" align="right" width="80">捐款年度：</td>
                      <td class="td02-c" width="80">
                      <%
                        SQL1="Select Begin_Year=Isnull(Min(Year(Donate_Date)),"&Year(Date())&") From DONATE"
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        If Not RS1.EOF Then Begin_Year=RS1("Begin_Year")
                        RS1.Close
                        Set RS1=Nothing
                        Response.Write "<SELECT Name='Donate_Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        For I= Cint(Year(Date())) To Cint(Begin_Year) Step -1
                          If I=Cint(Year(Date())-1) Then
                            Response.Write "<OPTION selected value='"&I&"'>"&I&"</OPTION>"
                          Else
                            Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
                          End If
                        Next
                        Response.Write "</SELECT>"
                      %>	
                      </td>
                      <td class="td02-c" align="right" width="80">勸募活動：</td>
                      <td class="td02-c" width="240">
                      <%
                        SQL="Select Act_Id,Act_ShortName From ACT Order By Act_id Desc"
                        FName="Act_Id"
                        Listfield="Act_ShortName"
                        menusize="1"
                        BoundColumn=""
                        call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                      %>	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">機構統編：</td>
                      <td class="td02-c"><input type="text" name="Uniform_No" size="16" class="font9" maxlength="8" value="<%=Uniform_No%>" onKeypress='if(event.keyCode==13){JavaScript:Export_OnClick();}'></td>
                      <td class="td02-c" align="right">許可文號：</td>
                      <td class="td02-c" colspan="3"><input type="text" name="Licence" size="30" class="font9" maxlength="60" value="<%=Licence%>" onKeypress='if(event.keyCode==13){JavaScript:Export_OnClick();}'></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">※注意事項：</td>
                      <td class="td02-c" colspan="5">
                      	1.本單據適用對象僅限於收據身分證(統一編號)資料已填寫者。<br />
                      	2.【許可文號】泛指專案核准文號、註管機關登記、立案文號(如有向法院辦理登記應加註法院登記簿<br />之&nbsp;&nbsp;&nbsp;冊、頁、號)
                      	或財團法人設立許可文號。<br />
                      	3.匯出後請將檔案另存成國稅局標準檔名<b>【捐款年度民國年.31.統一編號.txt】</b>
                      </td>
                    </tr>
                    <tr>
                      <td width="100%" colspan="6" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="6" width="100%"><%Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick()'><img src='../images/icon_TXT.gif' width='16' height='16'><br>匯出</button>"%></td>
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
  document.form.Donate_Year.focus();
}	
function Export_OnClick(){
  <%call CheckStringJ("Dept_Id","機構")%>
  <%call CheckStringJ("Donate_Year","捐款年度")%>
  <%call CheckStringJ("Uniform_No","機構統編")%>
  <%call CheckNumberJ("Uniform_No","機構統編")%>
  <%call ChecklenJ("Uniform_No",8,"機構統編")%>
  <%call ChecklenJ("Licence",60,"許可文號")%>
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();s
  }
}
--></script>