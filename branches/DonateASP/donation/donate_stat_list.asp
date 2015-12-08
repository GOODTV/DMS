<!--#include file="../include/dbfunctionJ.asp"-->
<%
  SQL_IS="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"' "
  Set RS_IS=Server.CreateObject("ADODB.Recordset")
  RS_IS.Open SQL_IS,Conn,1,1
  IsMember=RS_IS("IsMember")
  IsShopping=RS_IS("IsShopping")
  RS_IS.Close
  Set RS_IS=Nothing

  Donate_Date_Begin=DateAdd("M",-1,Year(Date())&"/"&Month(Date())&"/1")
  Donate_Date_End=DateAdd("D",-1,Year(Date())&"/"&Month(Date())&"/1")
  
  Donate_Amt_Kind="1000,3000,5000,10000,30000,50000,100000"
  SQL1="Select Donate_Amt_Kind From DEPT Where Dept_Id='"&Session("dept_id")&"'"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then
    If RS1("Donate_Amt_Kind")<>"" Then Donate_Amt_Kind=RS1("Donate_Amt_Kind")
  End If
  RS1.Close
  Set RS1=Nothing  
%>
<%Prog_Id="stat"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="donate_stat_list_qry.asp" target="main">
      <input type="hidden" name="action">
      <input type="hidden" name="DonateDesc" value="<%=DonateDesc%>">	
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
                      <td class="td02-c" align="right">機構：</td>
                      <td class="td02-c" colspan="3">
                        <%
                          If Session("comp_label")="1" Then
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' And Dept_Id In ("&Session("all_dept_type")&") Order By Comp_Label,Dept_Id"
                          Else
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id In ("&Session("all_dept_type")&") Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
                          End If
                          FName="Dept_Id"
                          Listfield="Comp_ShortName"
                          menusize="1"
                          BoundColumn=Session("dept_id")
                          call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                        %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width="90">捐款日期：</td>
                      <td class="td02-c" width="220"><%call Calendar("Donate_Date_Begin",Donate_Date_Begin)%> ~ <%call Calendar("Donate_Date_End",Donate_Date_End)%></td>
                      <td class="td02-c" align="right" align="right" width="75">募款活動：</td>
                      <td class="td02-c" width="375">
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
                      <td class="td02-c" align="right">統計項目：</td>
                      <td class="td02-c" colspan="3">
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind1" value="Category">類別&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind2" value="Sex">性別&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind3" value="Age">年齡&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind4" value="Education">教育程度&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind5" value="Occupation">職業別&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind6" value="Marriage">婚姻狀況&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind7" value="Religion">宗教信仰&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind8" value="City">通訊縣市&nbsp;
                      	<input type="radio" name="Stat_Kind" id="Stat_Kind9" value="Invoice_City">收據縣市<br />
                        <input type="radio" name="Stat_Kind" id="Stat_Kind10" value="Donate_Payment">捐款方式&nbsp;
                        <input type="radio" name="Stat_Kind" id="Stat_Kind11" value="Donate_Purpose">捐款用途&nbsp;
                        <input type="radio" name="Stat_Kind" id="Stat_Kind12" value="Donate_Amt" checked >捐款金額&nbsp;
                        <input type="text" name="Donate_Amt_Kind" size="57" class="font9" maxlength="100" value="<%=Donate_Amt_Kind%>">
                      </td>
                    </tr>
                    <%
                      If IsMember="Y" Or IsShopping="Y" Then
                    %>
                    <tr>
                      <td class="td02-c" align="right">款項分類：</td>
                      <td class="td02-c" colspan="3">
                        <input type='checkbox' name='Donate_Purpose_Type' id="Donate_Purpose_Type1" value="D" checked >愛心捐款
                        <%If IsMember="Y" Then%><input type='checkbox' name='Donate_Purpose_Type' id="Donate_Purpose_Type2" value="M" checked >會務繳費<%End If%>
                        <%If IsMember="Y" Then%><input type='checkbox' name='Donate_Purpose_Type' id="Donate_Purpose_Type3" value="A" checked >活動報名<%End If%>
                        <%If IsShopping="Y" Then%><input type='checkbox' name='Donate_Purpose_Type' id="Donate_Purpose_Type4" value="S" checked >商品義賣<%End If%>
                      </td>
                    </tr>
                    <%
                      Else
                        Response.Write "<input type=""hidden"" name=""Donate_Purpose_Type"" value=""D"">"
                      End If
                    %>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick()'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick()'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"
    			              Response.Write "<span id='movebar' style='display:none'><img border='0' src='../images/movebar.gif'></span>"
    			            %>
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
<script language="JavaScript"><!--\
function Window_OnLoad(){
  document.form.Donate_Date_Begin.focus();
}	
function Report_OnClick(){
  if(document.form.Donate_Date_Begin.value!=''){
    <%call CheckDateJ("Donate_Date_Begin","捐款日期起")%>
  }
  if(document.form.Donate_Date_End.value!=''){
    <%call CheckDateJ("Donate_Date_End","捐款日期迄")%>
  }
  if(document.form.Stat_Kind[11].checked){
    <%call CheckStringJ("Donate_Amt_Kind","捐款金額區間")%>
  }
  if(confirm('您是否確定要將查詢結果列表？')){
    document.form.target="report"
    document.form.action.value="report"
    window.open('','report','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600')
    document.form.submit();
  }
}

function Export_OnClick(){
  if(document.form.Donate_Date_Begin.value!=''){
    <%call CheckDateJ("Donate_Date_Begin","捐款日期起")%>
  }
  if(document.form.Donate_Date_End.value!=''){
    <%call CheckDateJ("Donate_Date_End","捐款日期迄")%>
  }
  if(document.form.Stat_Kind[11].checked){
    <%call CheckStringJ("Donate_Amt_Kind","捐款金額區間")%>
  }
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
}
--></script>