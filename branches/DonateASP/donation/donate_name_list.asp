<!--#include file="../include/dbfunctionJ.asp"-->
<%
  Invoice_Pre=""
  SQL1="Select Invoice_Pre From DEPT Where Dept_Id='"&Session("dept_id")&"'"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then  Invoice_Pre=RS1("Invoice_Pre")
  RS1.Close
  Set RS1=Nothing
%>
<%Prog_Id="donate_name_list"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="donate_name_list_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">捐款名冊</td>
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
                      <td class="td02-c" align="right">捐款機構：</td>
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
                      <td class="td02-c" align="right">類別：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL="Select Category=CodeDesc From CASECODE Where CodeType='Category' Order By Seq"
                        FName="Category"
                        Listfield="Category"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">身份別：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL="Select Donor_Type=CodeDesc From CASECODE Where CodeType='DonorType' Order By Seq"
                        FName="Donor_Type"
                        Listfield="Donor_Type"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width="100">通訊地址：</td>
                      <td class="td02-c" width="230"><%call CodeArea ("form","City","","Area","","Y")%></td>
                      <td class="td02-c" align="right" width="90">收據地址：</td>
                      <td class="td02-c" width="340"><%call CodeArea ("form","Invoice_City","","Invoice_Area","","N")%></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款日期：</td>
                      <td class="td02-c"><%call Calendar("Donate_Date_Begin","")%> ~ <%call Calendar("Donate_Date_End","")%></td>
                      <td class="td02-c" align="right">單筆金額：</td>
                      <td class="td02-c"><input type="text" name="Donate_Amt_Begin" size="9" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Export_OnClick();}'> ~ <input type="text" name="Donate_Amt_End" size="9" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Export_OnClick();}'></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">募款活動：</td>
                      <td class="td02-c">
                      <%
                        SQL="Select Act_Id,Act_ShortName From ACT Order By Act_id Desc"
                        FName="Act_Id"
                        Listfield="Act_ShortName"
                        menusize="1"
                        BoundColumn=""
                        call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                      %>
                      </td>
                      <td class="td02-c" align="right">收據編號：</td>
                      <td class="td02-c"><input type="text" name="Invoice_No_Begin" size="9" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Export_OnClick();}'> ~ <input type="text" name="Invoice_No_End" size="9" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Export_OnClick();}'><%If Invoice_Pre<>"" Then%>&nbsp;<font color="#ff0000">(&nbsp;機構前置碼不需輸入&nbsp;)</font><%End If%></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款人：</td>
                      <td class="td02-c"><input type="text" name="Donor_Name" size="19" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Export_OnClick();}'></td>
                      <td class="td02-c" align="right"> </td>
                      <td class="td02-c"> </td>
                    </tr> 
                    <tr>
                      <td class="td02-c" align="right">捐款方式：</td>
                      <td class="td02-c" colspan="3" width="660">
                      <%
                        SQL="Select Donate_Payment=CodeDesc From CASECODE Where CodeType='Payment' Order By Seq"
                        FName="Donate_Payment"
                        Listfield="Donate_Payment"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款用途：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL="Select Donate_Purpose=CodeDesc From CASECODE Where CodeType='Purpose' Order By Seq"
                        FName="Donate_Purpose"
                        Listfield="Donate_Purpose"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款類別：</td>
                      <td class="td02-c" colspan="3">
                        <input type="checkbox" name="Donate_Type" id="Donate_Type1" value="單次捐款">單次捐款
                        <input type="checkbox" name="Donate_Type" id="Donate_Type2"  value="長期捐款">長期捐款
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">入帳銀行：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL="Select Accoun_Bank=CodeDesc From CASECODE Where CodeType='Bank' Order By Seq"
                        FName="Accoun_Bank"
                        Listfield="Accoun_Bank"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">沖帳日期：</td>
                      <td class="td02-c" colspan="3"><%call Calendar("Accoun_Date_Begin","")%> ~ <%call Calendar("Accoun_Date_End","")%></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">會計科目：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL="Select Accounting_Title=CodeDesc From CASECODE Where CodeType='Accoun' Order By Seq"
                        FName="Accounting_Title"
                        Listfield="Accounting_Title"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">其他：</td>
                      <td class="td02-c" colspan="3">
                        <input type="checkbox" name="Issue_TypeM" value="M">手開收據
                        <input type="checkbox" name="Issue_TypeD" value="D">作廢收據
                      </td>
                    </tr>
                    <%
                      SQL_IS="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"' "
                      Set RS_IS=Server.CreateObject("ADODB.Recordset")
                      RS_IS.Open SQL_IS,Conn,1,1
                      IsDonation=RS_IS("IsDonation")
                      IsMember=RS_IS("IsMember")
                      IsShopping=RS_IS("IsShopping")
                      RS_IS.Close
                      Set RS_IS=Nothing
                      If IsMember="Y" Or IsShopping="Y" Then
                    %>
                    <tr>
                      <td class="td02-c" align="right">款項分類：</td>
                      <td class="td02-c" colspan="3">
                        <input type='checkbox' name='Donate_Purpose_Type' id="Donate_Purpose_Type1" value="D" checked >愛心捐款
                        <%If IsMember="Y" Then%><input type='checkbox' name='Donate_Purpose_Type' id="Donate_Purpose_Type2" value="M">會務繳費<%End If%>
                        <%If IsMember="Y" Then%><input type='checkbox' name='Donate_Purpose_Type' id="Donate_Purpose_Type3" value="A">活動報名<%End If%>
                        <%If IsShopping="Y" Then%><input type='checkbox' name='Donate_Purpose_Type' id="Donate_Purpose_Type4" value="S">商品義賣<%End If%>
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
                        Response.Write "<button id='query' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Query_OnClick()'><img src='../images/search.gif' width='20' height='20'><br>查詢</button>&nbsp;"
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
<script language="JavaScript"><!--
function Window_OnLoad(){
  document.form.Donate_Date_Begin.focus();
}	
function Query_OnClick(){
  document.form.target="main"
  document.form.action.value="query"
  document.form.submit();
  document.form.query.disabled=true
  movebar.style.display=""
}
function Report_OnClick(){
  if(confirm('您是否確定要將查詢結果列表？')){
    document.form.target="report"
    document.form.action.value="report"
    window.open('','report','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600')
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