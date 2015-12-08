<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="name_list"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="name_qry.asp" target="main">
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
                    <!--20130916 Mark by GoodTV-Tanya-->
                    <!--<tr>
                      <td class="td02-c" align="right">類別：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL="Select Category=CodeDesc From CASECODE Where codetype='Category' Order By Seq"
                        FName="Category"
                        Listfield="Category"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>-->
                    <tr>
                      <td class="td02-c" align="right">身份別：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL="Select Donor_Type=CodeDesc From CASECODE Where codetype='DonorType' Order By Seq"
                        FName="Donor_Type"
                        Listfield="Donor_Type"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width="90">通訊地址：</td>
                      <td class="td02-c" width="240"><%call CodeArea ("form","City","","Area","","Y")%></td>
                      <td class="td02-c" align="right" width="90">收據地址：</td>
                      <td class="td02-c" width="370"><%call CodeArea ("form","Invoice_City","","Invoice_Area","","N")%></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款日期：</td>
                      <td class="td02-c"><%call Calendar("Donate_Date_Begin","")%> ~ <%call Calendar("Donate_Date_End","")%></td>
                      <td class="td02-c" align="right">捐款總金額：</td>
                      <td class="td02-c"><input type="text" name="Donate_Total_Begin" size="9" class="font9"> ~ <input type="text" name="Donate_Total_End" size="9" class="font9"></td>
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
                      <td class="td02-c" align="right">捐款總次數：</td>
                      <td class="td02-c"><input type="text" name="Donate_No_Begin" size="9" class="font9"> ~ <input type="text" name="Donate_No_End" size="9" class="font9"></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款人：</td>
                      <td class="td02-c"><input type="text" name="Donor_Name" size="19" class="font9"></td>
                      <td class="td02-c" align="right">捐款人編號：</td>
                      <td class="td02-c"><input type="text" name="Donor_Id_Begin" size="9" class="font9"> ~ <input type="text" name="Donor_Id_End" size="9" class="font9"></td>
                    </tr>
                    <tr>
                      <td noWrap align="right">生日月份：</td>
                      <td>
                      <%
                        Response.Write "<SELECT Name='Birthday_Month' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        Response.Write "<OPTION> </OPTION>"
                        For I=1 To 12
                          Response.Write "<OPTION value='"&I&"'>"&I&"月</OPTION>"
                        Next
                        Response.Write "</SELECT>"
                      %>	
                      </td>
                      <td noWrap align="right">多久未捐款：</td>
                      <td>
                      <%
                        Response.Write "<SELECT Name='Donate_Over' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        Response.Write "<OPTION> </OPTION>"
                        Response.Write "<OPTION value='3'>3個月</OPTION>"
                        Response.Write "<OPTION value='6'>6個月</OPTION>"
                        Response.Write "<OPTION value='9'>9個月</OPTION>"
                        Response.Write "<OPTION value='12'>一年</OPTION>"
                        Response.Write "<OPTION value='24'>二年</OPTION>"
                        Response.Write "<OPTION value='36'>三年以上</OPTION>"
                        Response.Write "</SELECT>"
                      %>
                      </td>
                    </tr>
                    <tr>
                    	<td class="td02-c" align="right">文宣品：</td> 
                      <td class="td02-c" colspan="3">
                      	<!--20130729 Modify by GoodTV Tanya:修改文宣品項目-->
                      	<input type="checkbox" name="IsSendNews" value="Y">紙本月刊
                      	<input type="checkbox" name="IsDVD" value="Y">DVD
                      	<input type="checkbox" name="IsSendEpaper" value="Y">電子文宣
                        <!--<input type="checkbox" name="IsSendNews" value="Y">會訊
                        <input type="checkbox" name="IsSendEpaper" value="Y">電子報
                        <input type="checkbox" name="IsSendYNews" value="Y">年報
                        <input type="checkbox" name="IsBirthday" value="Y">生日卡
                        <input type="checkbox" name="IsXmas" value="Y">賀卡-->
                        
                      </td>
                    </tr>
                    <tr>
                    	<td class="td02-c" align="right">其他：</td> 
                      <td class="td02-c" colspan="3">
                        <input type="checkbox" name="IsAbroad" value="Y">海外地址
                      </td>
                    </tr>
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
  if(document.form.Donate_Date_Begin.value!=''){
    <%call CheckDateJ("Donate_Date_Begin","捐款日期")%>
  }
  if(document.form.Donate_Date_End.value!=''){
    <%call CheckDateJ("Donate_Date_End","捐款日期")%>
  }
  document.form.target="main"
  document.form.action.value="query"
  document.form.submit();
  document.form.query.disabled=true
  movebar.style.display=""
}
function Report_OnClick(){
  if(document.form.Donate_Date_Begin.value!=''){
    <%call CheckDateJ("Donate_Date_Begin","捐款日期")%>
  }
  if(document.form.Donate_Date_End.value!=''){
    <%call CheckDateJ("Donate_Date_End","捐款日期")%>
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
    <%call CheckDateJ("Donate_Date_Begin","捐款日期")%>
  }
  if(document.form.Donate_Date_End.value!=''){
    <%call CheckDateJ("Donate_Date_End","捐款日期")%>
  }
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
}
--></script>