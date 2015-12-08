<!--#include file="../include/dbfunctionJ.asp"-->
<%
year_now=year(Date)
month_now=month(Date)
//Donate_Date_Begin=DateAdd("M",-1,Year(Date())&"/"&Month(Date())&"/1")
//Donate_Date_End=DateAdd("D",-1,Year(Date())&"/"&Month(Date())&"/1")
Accumulate_Date_Begin=Year(Date())&"/1/1"
Accumulate_Date_End=Year(Date())&"/12/31"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title><%=request("subject")%>報表清單</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<table width="750" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3" align="center">
	<tr>
		<td><b><font color="#8B0000" size="3">週報</font></b>
		</td>
		<td><b><font color="#8B0000" size="3">月報</font></b>
		</td>
		<td><b><font color="#8B0000" size="3">季報</font></b>
		</td>
		<td><b><font color="#8B0000" size="3">年報</font></b>
		</td>
		<td><b><font color="#8B0000" size="3">其他</font></b>
		</td>
	</tr>
	<tr>
		<td height="20"><font size="2">建台/非建台奉獻級距表</font>
		</td>
		<td><font size="2">大額捐款人生日查詢表</font>
		</td>
		<td><font size="2">非建台奉獻統計分析表</font>
		</td>
		<td><font size="2">捐款總額與月刊索取明細表</font>
		</td>
		<td><font size="2">人數統計查詢報表</font>
		</td>
	</tr>
	<tr>
		<td height="20"><font size="2">新捐款人明細表</font>
		</td>
		<td><font size="2">定期奉獻授權到期明細表</font>
		</td>
		<td><font size="2">查詢單筆捐款金額累計</font>
		</td>
		<td><font size="2">海外捐款總額明細表</font>
		</td>
		<td><font size="2">捐款用途總額明細表</font>
		</td>
	</tr>
	<tr>
		<td height="20"><font size="2">查詢單筆捐款金額</font>
		</td>
		<td><font size="2">郵寄月刊明細表</font>
		</td>
		<td>
		</td>
		<td><font size="2">國內地區捐款總額明細表</font>
		</td>
		<td><font size="2">捐款總額明細表</font>
		</td>
	</tr>
	<tr>
		<td height="20">
		</td>
		<td><font size="2">郵寄電子報明細表</font>
		</td>
		<td>
		</td>
		<td><font size="2">期間身份別捐款總額明細表</font>
		</td>
		<td><font size="2">DVD贈品索取人報表</font>
		</td>
	</tr>
	<tr>
		<td height="20">
		</td>
		<td><font size="2">當月各日捐款方式統計表</font>
		</td>
		<td>
		</td>
		<td><font size="2">期間單筆捐款金額次數明細表</font>
		</td>
		<td><font size="2">雷同資料查詢</font>
		</td>
	</tr>
	<tr>
		<td height="20">
		</td>
		<td>
		</td>
		<td>
		</td>
		<td><font size="2">新增捐款人總額明細表</font>
		</td>
		<td>
		</td>
	</tr>
	<tr>
		<td height="20">
		</td>
		<td>
		</td>
		<td>
		</td>
		<td><font size="2">經常奉獻總額明細表</font>
		</td>
		<td>
		</td>
	</tr>
	<tr>
		<td height="20">
		</td>
		<td>
		</td>
		<td>
		</td>
		<td><font size="2">歷年來大陸奉獻天使累計奉獻明細表</font>
		</td>
		<td>
		</td>
	</tr>
	<tr>
		<td height="20">
		</td>
		<td>
		</td>
		<td>
		</td>
		<td>
		</td>
		<td>
		</td>
	</tr>
</table>
<table width="750" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3" align="center">
	<tr>
		<td>
			<form name="form"> 
				<select name="select" size="1" onchange="setChange()"> 
					<option value="">------------------------週報------------------------</option>
					<option value="week1">[週報]建台/非建台奉獻級距表</option>  
					<option value="week2">[週報]新捐款人明細表</option>
					<option value="week3">[週報]查詢單筆捐款金額</option>
					<option value="">------------------------月報------------------------</option>
					<option value="month1">[月報]大額捐款人生日查詢表</option>  
					<option value="month2">[月報]定期奉獻授權到期明細表</option> 
					<option value="month3">[月報]郵寄月刊明細表</option> 
					<option value="month4">[月報]郵寄電子報明細表</option>
					<option value="month5">[月報]當月各日捐款方式統計表</option>
					<option value="">------------------------季報------------------------</option>
					<option value="season1">[季報]非建台奉獻統計分析表</option>  
					<option value="season2">[季報]查詢單筆捐款金額累計</option> 
					<!--option value="season3">[季報]捐款用途各項總額明細表</option-->
					<option value="">------------------------年報------------------------</option>
					<option value="year2">[年報]捐款總額與月刊索取明細表</option> 
					<option value="year3">[年報]海外捐款總額明細表</option>
					<option value="year4">[年報]國內地區捐款總額明細表</option>  
					<option value="year5">[年報]期間身份別捐款總額明細表</option> 
					<option value="year6">[年報]期間單筆捐款金額次數明細表</option>
					<option value="year7">[年報]新增捐款人總額明細表</option>  
					<option value="year8">[年報]經常奉獻總額明細表</option> 
					<option value="year9">[年報]歷年來大陸奉獻天使累計奉獻明細表</option>
					<option value="">------------------------其他------------------------</option>
					<option value="other1">[其他]人數統計查詢報表</option>
					<option value="year1">[其他]捐款用途總額明細表</option>
					<option value="season4">[其他]捐款總額明細表</option>
					<option value="other2">[其他]DVD贈品索取人報表</option>
					<option value="other3">[其他]雷同資料查詢</option>
				</select> 
			</form>
		</td>
	</tr>
</table>
<div id="tb_week1" style="display:none" align="center">
	<form name="form_Week1" method="post" action="donate_week_report1_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">建台/非建台奉獻級距表</td>
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
                      <td class="td02-c" align="right">捐款用途：</td>
                      <td class="td02-c" colspan="3"><select name="Donate_Purpose">
										　<option value="建台">建台</option>
										　<option value="非建台">非建台</option>
										</select></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款日期：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Week1")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Week1")%></td>
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_week2" style="display:none" align="center">
<form name="form_Week2" method="post" action="donate_week_report3_qry.asp" target="main">
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
				  <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">新捐款人明細表</td>
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
				  <td class="td02-c" align="right"><font color="red">*</font>捐款日期：</td>
				  <td class="td02-c" width="210"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Week2")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Week2")%></td>
				  <td class="td02-c" align="right" width="90">捐款金額起訖：</td>
				  <td class="td02-c"><input type="text" name="Donate_Total_Begin" size="12" class="font9"> ~ <input type="text" name="Donate_Total_End" size="12" class="font9"></td>
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
				  <td class="td02-c" align="right">不含：</td>
				  <td class="td02-c" colspan="3">
					<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
					<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
					<input type="checkbox" name="Sex" value="D" checked>歿
				  </td>
                </tr>
				<!--#include file="../include/calendar2.asp"-->
				<tr>
				  <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
				</tr>
				<tr>
				  <td colspan="8" width="100%">
				  <%
						  Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
						  Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_week3" style="display:none" align="center">
<form name="form_Week3" method="post" action="donate_week_report2_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">查詢單筆捐款金額</td>
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
                      <td class="td02-c" align="right">捐款金額(大於)：</td>
                      <td class="td02-c" colspan="3"><input type="text" name="Donate_Amt" size="12" class="font9">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款期間：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Week3")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Week3")%></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">累計期間：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Accumulate_Date_Begin",Accumulate_Date_Begin,"form_Week3")%> ~ <%call Calendar2("Accumulate_Date_End",Accumulate_Date_End,"form_Week3")%></td>
                    </tr>
					<tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_month1" style="display:none" align="center">
    <form name="form_Month1" method="post" action="donate_month_report1_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">大額捐款人生日查詢表</td>
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
					  <td class="td02-c" width="130" align="right">捐款金額(大於)：</td>
					  <td class="td02-c"><input type="text" name="Donate_Amt" size="16" class="font9"></td>
                      <td class="td02-c" align="right" width="130">捐款總金額起訖：</td>
                      <td class="td02-c"><input type="text" name="Donate_Total_Begin" size="12" class="font9"> ~ <input type="text" name="Donate_Total_End" size="12" class="font9"></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width="130">身分別：</td>
                      <td class="td02-c" colspan="3" width="570">
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
                      <td class="td02-c" align="right">生日月份：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        Response.Write "<SELECT Name='Birthday_Month' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        Response.Write "<OPTION> </OPTION>"
                        For I=1 To 12
                          Response.Write "<OPTION value='"&I&"'>"&I&"月</OPTION>"
                        Next
                        Response.Write "</SELECT>"
                      %>
                      </td>
                    </tr>
					<tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
  </div>
<div id="tb_month2" style="display:none" align="center">
<form name="form_Month2" method="post" action="donate_month_report2_qry.asp" target="main">
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
				  <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">定期奉獻授權到期明細表</td>
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
				  <td class="td02-c" align="right" width="120">授權終止年月：</td>
                      <td class="td02-c">
                      <%
                        Response.Write "<SELECT Name='Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
						For i=year_now-5 to year_now+5
							IF i = year_now then
								Response.Write "<OPTION selected value='"&year_now&"'>"&year_now&"</OPTION>"
							Else
								Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
							END IF
						NEXT
						
                        Response.Write "</SELECT>年"
                      %>
					  
					  <%
                        Response.Write "<SELECT Name='Month' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        For I=1 To 12
							IF i=month_now then
								Response.Write "<OPTION selected value='"&month_now&"'>"&month_now&"</OPTION>"
							Else
								Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
							END IF
                        Next
                        Response.Write "</SELECT>月"
                      %>
                      </td>
				</tr>
				<tr>
				  <td class="td02-c" align="right" width="120">授權方式：</td>
				  <td class="td02-c" colspan="3" width="610">
				  <%
					SQL="Select distinct Donate_Payment From PLEDGE"
					FName="Donate_Payment"
					Listfield="Donate_Payment"
					BoundColumn=""
					call CheckBoxList (SQL,FName,Listfield,BoundColumn)
				  %>
				  </td>
				</tr>
				<tr>
				  <td class="td02-c" align="right">不含：</td>
				  <td class="td02-c" colspan="3">
					<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
					<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
					<input type="checkbox" name="Sex" value="D" checked>歿
				  </td>
				</tr>
				<!--#include file="../include/calendar2.asp"-->
				<tr>
				  <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
				</tr>
				<tr>
				  <td colspan="8" width="100%">
				  <%
						  Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
						  Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_month3" style="display:none" align="center">
<form name="form_Month3" method="post" action="donate_month_report3_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">郵寄月刊明細表</td>
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
                      <td class="td02-c" align="right">身分別：</td>
                      <td class="td02-c" colspan="3" width="530">
						  <input type="radio" name="IsMember" value="N" checked> 捐款人  
						  <input type="radio" name="IsMember" value="Y"> 非捐款人
                      </td>
                    </tr>
					<tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_month4" style="display:none" align="center">
<form name="form_Month4" method="post" action="donate_month_report4_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">郵寄電子報明細表</td>
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
					  <td class="td02-c" align="right">捐款金額(大於)：</td>
                      <td class="td02-c"><input type="text" name="Donate_Amt" size="12" class="font9">
                      </td>
                      <td class="td02-c" align="right">捐款累計金額(大於)：</td>
                      <td class="td02-c"><input type="text" name="Donate_Total_Amt" size="12" class="font9">
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
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_month5" style="display:none" align="center">
<form name="form_Month5" method="post" action="donate_month_report5_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">當月各日捐款方式統計表</td>
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
                      <td class="td02-c" align="left">
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
                      <td class="td02-c" align="right">統計年月：</td>
                      <td class="td02-c" align="left">
                      <%
                        Response.Write "<SELECT Name='Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
						For i=year_now-10 to year_now
							IF i = year_now then
								Response.Write "<OPTION selected value='"&year_now&"'>"&year_now&"</OPTION>"
							Else
								Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
							END IF
						NEXT
						
                        Response.Write "</SELECT>年"
                      %>
					  
					  <%
                        Response.Write "<SELECT Name='Month' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        month_now=month(Date)
                        For I=1 To 12
							IF i=month_now then
								Response.Write "<OPTION selected value='"&month_now&"'>"&month_now&"</OPTION>"
							Else
								Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
							END IF
                        Next
                        Response.Write "</SELECT>月"
                      %>
                      </td>
                    </tr>
                    
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_season1" style="display:none" align="center">
<form name="form_Season1" method="post" action="donate_season_report1_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">非建台奉獻統計分析表</td>
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
                      <td class="td02-c" align="right">捐款期間：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Season1")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Season1")%></td>
                    </tr>
                    
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_season2" style="display:none" align="center">
<form name="form_Season2" method="post" action="donate_season_report2_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">查詢單筆捐款金額累計</td>
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
                      <td class="td02-c" align="right">捐款金額(大於)：</td>
                      <td class="td02-c" colspan="3"><input type="text" name="Donate_Amt" size="12" class="font9">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款期間：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Season2")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Season2")%></td>
                    </tr>
                    <tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_season3" style="display:none" align="center">
<form name="form_Season3" method="post" action="donate_season_report3_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">捐款用途各項總額明細表</td>
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
					  <td class="td02-c" align="right">捐款金額(大於)：</td>
                      <td class="td02-c"><input type="text" name="Donate_Amt" size="12" class="font9">
                      </td>
                      <td class="td02-c" align="right">捐款累計金額(大於)：</td>
                      <td class="td02-c"><input type="text" name="Donate_Total_Amt" size="12" class="font9">
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
                    
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_season4" style="display:none" align="center">
<form name="form_Season4" method="post" action="donate_season_report4_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">捐款總額明細表</td>
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
					  <td class="td02-c" align="right">捐款金額(大於)：</td>
                      <td class="td02-c"><input type="text" name="Donate_Amt" size="12" class="font9">
                      </td>
                      <td class="td02-c" align="right"></td>
                      <td class="td02-c" width="200"></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款累計金額(大於)：</td>
                      <td class="td02-c"><input type="text" name="Donate_Total_Amt" size="12" class="font9">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款期間：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Season4")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Season4")%></td>
                    </tr>
					<tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_other2" style="display:none" align="center">
<form name="form_Other2" method="post" action="donate_other_report2_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">DVD贈品索取人報表</td>
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
                      <td class="td02-c" align="right"><font color="red">*</font>贈送日期</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Gift_Date_Begin",Gift_Date_Begin,"form_Other2")%> ~ <%call Calendar2("Gift_Date_End",Gift_Date_End,"form_Other2")%></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">物品名稱：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL="Select Linked2_Name From Linked2 L2 INNER JOIN Linked L ON L.Ser_No=L2.Linked_Id Where Linked_Name = '歷年天使DVD贈品' Order By Linked2_Seq"
                        FName="Linked2_Name"
                        Listfield="Linked2_Name"
                        BoundColumn=""
						menusize=1
                        call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                      %>
                      </td>
                    </tr>
                    <tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_other3" style="display:none" align="center">
<form name="form_Other3" method="post" action="donate_other_report3_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">雷同資料查詢</td>
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
					  <td class="td02-c" align="right"><font color="red">*</font>查詢：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_SameName" Id="Is_SameName" value="1" checked>同姓名
						<input type="checkbox" name="Is_SameAddress" Id="Is_SameAddress" value="2" onclick="AddressChange()">同地址
					  </td>
					</tr>
					<tbody id="Is_Address" style="display:none">
					<tr>
                      <td class="td02-c" align="right">地址：</td>
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="Address" id="Address1" value="1" checked >通訊地址&nbsp;
                        <input type="radio" name="Address" id="Address2" value="2">收據地址&nbsp;
                      </td>
                    </tr>
					</tbody>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_year1" style="display:none" align="center">
<form name="form_Year1" method="post" action="donate_year_report1_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">捐款用途總額明細表</td>
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
					  <td class="td02-c" align="right" width="130">捐款金額(大於)：</td>
					  <td class="td02-c"><input type="text" name="Donate_Amt" size="16" class="font9"></td>
                      <td class="td02-c" align="right">捐款總金額起訖：</td>
                      <td class="td02-c"><input type="text" name="Donate_Total_Begin" size="12" class="font9"> ~ <input type="text" name="Donate_Total_End" size="12" class="font9"></td>
                    </tr>
					<tr>
                      <td class="td02-c" align="right">捐款期間：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Year1")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Year1")%></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width="130">捐款用途：</td>
                      <td class="td02-c" colspan="3" width="570">
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
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_year2" style="display:none" align="center">
<form name="form_Year2" method="post" action="donate_year_report2_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">捐款總額與月刊索取明細表</td>
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
                      <td class="td02-c" align="right">捐款年：</td>
                      <td class="td02-c">
                      <%
                        Response.Write "<SELECT Name='Donate_Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
						For i=year_now-10 to year_now
							IF i = year_now then
								Response.Write "<OPTION selected value='"&year_now&"'>"&year_now&"</OPTION>"
							Else
								Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
							END IF
						NEXT
                        Response.Write "</SELECT>年"
                      %>
                      </td>
					  <td class="td02-c" align="right"></td>
                      <td class="td02-c" width="200"></td>
                    </tr>
					<tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_year3" style="display:none" align="center">
<form name="form_Year3" method="post" action="donate_year_report3_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">海外捐款總額明細表</td>
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
					  <td class="td02-c" align="right">捐款金額(大於)：</td>
					  <td class="td02-c" colspan="3"><input type="text" name="Donate_Amt" size="12" class="font9"></td>
                    </tr>
					<tr>
                      <td class="td02-c" align="right">捐款累計金額(大於)：</td>
                      <td class="td02-c" colspan="3"><input type="text" name="Donate_Total_Amt" size="16" class="font9">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款期間：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Year3")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Year3")%></td>
                    </tr>
                    <tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_year4" style="display:none" align="center">
<form name="form_Year4" method="post" action="donate_year_report4_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">國內地區捐款總額明細表</td>
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
					  <td class="td02-c" align="right">捐款金額(大於)：</td>
					  <td class="td02-c" colspan="3"><input type="text" name="Donate_Amt" size="12" class="font9"></td>
                    </tr>
					<tr>
                      <td class="td02-c" align="right">捐款累計金額(大於)：</td>
                      <td class="td02-c" colspan="3"><input type="text" name="Donate_Total_Amt" size="16" class="font9">
                      </td>
                    </tr>
					<tr>
                      <td class="td02-c" align="right">捐款期間：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Year4")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Year4")%></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">縣市：</td>
                      <td class="td02-c" colspan="3" width="480">
                      <%
                        SQL="select Zip_Code=Name from CODECITYNew where ParentCityID = '0' Order By Sort"
                        FName="Zip_Code"
                        Listfield="Zip_Code"
                        BoundColumn=""
                        call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                      %>
                      </td>
                    </tr>
                    <tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_year5" style="display:none" align="center">
<form name="form_Year5" method="post" action="donate_year_report5_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">期間身份別捐款總額明細表</td>
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
					  <td class="td02-c" align="right">捐款金額(大於)：</td>
					  <td class="td02-c"><input type="text" name="Donate_Amt" size="16" class="font9"></td>
					  <td class="td02-c" align="right">捐款累計金額(大於)：</td>
                      <td class="td02-c"><input type="text" name="Donate_Total_Amt" size="12" class="font9"></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">身分別：</td>
                      <td class="td02-c" colspan="3" width="530">
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
                      <td class="td02-c" align="right">捐款年：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        Response.Write "<SELECT Name='Donate_Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        For i=year_now-10 to year_now
							IF i = year_now then
								Response.Write "<OPTION selected value='"&year_now&"'>"&year_now&"</OPTION>"
							Else
								Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
							END IF
						NEXT
                        Response.Write "</SELECT>年"
                      %>
                      </td>
                    </tr>
					<tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_year6" style="display:none" align="center">
<form name="form_Year6" method="post" action="donate_year_report6_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">期間單筆捐款金額次數明細表</td>
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
					  <td class="td02-c" align="right">捐款金額(大於)：</td>
					  <td class="td02-c" colspan="3"><input type="text" name="Donate_Amt" size="16" class="font9"></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款年：</td>
                      <td class="td02-c">
                      <%
                        Response.Write "<SELECT Name='Donate_Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        For i=year_now-10 to year_now
							IF i = year_now then
								Response.Write "<OPTION selected value='"&year_now&"'>"&year_now&"</OPTION>"
							Else
								Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
							END IF
						NEXT
                        Response.Write "</SELECT>年"
                      %>
                      </td>
					  <td class="td02-c" align="right"></td>
                      <td class="td02-c" width="200"></td>
                    </tr>
					<tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_year7" style="display:none" align="center">
<form name="form_Year7" method="post" action="donate_year_report7_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">新增捐款人總額明細表</td>
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
					  <td class="td02-c" align="right">捐款金額(大於)：</td>
					  <td class="td02-c" colspan="3"><input type="text" name="Donate_Amt" size="16" class="font9"></td>
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
                      <td class="td02-c" align="right">捐款年：</td>
                      <td class="td02-c">
                      <%
                        Response.Write "<SELECT Name='Donate_Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        For i=year_now-10 to year_now
							IF i = year_now then
								Response.Write "<OPTION selected value='"&year_now&"'>"&year_now&"</OPTION>"
							Else
								Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
							END IF
						NEXT
                        Response.Write "</SELECT>年"
                      %>
                      </td>
					  <td class="td02-c" align="right">首捐年：</td>
                      <td class="td02-c">
                      <%
                        Response.Write "<SELECT Name='First_DonateYear' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        For i=year_now-10 to year_now
							Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
						NEXT
                        Response.Write "</SELECT>年"
                      %>
                      </td>
                    </tr>
					<tr>
					  <td class="td02-c" align="right">末捐日期(起)：</td>
					  <td class="td02-c" colspan="3"><%call Calendar2("Last_Donate_Date",Donate_Date_Begin,"form_Year7")%></td>
                    </tr>
					<tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_year8" style="display:none" align="center">
<form name="form_Year8" method="post" action="donate_year_report8_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">經常奉獻總額明細表</td>
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
					  <td class="td02-c" align="right">捐款金額(大於)：</td>
                      <td class="td02-c"><input type="text" name="Donate_Amt" size="12" class="font9">
                      </td>
                      <td class="td02-c" align="right"></td>
                      <td class="td02-c" width="200"></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款累計金額(大於)：</td>
                      <td class="td02-c"><input type="text" name="Donate_Total_Amt" size="16" class="font9">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款期間：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Year8")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Year8")%></td>
                    </tr>
					<tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_Abroad" value="A" checked>海外地址
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_year9" style="display:none" align="center">
<form name="form_Year9" method="post" action="donate_year_report9_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">歷年來大陸奉獻天使累計奉獻明細表</td>
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
                      <td class="td02-c" align="right">捐款累計金額(大於)：</td>
                      <td class="td02-c" width="450"><input type="text" name="Donate_Total_Amt" size="12" class="font9">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款期間：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Year9")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Year9")%></td>
                    </tr>
					<tr>
					  <td class="td02-c" align="right">不含：</td>
					  <td class="td02-c" colspan="3">
						<input type="checkbox" name="Is_ErrAddress" value="W" checked>錯址
						<input type="checkbox" name="Sex" value="D" checked>歿
					  </td>
					</tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
<div id="tb_other1" style="display:none" align="center">
<form name="form_Other1" method="post" action="donate_other_report1_qry.asp" target="main">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">人數統計查詢報表</td>
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
                      <td class="td02-c" align="right">查詢期間：</td>
                      <td class="td02-c" colspan="3"><%call Calendar2("Donate_Date_Begin",Donate_Date_Begin,"form_Other1")%> ~ <%call Calendar2("Donate_Date_End",Donate_Date_End,"form_Other1")%></td>
                    </tr>
                    
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
    			              Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick(this.form)'><img src='../images/print.gif' width='20' height='20'><br>報表</button>&nbsp;"
    			              Response.Write "<button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Export_OnClick(this.form)'><img src='../images/excel.gif' width='20' height='20'><br>匯出</button>"	
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
</div>
</body>
</html>
<script> 
function setChange() 
{ 
	if (document.form.select.value == "week1") { 
		document.all.tb_week1.style.display = "block"; 
	} 
	else { 
		document.all.tb_week1.style.display = "none"; 
	} 
	if (document.form.select.value == "week2") { 
		document.all.tb_week2.style.display = "block"; 
	} 
	else { 
		document.all.tb_week2.style.display = "none"; 
	} 
	if (document.form.select.value == "week3") { 
		document.all.tb_week3.style.display = "block"; 
	} 
	else { 
		document.all.tb_week3.style.display = "none"; 
	} 
	if (document.form.select.value == "month1") { 
		document.all.tb_month1.style.display = "block"; 
	} 
	else { 
		document.all.tb_month1.style.display = "none"; 
	} 
	if (document.form.select.value == "month2") { 
		document.all.tb_month2.style.display = "block"; 
	} 
	else { 
		document.all.tb_month2.style.display = "none"; 
	} 
	if (document.form.select.value == "month3") { 
		document.all.tb_month3.style.display = "block"; 
	} 
	else { 
		document.all.tb_month3.style.display = "none"; 
	} 
	if (document.form.select.value == "month4") { 
		document.all.tb_month4.style.display = "block"; 
	} 
	else { 
		document.all.tb_month4.style.display = "none"; 
	} 
	if (document.form.select.value == "month5") { 
		document.all.tb_month5.style.display = "block"; 
	} 
	else { 
		document.all.tb_month5.style.display = "none"; 
	} 
	if (document.form.select.value == "season1") { 
		document.all.tb_season1.style.display = "block"; 
	} 
	else { 
		document.all.tb_season1.style.display = "none"; 
	} 
	if (document.form.select.value == "season2") { 
		document.all.tb_season2.style.display = "block"; 
	} 
	else { 
		document.all.tb_season2.style.display = "none"; 
	} 
	if (document.form.select.value == "season3") { 
		document.all.tb_season3.style.display = "block"; 
	} 
	else { 
		document.all.tb_season3.style.display = "none"; 
	} 
	if (document.form.select.value == "season4") { 
		document.all.tb_season4.style.display = "block"; 
	} 
	else { 
		document.all.tb_season4.style.display = "none"; 
	} 
	if (document.form.select.value == "year1") { 
		document.all.tb_year1.style.display = "block"; 
	} 
	else { 
		document.all.tb_year1.style.display = "none"; 
	} 
	if (document.form.select.value == "year2") { 
		document.all.tb_year2.style.display = "block"; 
	} 
	else { 
		document.all.tb_year2.style.display = "none"; 
	} 
	if (document.form.select.value == "year3") { 
		document.all.tb_year3.style.display = "block"; 
	} 
	else { 
		document.all.tb_year3.style.display = "none"; 
	} 
	if (document.form.select.value == "year4") { 
		document.all.tb_year4.style.display = "block"; 
	} 
	else { 
		document.all.tb_year4.style.display = "none"; 
	} 
	if (document.form.select.value == "year5") { 
		document.all.tb_year5.style.display = "block"; 
	} 
	else { 
		document.all.tb_year5.style.display = "none"; 
	} 
	if (document.form.select.value == "year6") { 
		document.all.tb_year6.style.display = "block"; 
	} 
	else { 
		document.all.tb_year6.style.display = "none"; 
	} 
	if (document.form.select.value == "year7") { 
		document.all.tb_year7.style.display = "block"; 
	} 
	else { 
		document.all.tb_year7.style.display = "none"; 
	} 
	if (document.form.select.value == "year8") { 
		document.all.tb_year8.style.display = "block"; 
	} 
	else { 
		document.all.tb_year8.style.display = "none"; 
	} 
	if (document.form.select.value == "year9") { 
		document.all.tb_year9.style.display = "block"; 
	} 
	else { 
		document.all.tb_year9.style.display = "none"; 
	} 
	if (document.form.select.value == "other1") { 
		document.all.tb_other1.style.display = "block"; 
	} 
	else { 
		document.all.tb_other1.style.display = "none"; 
	} 
	if (document.form.select.value == "other2") { 
		document.all.tb_other2.style.display = "block"; 
	} 
	else { 
		document.all.tb_other2.style.display = "none"; 
	} 
	if (document.form.select.value == "other3") { 
		document.all.tb_other3.style.display = "block"; 
	} 
	else { 
		document.all.tb_other3.style.display = "none"; 
	} 
} 
function AddressChange(){
if (document.form_Other3.Is_SameAddress.checked) { 
		document.getElementById("Is_Address").style.display = "block"; 
	} 
	else { 
		document.getElementById("Is_Address").style.display = "none"; 
	} 
}
function Report_OnClick(formName){
	//alert (formName.action.value) ;
	//if (true) return;
	//alert(formName.name);
	if (formName.name=='form_Week2'){
	  if(document.form_Week2.Donate_Date_Begin.value==''){
	  <%call CheckStringK("Donate_Date_Begin","捐款日期","form_Week2")%>
	  }
	  if(document.form_Week2.Donate_Date_End.value==''){
	  <%call CheckStringK("Donate_Date_End","捐款日期","form_Week2")%>
	  }
	}
	if (formName.name=='form_Other2'){
	  if(document.form_Other2.Gift_Date_Begin.value==''){
	  <%call CheckStringK("Gift_Date_Begin","贈送日期","form_Other2")%>
	  }
	  if(document.form_Other2.Gift_Date_End.value==''){
	  <%call CheckStringK("Gift_Date_End","贈送日期","form_Other2")%>
	  }
	}
	if (formName.name=='form_Other3'){
	  if(document.form_Other3.Is_SameName.checked==false && document.form_Other3.Is_SameAddress.checked==false){
	  alert('請務必選取其中一個條件！');return false;
	  }
	}
	if(confirm('您是否確定要將查詢結果列表？')){
	formName.target="report"
	formName.action.value="report"
	window.open('','report','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600')
	formName.submit();
	}
}
function Export_OnClick(formName){
	if (formName.name=='form_Week2'){
	  if(document.form_Week2.Donate_Date_Begin.value==''){
	  <%call CheckStringK("Donate_Date_Begin","捐款日期","form_Week2")%>
	  }
	  if(document.form_Week2.Donate_Date_End.value==''){
	  <%call CheckStringK("Donate_Date_End","捐款日期","form_Week2")%>
	  }
	}
	if (formName.name=='form_Other2'){
	  if(document.form_Other2.Gift_Date_Begin.value==''){
	  <%call CheckStringK("Gift_Date_Begin","贈送日期","form_Other2")%>
	  }
	  if(document.form_Other2.Gift_Date_End.value==''){
	  <%call CheckStringK("Gift_Date_End","贈送日期","form_Other2")%>
	  }
	}
	if(confirm('您是否確定要將查詢結果匯出？')){
	formName.target='main';
	formName.action.value='export';
	formName.submit();
	}
}

</script>
<%
Function Calendar2 (FName,Val,formName)
  Response.Write "<input type=text name="&FName&" size=10 class=font9 value="&Val&">" & Chr(13) & Chr(10)
  Response.Write "<a href onclick=cal19.select(document."&formName&"."&FName&",'"&FName&"','yyyy/MM/dd');>" & Chr(13) & Chr(10)
  Response.Write "<img border=0 src=../images/date.gif width=16 height=14></a>" & Chr(13) & Chr(10)
End Function
Function CheckStringK (FName,ListField,formName)
  Response.Write "if(document."&formName&"." & FName & ".value==''){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位不可為空白！');" & Chr(13) & Chr(10)
  Response.Write "  document."&formName&"." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function
%>