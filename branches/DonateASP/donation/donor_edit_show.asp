<!--#include file="../include/dbfunctionJ.asp"-->
<%

SQL_IS="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"' "
Set RS_IS=Server.CreateObject("ADODB.Recordset")
RS_IS.Open SQL_IS,Conn,1,1
IsContribute=RS_IS("IsContribute")
IsMember=RS_IS("IsMember")
IsMemberPre=RS_IS("IsMemberPre")
IsShopping=RS_IS("IsShopping")
Donate_IsFdc=RS_IS("Donate_IsFdc")
RS_IS.Close
Set RS_IS=Nothing

SQL="Select * From DONOR Where Donor_id='"&request("donor_id")&"'"
Call QuerySQL(SQL,RS)
%>
<%Prog_Id="donor"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <%If request("ctype")<>"" Then%>
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <%Else%>
      <input type="hidden" name="ctype" value="donor_edit">
      <%End If%>
      <input type="hidden" name="donor_id" value="<%=request("donor_id")%>">
      <input type="hidden" name="Introducer_Id" value="<%=RS("Introducer_Id")%>">
      <input type="hidden" name="IsMemberPre" value="<%=IsMemberPre%>">
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
                  <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                    <tr>
                      <td width="99%">
                        <table border="1" cellpadding="2" style="border-collapse: collapse" width="100%" height="25" cellspacing="1">
                          <tr>
                            <td class="button2-bg"><img border="0" src="../images/red_arrow.gif" align="texttop">捐款人資料</td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                          	<td align="right">捐款人編號：</td>
                            <td align="left" colspan="7">&nbsp;<%=RS("Donor_ID")%></td>
                          </tr>
                          <tr>
                            <td align="right">捐款人：</td>
                            <td align="left" colspan="3"><input type="text" name="Donor_Name" size="46" class="font9" maxlength="100" value="<%=Data_Minus(RS("Donor_Name"))%>"></td>
                            <td align="left" colspan="4">
                              <!--類別：
                              <%
                                'SQL="Select Category=CodeDesc From CASECODE Where codetype='Category' Order By Seq"
                                'FName="Category"
                                'Listfield="Category"
                                'menusize="1"
                                'BoundColumn=RS("category")
                                'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>-->
                              &nbsp;性別：
                              <%
                                SQL="Select Sex=CodeDesc From CASECODE Where codetype='Sex' Order By Seq"
                                FName="Sex"
                                Listfield="Sex"
                                menusize="1"
                                BoundColumn=RS("Sex")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                              &nbsp;稱謂：
                              <%
                                SQL="Select Title=CodeDesc From CASECODE Where codetype='Title' Order By Seq"
                                FName="Title"
                                Listfield="Title"
                                menusize="1"
                                BoundColumn=RS("Title")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">身份別：</td>
                            <td align="left" colspan="7">
                              <%
                                SQL="Select Donor_Type=CodeDesc From CASECODE Where codetype='DonorType' Order By Seq"
                                FName="Donor_Type"
                                Listfield="Donor_Type"
                                BoundColumn=RS("Donor_Type")
                                call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                              %>
                            </td>
                          </tr>
                          <tr>
                            <td align="right" width="12%">身分證/統編：</td>
                            <td align="left" width="15%"><input type="text" name="IDNo" size="15" class="font9" maxlength="10" onKeyUp="javascript:UCaseIDNO();" value="<%=RS("IDNo")%>"></td>
                            <td align="right" width="10%">出生日期：</td>
                            <td align="left" width="15%"><%call Calendar("Birthday",RS("Birthday"))%></td>
                            <input type="hidden" name="Education" size="15" class="font9" maxlength="10" value="<%=RS("Education")%>">
                            <input type="hidden" name="Occupation" size="15" class="font9" maxlength="10" value="<%=RS("Occupation")%>">
                            <input type="hidden" name="Marriage" size="15" class="font9" maxlength="10" value="<%=RS("Marriage")%>">
                            <input type="hidden" name="Religion" size="15" class="font9" maxlength="10" value="<%=RS("Religion")%>">
                            <input type="hidden" name="ReligionName" size="15" class="font9" maxlength="10" value="<%=RS("ReligionName")%>">
                            <td align="right">手機：</td>
                            <td align="left"><input type="text" name="Cellular_Phone" size="15" class="font9" maxlength="40" value="<%=RS("Cellular_Phone")%>"></td>
                            <!--<td align="right" width="10%">教育程度：</td>
                            <td align="left" width="15%">
                            <%
                              'SQL="Select Education=CodeDesc From CASECODE Where codetype='Education' Order By Seq"
                              'FName="Education"
                              'Listfield="Education"
                              'menusize="1"
                              'BoundColumn=RS("Education")
                              'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td align="right" width="8%">職業別：</td>
                            <td align="left" width="15%">
                            <%
                              'SQL="Select Occupation=CodeDesc From CASECODE Where codetype='Occupation' Order By Seq"
                              'FName="Occupation"
                              'Listfield="Occupation"
                              'menusize="1"
                              'BoundColumn=RS("Occupation")
                              'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>-->
                          </tr>
                          <!--<tr>
                            <td align="right">婚姻狀況：</td>
                            <td align="left">
                            <%
                              'SQL="Select Marriage=CodeDesc From CASECODE Where codetype='Marriage' Order By Seq"
                              'FName="Marriage"
                              'Listfield="Marriage"
                              'menusize="1"
                              'BoundColumn=RS("Marriage")
                              'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td align="right">宗教信仰：</td>
                            <td align="left">
                            <%
                              'SQL="Select Religion=CodeDesc From CASECODE Where codetype='Religion' Order By Seq"
                              'FName="Religion"
                              'Listfield="Religion"
                              'menusize="1"
                              'BoundColumn=RS("Religion")
                              'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td colspan="4">
                            	&nbsp;&nbsp;&nbsp;&nbsp;道場/教會名稱：
                              <input type="text" name="ReligionName" size="39" class="font9" maxlength="40" value="<%'=Data_Minus(RS("ReligionName"))%>">
                            </td>
                          </tr>-->                            
                          <tr>                            
                            <td align="right">電話：</td>
                            <td align="left" colspan="3"><input type="text" name="Tel_Office_Loc" size="5" class="font9" maxlength="5" value="<%=RS("Tel_Office_Loc")%>">-<input type="text" name="Tel_Office" size="15" class="font9" maxlength="40" value="<%=RS("Tel_Office")%>">-<input type="text" name="Tel_Office_Ext" size="5" class="font9" maxlength="5" value="<%=RS("Tel_Office_Ext")%>"></td>
                            <input type="hidden" name="Tel_Home" size="15" class="font9" maxlength="10" value="<%=RS("Tel_Home")%>">
                            <!--<td align="right">電話(夜)：</td>
                            <td align="left"><input type="text" name="Tel_Home" size="15" class="font9" maxlength="40" value="<%'=RS("Tel_Home")%>"></td>-->
                            <td align="right">傳真：</td>
                            <td align="left"><input type="text" name="Fax_Loc" size="5" class="font9" maxlength="5" value="<%=RS("Fax_Loc")%>">-<input type="text" name="Fax" size="15" class="font9" maxlength="40" value="<%=RS("Fax")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">E-Mail：</td>
                            <td align="left"><input type="text" name="Email" size="25" class="font9" maxlength="80" value="<%=RS("Email")%>"></td>
                            <td align="right">聯絡人：</td>
                            <td align="left"><input type="text" name="Contactor" size="15" class="font9" maxlength="40" value="<%=Data_Minus(RS("Contactor"))%>"></td>
                            <input type="hidden" name="OrgName" size="15" class="font9" maxlength="10" value="<%=RS("OrgName")%>">
                            <!--<td align="right">服務單位：</td>
                            <td align="left"><input type="text" name="OrgName" size="15" class="font9" maxlength="40" value="<%'=RS("OrgName")%>"></td>-->
                            <td align="right">職稱：</td>
                            <td align="left"><input type="text" name="JobTitle" size="15" class="font9" maxlength="40" value="<%=RS("JobTitle")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">通訊地址：</td>
                            <td align="left" colspan="7">
                              <input type="radio" name="IsAbroad" id="IsAbroadY" value="N" <%If RS("IsAbroad")="N" Or RS("IsAbroad")="" Then Response.Write "checked" End If%> >台灣本島
                              <%call CodeCity2 ("form","ZipCode",RS("ZipCode"),"City",RS("City"),"Area",RS("Area"),"Y")%>
                            	<input type='text' class='font9' name='Street' size='25' maxlength='40' value='<%=RS("Street")%>'>大道/路/街/部落
                              <br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                              <%
                                SQL="Select SectionType=CodeDesc From CASECODE Where codetype='SectionType' Order By Seq"
                                FName="SectionType"
                                Listfield="SectionType"
                                menusize="1"
                                BoundColumn=RS("Section")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>段
                            	<input type='text' class='font9' name='Lane' size='5' maxlength='10' value='<%=RS("Lane")%>'>巷
                            	<input type='text' class='font9' name='Alley' size='5' maxlength='10' value='<%=RS("Alley")%>'>
                              <%
                                SQL="Select AlleyType=CodeDesc From CASECODE Where codetype='AlleyType' Order By Seq"
                                FName="AlleyType"
                                Listfield="AlleyType"
                                menusize="1"
                                If RS("Address")<>"" Then
                                  If Instr(RS("Address"),"弄")>0 Then
                                    BoundColumn="弄"
                                  ElseIf Instr(RS("Address"),"衖")>0 Then
                                    BoundColumn="衖"
                                  Else
                                    BoundColumn=""
                                  End If
                                Else
                                  BoundColumn=""
                                End If
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            	<input type='text' class='font9' name='HouseNo' size='5' maxlength='10' value='<%=RS("HouseNo")%>'>-<input type='text' class='font9' name='HouseNoSub' size='5' maxlength='10' value='<%=RS("HouseNoSub")%>'>號
                            	<input type='text' class='font9' name='Floor' size='5' maxlength='10' value='<%=RS("Floor")%>'>-<input type='text' class='font9' name='FloorSub' size='5' maxlength='10' value='<%=RS("FloorSub")%>'>樓
                            	<input type='text' class='font9' name='Room' size='5' maxlength='10' value='<%=RS("Room")%>'>室
                              <br />
                              <input type="radio" name="IsAbroad" id="IsAbroadN" value="Y" <%If RS("IsAbroad")="Y" Then Response.Write "checked" End If%> >海外地址
                              <input type='text' class='font9' name='OverseasAddress' size='50' maxlength='100' value='<%=RS("OverseasAddress")%>'>
                              
                              國家/省城市/區
							  <input type='text' class='font9' name='OverseasCountry' size='20' maxlength='50' value='<%=RS("OverseasCountry")%>'>
                              
                            </td>
                          </tr>
                          <tr>
                            <td align="right">文宣品：</td>
                            <td align="left" colspan="7">
                            	<input type="checkbox" name="IsSendNews" value="Y" <%If RS("IsSendNews")="Y" Then Response.Write "checked" End If%> >會訊
                              <input type="checkbox" name="IsSendEpaper" value="Y" <%If RS("IsSendEpaper")="Y" Then Response.Write "checked" End If%> >電子報
                              <input type="checkbox" name="IsSendYNews" value="Y" <%If RS("IsSendYNews")="Y" Then Response.Write "checked" End If%> >年報
                              <input type="checkbox" name="IsBirthday" value="Y" <%If RS("IsBirthday")="Y" Then Response.Write "checked" End If%> >生日卡
                              <input type="checkbox" name="IsXmas" value="Y" <%If RS("IsXmas")="Y" Then Response.Write "checked" End If%> >賀卡
                            </td>
                          </tr>
                          <!--#include file="../include/calendar2.asp"-->
                          <tr>
                            <td width="99%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                          </tr>
                          <tr>
                            <td align="right">收據開立：</td>
                            <td align="left" colspan="7">
                              <%
                                SQL="Select Invoice_Type=CodeDesc From CASECODE Where codetype='InvoiceType' Order By Seq"
                                FName="Invoice_Type"
                                Listfield="Invoice_Type"
                                menusize="1"
                                BoundColumn=RS("Invoice_Type")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                              <input type="checkbox" name="IsSameAddress" value="Y" OnClick="SameAddress_OnClick()"><font color="#ff0000">收據資料同上</font>	
                              <input type="checkbox" name="IsAnonymous" value="Y" <%If RS("IsAnonymous")="Y" Then Response.Write "checked" End If%> >徵信錄匿名
                            	<!--&nbsp;匿名：
                            	<input type="text" name="NickName" size="13" class="font9" maxlength="20" value="<%=RS("NickName")%>">-->
                                <input type="hidden" name="NickName" size="13" class="font9" maxlength="20" value="<%=RS("NickName")%>">
                            	&nbsp;收據稱謂：
                              <%
                                SQL="Select Title2=CodeDesc From CASECODE Where codetype='Title' Order By Seq"
                                FName="Title2"
                                Listfield="Title2"
                                menusize="1"
                                BoundColumn=RS("Title2")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            </td>	
                          </tr>
                          <tr>
                            <td align="right">收據地址：</td>
                            <td align="left" colspan="7">
                              <input type="radio" name="IsAbroad_Invoice" id="IsAbroad_InvoiceY" value="N" <%If RS("IsAbroad_Invoice")="N" Or RS("IsAbroad_Invoice")="" Then Response.Write "checked" End If%> >台灣本島
                            	<%call CodeCity2 ("form","Invoice_ZipCode",RS("Invoice_ZipCode"),"Invoice_City",RS("Invoice_City"),"Invoice_Area",RS("Invoice_Area"),"N")%>
                            	<input type='text' class='font9' name='Invoice_Street' size='15' maxlength='40' value='<%=RS("Invoice_Street")%>'>大道/路/街/部落
                              <br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                              <%
                                SQL="Select Invoice_Section=CodeDesc From CASECODE Where codetype='SectionType' Order By Seq"
                                FName="Invoice_Section"
                                Listfield="Invoice_Section"
                                menusize="1"
                                BoundColumn=RS("Invoice_Section")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>段
                            	<input type='text' class='font9' name='Invoice_Lane' size='5' maxlength='10' value='<%=RS("Invoice_Lane")%>'>巷
                            	<input type='text' class='font9' name='Invoice_Alley' size='5' maxlength='10' value='<%=RS("Invoice_Alley")%>'>
                              <%
                                SQL="Select Invoice_AlleyType=CodeDesc From CASECODE Where codetype='AlleyType' Order By Seq"
                                FName="Invoice_AlleyType"
                                Listfield="Invoice_AlleyType"
                                menusize="1"
                                If RS("Invoice_Address")<>"" Then
                                  If Instr(RS("Invoice_Address"),"弄")>0 Then
                                    BoundColumn="弄"
                                  ElseIf Instr(RS("Invoice_Address"),"衖")>0 Then
                                    BoundColumn="衖"
                                  Else
                                    BoundColumn=""
                                  End If
                                Else
                                  BoundColumn=""
                                End If
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            	<input type='text' class='font9' name='Invoice_HouseNo' size='5' maxlength='10' value='<%=RS("Invoice_HouseNo")%>'>-<input type='text' class='font9' name='Invoice_HouseNoSub' size='5' maxlength='10' value='<%=RS("Invoice_HouseNoSub")%>'>號
                            	<input type='text' class='font9' name='Invoice_Floor' size='5' maxlength='10' value='<%=RS("Invoice_Floor")%>'>-<input type='text' class='font9' name='Invoice_FloorSub' size='5' maxlength='10' value='<%=RS("Invoice_FloorSub")%>'>樓
                            	<input type='text' class='font9' name='Invoice_Room' size='5' maxlength='10' value='<%=RS("Invoice_Room")%>'>室
                              <br />
                              <input type="radio" name="IsAbroad_Invoice" id="IsAbroad_InvoiceN" value="Y" <%If RS("IsAbroad_Invoice")="Y" Then Response.Write "checked" End If%> >海外地址
                              <input type='text' class='font9' name='Invoice_OverseasAddress' size='50' maxlength='100' value='<%=RS("Invoice_OverseasAddress")%>'>
                              
                              國家/省城市/區
							  <input type='text' class='font9' name='Invoice_OverseasCountry' size='20' maxlength='50' value='<%=RS("Invoice_OverseasCountry")%>'>
                              
                            </td>
                          </tr>
                          <tr>
                            <td align="right">收據抬頭：</td>
                            <td align="left" colspan="3"><input type="text" name="Invoice_Title" size="46" class="font9" maxlength="100" value="<%=Data_Minus(RS("Invoice_Title"))%>"></td>
                            <td align="left" colspan="4">
                              收據身分證/統編：
                              <input type="text" name="Invoice_IDNo" size="15" class="font9" maxlength="10" onKeyUp="javascript:UCaseInvoiceIDNo();" value="<%=RS("Invoice_IDNo")%>">
                            </td>
                          </tr>
                          <tr> 
							
                            <td align="right"> 徵信錄原則：</td>
                            <td align="left" >
                            <%
                              SQL="Select Report_Type=CodeDesc From CASECODE Where codetype='ReportType' Order By Seq"
                              FName="Report_Type"
                              Listfield="Report_Type"
                              menusize="1"
                              BoundColumn=RS("Report_Type")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
						
                            <td align="right">群組：</td>
                            <td align="left" >
                            <%
                              SQL="Select Report_Group=CodeDesc From CASECODE Where codetype='ReportGroup' Order By Seq"
                              FName="Report_Group"
                              Listfield="Report_Group"
                              menusize="1"
                              BoundColumn=RS("Report_Group")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
						
                            <td align="right">收據抬頭人數：</td>
                            <td align="left" ><input type="text" name="Invoice_Title_Man" size="2" class="font9" maxlength="2" value="<%=RS("Invoice_Title_Man")%>"></td>
                          </tr>
                          <%If Donate_IsFdc="Y" Then%>
                          <tr>
                            <td align="left" colspan="8">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="IsFdc" value="Y" <%If RS("IsFdc")="Y" Then Response.Write "checked" End If%> ><font color="#ff0000">本捐款人願意將捐款資料上傳至國稅局以供報稅用</font>&nbsp;(&nbsp;請填寫捐款人收據身分證/統編&nbsp;)</td>
                          </tr>
                          <%Else%>
                          <input type="hidden" name="IsFdc" value="<%=RS("IsFdc")%>">
                          <%End If%>
                          <tr>
                            <td align="right">捐款人備註：</td>
                            <td align="left" colspan="7">
                            	<textarea name="Remark" rows="6" cols="81" class="font9"><%=RS("Remark")%></textarea>
                            </td>
                          </tr>
						  <tr>
                            <td align="right">天使想說的話：</td>
                            <td align="left" colspan="7">
                            	<font><%=RS("ToGOODTV")%></font>
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                          </tr>
                          <tr>  
                            <td align="left" colspan="8">
                              <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                                <tr>
                          	      <td align="right">首次捐款日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateD" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateD")%>"></td>
                                  <td align="right">最近捐款日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateD" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateD")%>"></td>
                                  <td align="right">累計捐款次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoD" size="13" class="font9t" readonly value="<%if RS("Donate_NoD")<>"" And isnull(RS("Donate_NoD"))=false Then Response.write FormatNumber(RS("Donate_NoD"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計捐款金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalD" size="13" class="font9t" readonly value="<%if RS("Donate_TotalD")<>"" And isnull(RS("Donate_TotalD"))=false Then Response.write FormatNumber(RS("Donate_TotalD"),0)%>" style="text-align: right"></td>
                                </tr>
                                <%If IsContribute="Y" Then%>
                                <tr>
                          	      <td align="right">首次捐物日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateC" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateC")%>"></td>
                                  <td align="right">最近捐物日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateC" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateC")%>"></td>
                                  <td align="right">累計捐物次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoC" size="13" class="font9t" readonly value="<%if RS("Donate_NoC")<>"" And isnull(RS("Donate_NoC"))=false Then Response.write FormatNumber(RS("Donate_NoC"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計折合現金：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalC" size="13" class="font9t" readonly value="<%if RS("Donate_TotalC")<>"" And isnull(RS("Donate_TotalC"))=false Then Response.write FormatNumber(RS("Donate_TotalC"),0)%>" style="text-align: right"></td>
                                </tr>
                                <%End If%>
                                <%If IsMember="Y" Then%>
                                <tr>
                          	      <td align="right">首次繳費日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateM" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateM")%>"></td>
                                  <td align="right">最近繳費日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateM" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateM")%>"></td>
                                  <td align="right">累計繳費次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoM" size="13" class="font9t" readonly value="<%if RS("Donate_NoM")<>"" And isnull(RS("Donate_NoM"))=false Then Response.write FormatNumber(RS("Donate_NoM"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計繳費金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalM" size="13" class="font9t" readonly value="<%if RS("Donate_TotalM")<>"" And isnull(RS("Donate_TotalM"))=false Then Response.write FormatNumber(RS("Donate_TotalM"),0)%>" style="text-align: right"></td>
                                </tr>
                                <tr>
                          	      <td align="right">首次報名日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateA" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateA")%>"></td>
                                  <td align="right">最近報名日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateA" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateA")%>"></td>
                                  <td align="right">累計報名次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoA" size="13" class="font9t" readonly value="<%if RS("Donate_NoA")<>"" And isnull(RS("Donate_NoA"))=false Then Response.write FormatNumber(RS("Donate_NoA"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計報名金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalA" size="13" class="font9t" readonly value="<%if RS("Donate_TotalA")<>"" And isnull(RS("Donate_TotalA"))=false Then Response.write FormatNumber(RS("Donate_TotalA"),0)%>" style="text-align: right"></td>
                                </tr>
                                <%End If%>
                                <%If IsShopping="Y" Then%>
                                <tr>
                          	      <td align="right">首次購物日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateS" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateS")%>"></td>
                                  <td align="right">最近購物日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateS" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateS")%>"></td>
                                  <td align="right">累計購物次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoS" size="13" class="font9t" readonly value="<%if RS("Donate_NoS")<>"" And isnull(RS("Donate_NoS"))=false Then Response.write FormatNumber(RS("Donate_NoS"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計購物金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalS" size="13" class="font9t" readonly value="<%if RS("Donate_TotalS")<>"" And isnull(RS("Donate_TotalS"))=false Then Response.write FormatNumber(RS("Donate_TotalS"),0)%>" style="text-align: right"></td>
                                </tr>
                                <%End If%>
                                <%If IsMember="Y" Or IsShopping="Y" Then%>
                                <tr>
                          	      <td align="right">首次非捐日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateND" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateND")%>"></td>
                                  <td align="right">最近非捐日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateND" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateND")%>"></td>
                                  <td align="right">累計非捐次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoND" size="13" class="font9t" readonly value="<%if RS("Donate_NoND")<>"" And isnull(RS("Donate_NoND"))=false Then Response.write FormatNumber(RS("Donate_NoND"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計非捐金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalND" size="13" class="font9t" readonly value="<%if RS("Donate_TotalND")<>"" And isnull(RS("Donate_TotalND"))=false Then Response.write FormatNumber(RS("Donate_TotalND"),0)%>" style="text-align: right"></td>
                                </tr>
                                <%End If%>
                                <tr>
                                  <td align="right">資料建檔日期：</td>
                                  <td align="left"><input type="text" name="Create_Date" size="13" class="font9t" readonly value="<%=RS("Create_Date")%>"></td>
                                  <td align="right">資料建檔人員：</td>
                                  <td align="left"><input type="text" name="Create_User" size="13" class="font9t" readonly value="<%=RS("Create_User")%>"></td>
                            	    <td align="right">最後異動日期：</td>
                                  <td align="left"><input type="text" name="LastUpdate_Date" size="13" class="font9t" readonly value="<%=RS("LastUpdate_Date")%>"></td>
                                  <td align="right">最後異動人員：</td>
                                  <td align="left"><input type="text" name="LastUpdate_User" size="13" class="font9t" readonly value="<%=RS("LastUpdate_User")%>"></td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                          <tr>
                            <td align="center" colspan="8">                               
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
