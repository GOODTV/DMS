<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="donor"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="donor_list.asp" target="detail">
      <input type="hidden" name="action">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="1050"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
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
	          <table width="1050" border="0" cellspacing="0" cellpadding="0" align="center">
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
                      <td class="td02-c" align="right">捐款人：</td>
                      <td class="td02-c">
                      	<input type="text" name="Donor_Name" size="11" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                        &nbsp;
                      	捐款人編號：
		                    <input type="text" name="Member_No" size="11" class="font9" onKeyUp="javascript:UCaseMemberNo();" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
		                    &nbsp;
                      	舊編號：
		                    <input type="text" name="Donor_Id_Old" size="7" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                      	&nbsp;
                      	聯絡電話：
		                    <input type="text" name="Tel_Office" size="11" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
		                    &nbsp;
                        身份別：
                        <%
                          SQL="Select Donor_Type=CodeDesc From CASECODE Where codetype='DonorType' Order By Seq"
                          FName="Donor_Type"
                          Listfield="Donor_Type"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                      </td>
                      <td class="td02-c">
                        <input type="button" value=" 查 詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">	
                        <!--20131105Modify by GoodTV Tanya:增加「雷同資料查詢」--> <br/><font color="blue">1.捐款人編號,舊編號,台灣地址是精準查詢欄位<br/>2.其他欄位是雷同查詢欄位</font>
                          <!--20140620 取消「雷同資料查詢」功能
                        <input type="button" value="  雷同資料查詢  " name="fuzzy_query" class="addbutton" style="cursor:hand" onClick="FuzzyQuery_OnClick()">-->
                      </td>
                    </tr>     
					<tr>
						<td class="td02-c" align="right">收據抬頭：</td>
						<td class="td02-c"><input type="text" name="Invoice_Title" size="15" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                        &nbsp;
						身份證/統編：
		                    <input type="text" name="IDNo" size="10" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                      	&nbsp;
						文宣品：
						<!-- <input type="checkbox" name="IsSendNews" value="Y">會訊
                        <input type="checkbox" name="IsSendEpaper" value="Y">電子報
                        <input type="checkbox" name="IsSendYNews" value="Y">年報
                        <input type="checkbox" name="IsBirthday" value="Y">生日卡
                        <input type="checkbox" name="IsXmas" value="Y">賀卡 -->
						紙本月刊本數<input type="text" class="font9" name="IsSendNewsNum" size='3' maxlength='3'>                            	
                            	<input type="checkbox" name="IsDVD" value="Y">DVD
                              <input type="checkbox" name="IsSendEpaper" value="Y">電子文宣
						</td>
						<td class="td02-c">
                      	<input type="button" value=" 列 表 " name="report" class="cbutton" style="cursor:hand" onClick="Report_OnClick()">
                        <input type="button" value=" 匯 出 " name="export" class="cbutton" style="cursor:hand" onClick="Export_OnClick()">
                      	<input type="button" value="手機名單" name="mobile" class="cbutton" style="cursor:hand" onClick="Mobile_OnClick()">
                      	<input type="button" value="EMail名單" name="email" class="cbutton" style="cursor:hand" onClick="EMail_OnClick()">	
						</td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width='64'>地址：</td>
                      <td class="td02-c" colspan="3">
                      	<table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="75">
                            	<input type="checkbox" name="IsAbroad" value="Y" onclick="IsAbroad_OnClick()">海外地址                            	
                            </td>
                            <!--20131112 Add by GoodTV Tanya:增加海外地址查詢欄位-->
                            <td id="donor_abroad_address" style="display:none" colspan="9"> 
                            	<input type='text' class='font9' name='IsAbroad_Address' size='92' maxlength='200' value='<%=Request("IsAbroad_Address")%>'></td>
                            <td id="donor_address" style="display:block" colspan="9">                            	
							  							<%call CodeCity2 ("form","ZipCode",Request("ZipCode"),"City",Request("City"),"Area",Request("Area"),"Y")%>
                              <input type='text' class='font9' name='Street' size='25' maxlength='40' value='<%=Request("Street")%>'>大道/路/街/部落                              
                              <%
                                SQL="Select SectionType=CodeDesc From CASECODE Where codetype='SectionType' Order By Seq"
                                FName="SectionType"
                                Listfield="SectionType"
                                menusize="1"
                                BoundColumn=Request("Section")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>段
                            	<input type='text' class='font9' name='Lane' size='5' maxlength='10' value='<%=Request("Lane")%>'>巷
                            	<input type='text' class='font9' name='Alley' size='5' maxlength='10' value='<%=Request("Alley")%>'>
                              <%
                                SQL="Select AlleyType=CodeDesc From CASECODE Where codetype='AlleyType' Order By Seq"
                                FName="AlleyType"
                                Listfield="AlleyType"
                                menusize="1"
                                If Request("Address")<>"" Then
                                  If Instr(Request("Address"),"弄")>0 Then
                                    BoundColumn="弄"
                                  ElseIf Instr(Request("Address"),"衖")>0 Then
                                    BoundColumn="衖"
                                  Else
                                    BoundColumn=""
                                  End If
                                Else
                                  BoundColumn=""
                                End If
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            	<input type='text' class='font9' name='HouseNo' size='5' maxlength='10' value='<%=Request("HouseNo")%>'>之<input type='text' class='font9' name='HouseNoSub' size='5' maxlength='10' value='<%=Request("HouseNoSub")%>'>號
                            	<input type='text' class='font9' name='Floor' size='5' maxlength='10' value='<%=Request("Floor")%>'>樓之<input type='text' class='font9' name='FloorSub' size='5' maxlength='10' value='<%=Request("FloorSub")%>'>
                            	<input type='text' class='font9' name='Room' size='5' maxlength='10' value='<%=Request("Room")%>'>室
                            </td>
                          </tr>
                      	</table>
                      </td>                      
                    </tr>
                    <tr>
                      <td class="td02-c" colspan="3" width="100%"><iframe name="detail" src="donor_list.asp" height="415" width="100%" frameborder="0" scrolling="auto"></iframe></td>
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
  document.form.Donor_Name.focus();
}
function UCaseMemberNo(){
  document.form.Member_No.value=document.form.Member_No.value.toUpperCase();
}		
function Query_OnClick(){
	//20140122 Add by GoodTV Tanya:捐款人編號&舊編號驗證是否填入數字
	if(document.form.Member_No.value!=''){  	
  	<%call CheckNumberJ("Member_No","捐款人編號")%>
	}
	if(document.form.Donor_Id_Old.value!=''){  	
  	<%call CheckNumberJ("Donor_Id_Old","舊編號")%>
	}
	
  document.form.target="detail"
  document.form.action.value="query"
  document.form.submit();
}
function FuzzyQuery_OnClick(){
	//20140122 Add by GoodTV Tanya:捐款人編號&舊編號驗證是否填入數字
	if(document.form.Member_No.value!=''){  	
  	<%call CheckNumberJ("Member_No","捐款人編號")%>
	}
	if(document.form.Donor_Id_Old.value!=''){  	
  	<%call CheckNumberJ("Donor_Id_Old","舊編號")%>
	}
	
  document.form.target="detail"
  document.form.action.value="fuzzy_query"
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
function IsAbroad_OnClick(){
  if(document.form.IsAbroad.checked){
    donor_address.style.display='none';
    document.form.City.value='';
    ClearOption(document.form.Area);
    document.form.Area.options[0] = new Option('鄉鎮市區','');
    donor_abroad_address.style.display='block';
  }else{
    donor_address.style.display='block';
    donor_abroad_address.style.display='none';
  }
}
function Mobile_OnClick(){
  if(confirm('您是否確定要將查詢結果匯出手機名單？')){
    document.form.target='main';
    document.form.action.value='mobile';
    document.form.submit();
  }
}
function EMail_OnClick(){
  if(confirm('您是否確定要將查詢結果匯出EMail名單？')){
    document.form.target='main';
    document.form.action.value='email';
    document.form.submit();
  }
}
--></script>