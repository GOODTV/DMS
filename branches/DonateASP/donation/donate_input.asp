<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="donate_input"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
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
                      <td class="td02-c" width="85" align="right">捐款人：</td>
                      <td class="td02-c" width="615">
                      	<input type="text" name="Donor_Name" size="13" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}' value="<%=request("Donor_Name")%>">
                        &nbsp;
                      	捐款人編號：
		                    <input type="text" name="Donor_Id" size="13" class="font9" maxlength="10" value="<%=request("Donor_Id")%>" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
		                    <!--20131008 Modify by GoodTV Tanya:Add查詢「舊編號」-->
		                    &nbsp;
		                    舊編號：
		                    <input type="text" name="Donor_Id_Old" size="7" class="font9" maxlength="7" value="<%=request("Donor_Id_Old")%>" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                        <!--
						&nbsp;
                      	會員編號：
		                    <input type="text" name="Member_No" size="13" class="font9" onkeyup="javascript:UCaseMemberNo();" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}' value="<%=request("Member_No")%>">
                        -->
						&nbsp;
                        身份別：
                        <%
                          SQL1="Select Donor_Type=CodeDesc From CASECODE Where codetype='DonorType' Order By Seq"
                          FName="Donor_Type"
                          Listfield="Donor_Type"
                          menusize="1"
                          BoundColumn=request("Donor_Type")
                          call OptionList (SQL1,FName,Listfield,BoundColumn,menusize)
                        %>
                      </td>
                      <td class="td02-c" width="80">
                        <input type="button" value=" 查 詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_Without_Member_No_OnClick()">	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">身分證/統編：</td>
                      <td class="td02-c" >
                      	<input type="text" name="IDNo" size="13" class="font9" maxlength="10" onkeyup="javascript:UCaseIDNO();" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}' value="<%=request("IDNo")%>">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                      	聯絡電話：
		                    <input type="text" name="Tel_Office" size="13" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}' value="<%=request("Tel_Office")%>">
		                    &nbsp;
                      	地址：
                      	<%call CodeArea ("form","City",request("City"),"Area",request("Area"),"Y")%><input type="text" class="font9" name="Address" size="18" maxlength="80" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}' value="<%=request("Address")%>">	
                      </td>
                      <td class="td02-c">	
                        <input type="button" value=" 取 消 " name="report" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" colspan="9" width="100%" align="center">
                      	<!--20131007 Modify by GoodTV Tanya:查詢結果類別改讀者註記-->
                      	<!--20131008 Modify by GoodTV Tanya:Add查詢「舊編號」-->
                      <%
                        If request("action")="query" Then
                          SQL1="Select Donor_Id,編號=(Case When Member_No<>'' Then Member_No Else CONVERT(nvarchar(20),Donor_Id) End),捐款人=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),讀者=Case When IsMember='Y' Then 'V' Else '' End,身份別=Donor_Type,手機=Cellular_Phone,電話日=Tel_Office,EMail=Email,身份證=IDNo,聯絡人=Contactor,通訊地址=(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+Address Else DONOR.ZipCode+A.mValue+Address End End),最近捐款日期=CONVERT(VarChar,Last_DonateDate,111),舊編號=IsNull(Donor_Id_Old,'') " & _
                              "From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where 1=1 "
                          If request("Donor_Name")<>"" Then SQL1=SQL1 & "And (Donor_Name Like N'%"&request("Donor_Name")&"%' Or NickName Like N'%"&request("Donor_Name")&"%' Or Contactor Like N'%"&request("Donor_Name")&"%' Or Invoice_Title Like N'%"&request("Donor_Name")&"%') "
                          If request("Donor_Id")<>"" Then SQL1=SQL1 & "And Donor_Id Like '%"&request("Donor_Id")&"%' "
                          If request("Donor_Id_Old")<>"" Then SQL1=SQL1 & "And Donor_Id_Old Like '%"&request("Donor_Id_Old")&"%' "
                          If request("Member_No")<>"" Then SQL1=SQL1 & "And Member_No Like '%"&request("Member_No")&"%' "
                          If request("Category")<>"" Then SQL1=SQL1 & "And Category = '"&request("Category")&"' "
                          If request("Donor_Type")<>"" Then SQL1=SQL1 & "And Donor_Type Like '%"&request("Donor_Type")&"%' "
                          If request("IDNo")<>"" Then SQL1=SQL1 & "And (IDNo Like '%"&request("IDNo")&"%' Or Invoice_IDNo Like '%"&request("IDNo")&"%') "
                          If request("Tel_Office")<>"" Then SQL1=SQL1 & "And (Tel_Office Like '%"&request("Tel_Office")&"%' Or Tel_Home Like '%"&request("Tel_Office")&"%' Or Cellular_Phone Like '%"&request("Tel_Office")&"%') "
                          If request("City")<>"" Then SQL1=SQL1 & "And (City = '"&request("City")&"' Or Invoice_City = '"&request("City")&"') "
                          If request("Area")<>"" Then SQL1=SQL1 & "And (Area = '"&request("Area")&"' Or Invoice_Area = '"&request("Area")&"') "
                          If request("Address")<>"" Then SQL1=SQL1 & "And (Address Like N'%"&request("Address")&"%' Or Invoice_Address Like N'%"&request("Address")&"%') "
                          SQL1=SQL1&"Order By Donor_Id Desc"
                          Set RS1 = Server.CreateObject("ADODB.RecordSet")
                          RS1.Open SQL1,Conn,1,1
                          If Not RS1.EOF Then
                            FieldsCount = RS1.Fields.Count-1
                            Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
                            Response.Write "  <tr>"
                            Response.Write "    <td width='85%' align='center' style='color:#000080'>** 請選擇您要輸入的捐款人 **</td>"	
                            Response.Write "    <td width='15%' align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'><a href='donor_add.asp?ctype=donate&donor_name2="&request("Donor_Name")&"&tel_office2="&request("Tel_Office")&"&category2="&request("Category")&"&donor_type2="&request("Donor_Type")&"&idno2="&request("IDNo")&"&city2="&request("City")&"&area2="&request("Area")&"&address2="&request("Address")&"' target='main'>新增捐款人資料</a></span></td>"
                            Response.Write "  </tr>"
                            Response.Write "</table>"
                            Response.Write "<table border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
                            Response.Write "  <tr>"
                            For J = 1 To FieldsCount
                              Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J).Name & "</span></font></td>"
                            Next
                            Response.Write "</tr>"
                            While Not RS1.EOF
                              Response.Write "<a href='donate_add.asp?ctype=donate_input&donor_id="&RS1(0)&"' target='main'>"
                              Response.Write "<tr style='cursor:hand' bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
                              For J = 1 To FieldsCount
                                Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(J))&"</span></td>"
                              Next
                              Response.Flush
                              Response.Clear
                              RS1.MoveNext
                              Response.Write "</tr>"
                              Response.Write "</a>"
                            Wend
                            Response.Write "</table>"	
                          Else
                            session("errnumber")=1
                            session("msg")="查無相關資料，請先新增捐款人資料 ！"
                            Response.Redirect "donor_add.asp?ctype=donate&donor_name2="&request("Donor_Name")&"&tel_office2="&request("Tel_Office")&"&category2="&request("Category")&"&donor_type2="&request("Donor_Type")&"&idno2="&request("IDNo")&"&city2="&request("City")&"&area2="&request("Area")&"&address2="&request("Address")&""
                          End If
                          RS1.Close
                          Set RS1=Nothing 
                        Else
                          Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
                          Response.Write "  <tr>"
                          Response.Write "    <td width='85%' align='center' style='color:#ff0000'>** 請先輸入查詢條件 **</td>"	
                          Response.Write "    <td width='15%' align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='donor_add.asp?ctype=donate' target='main'>新增捐款人資料</a></span></td>"
                          Response.Write "  </tr>"
                          Response.Write "</table>"
                        End If
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
<!-- 新增Query_Without_Member_No_OnClick的method給沒有Member_No的query用 -->
<script language="JavaScript"><!--
function Window_OnLoad(){
  if(document.form.City.value!=''&&document.form.Area.value==''){
    ChgCity(document.form.City.value,document.form.Area);
  }
  document.form.Donor_Name.focus();
}
function UCaseMemberNo(){
  document.form.Member_No.value=document.form.Member_No.value.toUpperCase();
}	
function UCaseIDNO(){
  document.form.IDNo.value=document.form.IDNo.value.toUpperCase();
}	
function Query_OnClick(){
  if(document.form.Donor_Name.value==''&&document.form.Donor_Id.value==''&&document.form.Member_No.value==''&&document.form.Donor_Type.value==''&&document.form.IDNo.value==''&&document.form.Tel_Office.value==''&&document.form.City.value==''&&document.form.Area.value==''&&document.form.Address.value==''){
    document.form.Donor_Name.focus();
    alert('請輸入查詢條件！');
    return;
  }
  document.form.action.value='query';
  document.form.submit();
}
function Cancel_OnClick(){
  location.href='donate.asp';
}
function Query_Without_Member_No_OnClick(){
	//20131008Modify by GoodTV Tanya:Add查詢「舊編號」
  if(document.form.Donor_Name.value==''&&document.form.Donor_Id.value==''&&document.form.Donor_Type.value==''&&document.form.IDNo.value==''&&document.form.Tel_Office.value==''&&document.form.City.value==''&&document.form.Area.value==''&&document.form.Address.value==''&&document.form.Donor_Id_Old.vale==''){
    document.form.Donor_Name.focus();
    alert('請輸入查詢條件！');
    return;
  }
  document.form.action.value='query';
  document.form.submit();
}
--></script>