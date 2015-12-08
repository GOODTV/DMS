<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  SQL1="ACT"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  RS1.Addnew
  RS1("Dept_Id")=request("Dept_Id")
  RS1("Act_Name")=request("Act_Name")
  RS1("Act_ShortName")=request("Act_ShortName")
  RS1("Act_OrgName")=request("Act_OrgName")
  RS1("Act_OrgName2")=request("Act_OrgName2")
  RS1("Act_Subject")=request("Act_Subject")
  If request("Act_BeginDate")<>"" Then
    RS1("Act_BeginDate")=request("Act_BeginDate")
  Else
    RS1("Act_BeginDate")=null
  End If
  If request("Act_EndDate")<>"" Then
    RS1("Act_EndDate")=request("Act_EndDate")
  Else
    RS1("Act_EndDate")=null
  End If
  RS1("Act_Licence")=request("Act_Licence")
  RS1("Act_Flg")=request("Act_Flg")
  RS1("Act_Flg2")=request("Act_Flg2")
  RS1("Act_Pre")=request("Act_Pre")
  RS1("Act_Rule_Type")=request("Act_Rule_Type")
  RS1("Act_Rule_YMD")=request("Act_Rule_YMD")
  RS1("Act_Rule_Len")=request("Act_Rule_Len")
  RS1("Act_Rule_Pub")=request("Act_Rule_Pub")
  RS1("Act_Invoice")=request("Act_Invoice")
  RS1("Act_Pre2")=request("Act_Pre2")
  RS1("Act_Rule_Type2")=request("Act_Rule_Type2")
  RS1("Act_Rule_YMD2")=request("Act_Rule_YMD2")
  RS1("Act_Rule_Len2")=request("Act_Rule_Len2")
  RS1("Act_Rule_Pub2")=request("Act_Rule_Pub2")  
  RS1("Remark")=request("Remark")
  RS1("Core_Value")=request("Core_Value")
  RS1("Act_MainPower")=request("Act_MainPower")
  RS1("OutCome")=request("OutCome")
  If request("Seq")<>"" Then
    RS1("Seq")=request("Seq")
  Else
    RS1("Seq")="999"
  End If
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="勸募活動新增成功 ！"
  SQL="Select @@IDENTITY as ser_no"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then
    Response.Redirect "act_edit2.asp?act_id="&RS("ser_no") 
  End If
End If

IsAct="N"
Act_Flg="N"
SQL1="Select IsAct,Donate_Invoice From DEPT Where Dept_Id='"&Session("dept_id")&"'"
Call QuerySQL(SQL1,RS1)
If RS1("IsAct")="Y" Then
  IsAct="Y"
  Act_Flg="Y"
End If
Donate_Invoice=RS1("Donate_Invoice")
RS1.Close
Set RS1=Nothing

Seq=1
SQL1="Select Seq=Isnull(Max(Seq),0) From ACT"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then
  Seq=Cint(RS1("Seq"))+1
End If
RS1.Close
Set RS1=Nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>新增專案活動</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="IsAct" value="<%=IsAct%>">
      <input type="hidden" name="Act_Flg" value="<%=Act_Flg%>">	
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">新增專案活動</td>
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
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td align="right">主責單位：</td>
                            <td>
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
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                  　        </td>
                            <td align="right">主責人：</td>
                            <td align="left"><input type="text" name="Act_MainPower" size="35" class="font9" maxlength="60"></td>
                          </tr>
                          <tr>
                            <td align="right" width="140">活動名稱：</td>
                            <td align="left" width="230"><input type="text" name="Act_Name" size="35" class="font9" maxlength="60"></td>
                            <td align="right" width="80">活動簡稱：</td>
                            <td align="left" width="330"><input type="text" name="Act_ShortName" size="35" class="font9" maxlength="30"></td>
                          </tr>
                          <tr>
                            <td align="right">主辦單位：</td>
                            <td align="left"><input type="text" name="Act_OrgName" size="35" class="font9" maxlength="60"></td>
                            <td align="right">協辦單位：</td>
                            <td align="left"><input type="text" name="Act_OrgName2" size="35" class="font9" maxlength="60"></td>
                          </tr>
                          <tr>
                            <td align="right"">活動主題：</td>
                            <td align="left"><input type="text" name="Act_Subject" size="35" class="font9" maxlength="60"></td>
                            <td align="right">活動期間：</td>
                            <td align="left"><%call Calendar("Act_BeginDate",Date())%> ~ <%call Calendar("Act_EndDate","")%></td>
                          </tr>
                          <tr>
                            <td align="right">勸募許可文號：</td>
                            <td align="left"><input type="text" name="Act_Licence" size="35" class="font9" maxlength="80"></td>
                            <td align="right">核心價值：</td>
                            <td align="left">
                            <%
                              SQL="select CodeDesc as Core_Value from casecode where codetype='Core_Value' order by Seq"
                              FName="Core_Value"
                              Listfield="Core_Value"
                              menusize="1"
                              BoundColumn=""
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">方案結果：</td>
                            <td align="left">
                              <select name='outcome'>
                              	<option value=''> </option>
                              	<option value='執行'>執行</option>
                              	<option value='結案'>結案</option>
                              </select>
                            </td>
                            <td align="right"">排序：</td>
                            <td align="left"><input type="text" name="Seq" size="8" class="font9" maxlength="3" value="<%=Seq%>"></td>
                          </tr>
                          <%If IsAct="Y" Then%>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">捐款收據取號規則：</td>
                            <td align="left" colspan="3">
                              <input type="radio" name="Act_Flg2" id="Act_Flg2_1" value="Y" checked >自訂取號規則<font color="#ff0000">(請填寫下方資料)</font>&nbsp;
                              <input type="radio" name="Act_Flg2" id="Act_Flg2_2" value="N">依現有捐款取號規則
                            </td>
                          </tr>
                          <tr>
                            <td align="right">捐款收據前置碼：</td>
                            <td align="left" colspan="3">
                              <input type="text" name="Act_Pre" size="10" class="font9" maxlength="5" onkeyup="javascript:UCasePre();">
                              &nbsp;&nbsp;
                              收據取號規則<span lang="en-us">：</span>
                              <select size="1" name="Act_Rule_Type" OnChange='JavaScript:Chg_RuleType(this.value);'>
		                            <option value="R">民國年</option>
		                            <option value="A" selected >西元年</option>
		                            <option value="S">流水號</option>
		                          </select>
		                          <span lang="en-us">+</span>
                              <select size="1" name="Act_Rule_YMD">
		                            <option value="Y">年序號</option>
		                            <option value="M" selected >月序號</option>
		                            <option value="D">日序號</option>
		                          </select>
		                          <span lang="en-us">+</span>
		                          流水號長度<span lang="en-us">：</span>
		                          <input type="text" name="Act_Rule_Len" size="2" class="font9" maxlength="2" value="5">
                              &nbsp;&nbsp;
                              多單位流水號取號<span lang="en-us">：</span>
                              <select size="1" name="Act_Rule_Pub">
		                            <option value="Y" selected >共用</option>
		                            <option value="N">獨立</option>
		                          </select>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">捐款年度收據取號規則：</td>
                            <td align="left" colspan="3">
                              <select size="1" name="Act_Invoice">
		                            <option value="Y" <%If Donate_Invoice="Y" Then%>selected<%End If%> >每筆捐款均產生獨立收據編號</option>
		                            <option value="N" <%If Donate_Invoice="N" Then%>selected<%End If%> >所有捐款匯整成一個收據編號</option>
		                          </select>
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">物品捐贈收據前置碼：</td>
                            <td align="left" colspan="3">
                              <input type="text" name="Act_Pre2" size="10" class="font9" maxlength="5" onkeyup="javascript:UCasePre2();">
                              &nbsp;&nbsp;
                              收據取號規則<span lang="en-us">：</span>
                              <select size="1" name="Act_Rule_Type2" OnChange='JavaScript:Chg_RuleType2(this.value);'>
                              	<option value="P" selected >同捐款</option>
		                            <option value="R">民國年</option>
		                            <option value="A">西元年</option>
		                            <option value="S">流水號</option>
		                          </select>
		                          <span lang="en-us">+</span>
                              <select size="1" name="Act_Rule_YMD2">
                              	<option value=""> </option> 
		                            <option value="Y">年序號</option>
		                            <option value="M" selected >月序號</option>
		                            <option value="D">日序號</option>
		                          </select>
		                          <span lang="en-us">+</span>
		                          流水號長度<span lang="en-us">：</span>
		                          <input type="text" name="Act_Rule_Len2" size="2" class="font9" maxlength="2" value="5">
                              &nbsp;&nbsp;
                              多單位流水號取號<span lang="en-us">：</span>
                              <select size="1" name="Act_Rule_Pub2">
		                            <option value="Y" selected >共用</option>
		                            <option value="N">獨立</option>
		                          </select>
                            </td>
                          </tr>
                          <tr>
                            <td colspan="4">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="IsSameInvoice" value="Y" OnClick="IsSameInvoice_OnClick()"><font color="#ff0000">同捐款取號規則</font></td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <%Else%>
                          <input type="hidden" name="Act_Flg2" value="N">
                          <input type="hidden" name="Act_Pre" value="">
                          <input type="hidden" name="Act_Rule_Type" value="A">
                          <input type="hidden" name="Act_Rule_YMD" value="M">
                          <input type="hidden" name="Act_Rule_Len" value="5">
                          <input type="hidden" name="Act_Rule_Pub" value="Y">
                          <input type="hidden" name="Act_Invoice" value="Y">	
                          <input type="hidden" name="Act_Pre2" value="">
                          <input type="hidden" name="Act_Rule_Type2" value="A">
                          <input type="hidden" name="Act_Rule_YMD2" value="M">
                          <input type="hidden" name="Act_Rule_Len2" value="5">
                          <input type="hidden" name="Act_Rule_Pub2" value="Y">
                          <%End If%>
                          <tr>
                            <td align="right">備註：</td>
                            <td align="left" colspan="3">
                            	<textarea name="Remark" rows="6" cols="81" class="font9"></textarea>
                            </td>
                          </tr>
                          <!--#include file="../include/calendar2.asp"-->
                          <tr>
                            <td align="center" colspan="4">
                               <input type="button" value=" 存 檔 " name="save" class="cbutton" style="cursor:hand" onClick="Save_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value=" 取 消 " name="cancel" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">
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
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function UCasePre(){
  document.form.Act_Pre.value=document.form.Act_Pre.value.toUpperCase();
}
function Chg_RuleType(){
  if(document.form.Act_Rule_Type.value=='S'){
    document.form.Act_Rule_YMD.value='';
  }else{
    document.form.Act_Rule_YMD.value='M';
  }
}
function UCasePre2(){
  document.form.Act_Pre2.value=document.form.Act_Pre2.value.toUpperCase();
}
function Chg_RuleType2(){
  if(document.form.Act_Rule_Type2.value=='S'){
    document.form.Act_Rule_YMD2.value='';
  }else{
    document.form.Act_Rule_YMD2.value='M';
  }
}
function IsSameInvoice_OnClick(){
  if(document.form.IsSameInvoice.checked){
    document.form.Act_Pre2.value=document.form.Act_Pre.value;
    document.form.Act_Rule_Type2.value='P';
    document.form.Act_Rule_YMD2.value=document.form.Act_Rule_YMD.value;
    document.form.Act_Rule_Len2.value=document.form.Act_Rule_Len.value;
    document.form.Act_Rule_Pub2.value=document.form.Act_Rule_Pub.value;
  }
}
function Save_OnClick(){
  <%call CheckStringJ("Act_Name","活動名稱")%>
  <%call ChecklenJ("Act_Name",60,"活動名稱")%>
  <%call CheckStringJ("Act_ShortName","活動簡稱")%>
  <%call ChecklenJ("Act_ShortName",60,"活動簡稱")%>
  <%call ChecklenJ("Act_OrgName",60,"主辦單位")%>
  <%call ChecklenJ("Act_OrgName2",60,"協辦單位")%>
  <%call ChecklenJ("Act_Subject",60,"活動主題")%>
  <%call CheckStringJ("Act_BeginDate","活動起日")%>
  if(document.form.Act_EndDate.value==''){
    document.form.Act_EndDate.value='2099/12/31';
  }
  <%call DiffDateJ("Act_BeginDate","Act_EndDate")%> 
  <%call ChecklenJ("Act_Licence",80,"勸募許可文號")%>
  if(document.form.IsAct.value=='Y'){
    if(document.form.Act_Flg2[0].checked){
      <%call CheckStringJ("Act_Pre","捐款收據前置碼")%>
      <%call ChecklenJ("Act_Pre",5,"捐款收據前置碼")%>
      if(document.form.Act_Rule_Type.value!='S'){
        <%call CheckStringJ("Act_Rule_YMD","序號")%>
      }
      <%call CheckStringJ("Act_Rule_Len","捐款流水號長度")%>
      <%call ChecklenJ("Act_Rule_Len",2,"捐款流水號長度")%>
      <%call CheckStringJ("Act_Pre2","物品捐贈前置碼")%>
      <%call ChecklenJ("Act_Pre2",5,"物品捐贈前置碼")%>
      if(document.form.Act_Rule_Type2.value!='S'){
        <%call CheckStringJ("Act_Rule_YMD2","序號")%>
      }
      <%call CheckStringJ("Act_Rule_Len2","物品捐贈流水號長度")%>
      <%call ChecklenJ("Act_Rule_Len2",2,"物品捐贈流水號長度")%>
    }
  }
  if(document.form.Seq.value==''){
   alert('排序  欄位不可為空白！');
    document.form.Seq.focus();
    return;
  }
  if(isNaN(Number(document.form.Seq.value))==true){
    alert('排序  欄位必須為數字！');
    document.form.Seq.focus();
    return;
  }  
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  location.href='act2.asp';
}
--></script>