<!--#include file="../include/dbfunctionJ.asp"-->
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>會員繳費資料維護</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <form name="form" method="post" action="member_donate_list.asp" target="detail">
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
                      <td class="table62-bg">　</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">會員繳費資料維護</td>
                      <td class="table63-bg">　</td>
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
                <td class="table62-bg">　</td>
                <td valign="top">
                  <table width="100%" border="0" cellpadding="1" cellspacing="1">
                    <tr>
                      <td class="td02-c" width="70" align="right">機構：</td>
                      <td class="td02-c" width="650">
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
                      	&nbsp;
                      	會員姓名：
                      	<input type="text" name="Donor_Name" size="15" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                      	&nbsp;
                      	會員編號：
                      	<input type="text" name="Member_No" size="15" class="font9" onKeyUp="javascript:UCaseMemberNo();" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>	
                        &nbsp;
                        繳費方式：
                        <%
                          SQL="Select Donate_Payment=CodeDesc From CASECODE Where codetype='Payment' Order By Seq"
                          FName="Donate_Payment"
                          Listfield="Donate_Payment"
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
                      <td class="td02-c" width="70" align="right">繳費用途：</td>
                      <td class="td02-c" width="650">
                        <%
                          SQL="Select Donate_Purpose=CodeDesc From CASECODE Where CodeType='Purpose' And CodeKind In ('M','A') Order By Seq"
                          FName="Donate_Purpose"
                          Listfield="Donate_Purpose"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        &nbsp;
                        傳票號碼：
                      	<input type="text" name="Donation_NumberNo" size="11" class="font9" onKeyUp="javascript:UCaseNumberNo();" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>	
                      	&nbsp;
                        劃撥&nbsp;/&nbsp;匯款單號：
                        <input type="text" name="Donation_SubPoenaNo" size="10" class="font9" onKeyUp="javascript:UCaseSubPoenaNo();" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>	
                        &nbsp;
                        經手人： 
                        <%
                          If Session("user_group")<>"系統管理" And Session("user_group")<>"捐款管理" Then
                            SQL="Select Create_User=user_name From userfile Where user_name='"&Session("user_name")&"' Order By dept_id,user_group"
                            FName="Create_User"
                            Listfield="Create_User"
                            menusize="1"
                            BoundColumn=Session("user_name")
                            BoundNull="N"
                            call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,BoundNull)
                          Else
                            If Session("user_id")<>"npois" Then
                              SQL="Select Create_User=user_name From userfile Where user_id<>'npois' Order By dept_id,user_group"
                            Else
                              SQL="Select Create_User=user_name From userfile Order By dept_id,user_group"
                            End If
                            FName="Create_User"
                            Listfield="Create_User"
                            menusize="1"
                            call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                          End If
                        %>                                               
                      </td>
                      <td class="td02-c" width="80">
                        <input type="button" value=" 列 表 " name="report" class="cbutton" style="cursor:hand" onClick="Report_OnClick()">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">繳費日期：</td>
                      <td class="td02-c">
                        <%call Calendar("Donate_Date_B",Year(Date())&"/1/1")%> ~ <%call Calendar("Donate_Date_E","")%>
                        &nbsp;
                        收據開立：
                        <%
                          SQL="Select Invoice_Type=CodeDesc From CASECODE Where codetype='InvoiceType' Order By Seq"
                          FName="Invoice_Type"
                          Listfield="Invoice_Type"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        &nbsp;
                        收據編號：
                        <input type="text" name="Invoice_No_B" size="11" class="font9"> ~ <input type="text" name="Invoice_No_E" size="11" class="font9">
                      </td>
                      <td class="td02-c">
                        <input type="button" value=" 匯 出 " name="report" class="cbutton" style="cursor:hand" onClick="Export_OnClick()">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">沖帳日期：</td>
                      <td class="td02-c">
                        <%call Calendar("Accoun_Date_B","")%> ~ <%call Calendar("Accoun_Date_E","")%>
                        &nbsp;
                        繳費類別：
                        <select size="1" name="Donate_Type" class="font9">
                          <option value=""> </option>
                          <option value="單次捐款">單次捐款</option>
                          <option value="長期捐款">長期捐款</option>
                        </select>
                        &nbsp;
                        入帳銀行：
                        <%
                          SQL="Select Accoun_Bank=CodeDesc From CASECODE Where codetype='Bank' Order By Seq"
                          FName="Accoun_Bank"
                          Listfield="Accoun_Bank"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                      </td>
                      <td class="td02-c"> </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">會計科目：</td>
                      <td class="td02-c">
                        <%
                          SQL="Select Accounting_Title=CodeDesc From CASECODE Where codetype='Accoun' Order By Seq"
                          FName="Accounting_Title"
                          Listfield="Accounting_Title"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                      	&nbsp;
                      	募款活動：
                        <%
                          SQL="Select Act_Id,Act_ShortName From ACT Order By Act_id Desc"
                          FName="Act_Id"
                          Listfield="Act_ShortName"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        &nbsp;
                        <input type="checkbox" name="Issue_TypeM" value="M">手開收據
                        <input type="checkbox" name="Issue_TypeD" value="D">作廢收據
                      </td>
                      <td class="td02-c"> </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">款項用途：</td>
                      <td class="td02-c">
                        <input type='checkbox' name='Donate_Purpose_Type' id="Donate_Purpose_Type2" value="M" checked >會務繳費
                        <input type='checkbox' name='Donate_Purpose_Type' id="Donate_Purpose_Type3" value="A" checked >活動報名		
                      </td>
                      <td class="td02-c"> </td>
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <%
                        Donate_Date_B=Year(Date())&"/1/1"
                        CreateUser=""
                        If Session("user_group")<>"系統管理" And Session("user_group")<>"捐款管理" Then CreateUser=Session("user_name")
                      %>
                      <td class="td02-c" colspan="9" width="100%"><iframe name="detail" src="member_donate_list.asp?Dept_Id=<%=Session("dept_id")%>&Donate_Date_B=<%=Donate_Date_B%>&Create_User=<%=CreateUser%>&Donate_Purpose_Type=M,A" height="380" width="100%" frameborder="0" scrolling="auto"></iframe></td>
                    </tr>
                  </table>
                </td>
                <td class="table63-bg">　</td>
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
function UCaseMemberNo(){
  document.form.Member_No.value=document.form.Member_No.value.toUpperCase();
}
function UCaseNumberNo(){
  document.form.Donation_NumberNo.value=document.form.Donation_NumberNo.value.toUpperCase();
}
function UCaseSubPoenaNo(){
  document.form.Donation_SubPoenaNo.value=document.form.Donation_SubPoenaNo.value.toUpperCase();
}
function Query_OnClick(){
  if(document.form.Donate_Date_B.value!=''){
    <%call CheckDateJ("Donate_Date_B","捐款日期起")%>
  }
  if(document.form.Donate_Date_E.value!=''){
    <%call CheckDateJ("Donate_Date_E","捐款日期迄")%>
  }
  if(document.form.Accoun_Date_B.value!=''){
    <%call CheckDateJ("Accoun_Date_B","沖帳日期起")%>
  }
  if(document.form.Accoun_Date_E.value!=''){
    <%call CheckDateJ("Accoun_Date_E","沖帳日期迄")%>
  }  
  document.form.target="detail"
  document.form.action.value="query"
  document.form.submit();
}
function Report_OnClick(){
  if(document.form.Donate_Date_B.value!=''){
    <%call CheckDateJ("Donate_Date_B","捐款日期起")%>
  }
  if(document.form.Donate_Date_E.value!=''){
    <%call CheckDateJ("Donate_Date_E","捐款日期迄")%>
  }
  if(document.form.Accoun_Date_B.value!=''){
    <%call CheckDateJ("Accoun_Date_B","沖帳日期起")%>
  }
  if(document.form.Accoun_Date_E.value!=''){
    <%call CheckDateJ("Accoun_Date_E","沖帳日期迄")%>
  }
  if(confirm('您是否確定要將查詢結果列表？')){
    document.form.target="report"
    document.form.action.value="report"
    window.open('','report','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=800,height=600')
    document.form.submit();
  }
}

function Export_OnClick(){
  if(document.form.Donate_Date_B.value!=''){
    <%call CheckDateJ("Donate_Date_B","捐款日期起")%>
  }
  if(document.form.Donate_Date_E.value!=''){
    <%call CheckDateJ("Donate_Date_E","捐款日期迄")%>
  }
  if(document.form.Accoun_Date_B.value!=''){
    <%call CheckDateJ("Accoun_Date_B","沖帳日期起")%>
  }
  if(document.form.Accoun_Date_E.value!=''){
    <%call CheckDateJ("Accoun_Date_E","沖帳日期迄")%>
  }
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
}
--></script>