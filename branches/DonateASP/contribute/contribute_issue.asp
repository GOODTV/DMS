<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="issue"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="contribute_issue_list.asp" target="detail">
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
                      <td class="td02-c" width="70" align="right">機構：</td>
                      <td class="td02-c" width="550">
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
                      	領取人：
                      	<input type="text" name="Issue_Processor" size="13" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                        &nbsp;
                        領取用途：
                        <%
                          SQL="Select Issue_Purpose=CodeDesc From CASECODE Where CodeType='Purpose3' Order By Seq"
                          FName="Issue_Purpose"
                          Listfield="Issue_Purpose"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
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
                      <td class="td02-c" width="180">
                        <input type="button" value=" 查 詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">&nbsp;
                        <input type="button" value=" 列 表 " name="report" class="cbutton" style="cursor:hand" onClick="Report_OnClick()">&nbsp;
                        <input type="button" value=" 匯 出 " name="report" class="cbutton" style="cursor:hand" onClick="Export_OnClick()">		
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">領取日期：</td>
                      <td class="td02-c" colspan="2">
                        <%call Calendar("Issue_Date_B",Year(Date())&"/1/1")%> ~ <%call Calendar("Issue_Date_E","")%>
                        &nbsp;
                        領取編號：
                        <input type="text" name="Invoice_No_B" size="11" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'> ~ <input type="text" name="Invoice_No_E" size="11" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                        &nbsp;
                        <input type="checkbox" name="Issue_TypeM" value="M">手開
                        <input type="checkbox" name="Issue_TypeD" value="D">作廢	
                      </td>
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <%
                        Donate_Date_B=Year(Date())&"/1/1"
                        CreateUser=""
                        If Session("user_group")<>"系統管理" And Session("user_group")<>"捐款管理" Then CreateUser=Session("user_name")
                      %>
                      <td class="td02-c" colspan="9" width="100%"><iframe name="detail" src="contribute_issue_list.asp?Dept_Id=<%=Session("dept_id")%>&Donate_Date_B=<%=Donate_Date_B%>&Create_User=<%=CreateUser%>" height="380" width="100%" frameborder="0" scrolling="auto"></iframe></td>
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
  document.form.Issue_Processor.focus();
}	
function Query_OnClick(){
  if(document.form.Issue_Date_B.value!=''){
    <%call CheckDateJ("Issue_Date_B","領取日期起")%>
  }
  if(document.form.Issue_Date_E.value!=''){
    <%call CheckDateJ("Issue_Date_E","領取日期迄")%>
  }
  document.form.target="detail"
  document.form.action.value="query"
  document.form.submit();
}
function Report_OnClick(){
  if(document.form.Issue_Date_B.value!=''){
    <%call CheckDateJ("Issue_Date_B","領取日期起")%>
  }
  if(document.form.Issue_Date_E.value!=''){
    <%call CheckDateJ("Issue_Date_E","領取日期迄")%>
  }
  if(confirm('您是否確定要將查詢結果列表？')){
    document.form.target="report"
    document.form.action.value="report"
    window.open('','report','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=800,height=600')
    document.form.submit();
  }
}

function Export_OnClick(){
  if(document.form.Issue_Date_B.value!=''){
    <%call CheckDateJ("Issue_Date_B","領取日期起")%>
  }
  if(document.form.Issue_Date_E.value!=''){
    <%call CheckDateJ("Issue_Date_E","領取日期迄")%>
  }
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
}
--></script>