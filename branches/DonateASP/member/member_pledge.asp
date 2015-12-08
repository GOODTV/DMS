<!--#include file="../include/dbfunctionJ.asp"-->
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>固定轉帳授權書資料維護</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <form name="form" method="post" action="pledge_list.asp" target="detail">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">固定轉帳授權書</td>
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
                  <table width="100%" border="0" cellpadding="1" cellspacing="3">
                    <tr>
                      <td class="td02-c" width="9%" align="right">機構：</td>
                      <td class="td02-c" width="15%">
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
                      <td class="td02-c" width="9%" align="right">捐款人：</td>
                      <td class="td02-c" width="15%"><input type="text" name="Donor_Name" size="15" class="font9" maxlength="40" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'></td>
                      <td class="td02-c" width="14%" align="right">捐款授權到期日：</td>
                      <td class="td02-c" width="18%">
                      <%
                        Response.Write "<SELECT Name='From_Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        Response.Write "<OPTION>" & " " & "</OPTION>"
                        For I= Year(Date()) To Year(Date())+10
                          Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
                        Next
                        Response.Write "</SELECT>&nbsp;年&nbsp;"
                        Response.Write "<SELECT Name='From_Month' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        Response.Write "<OPTION>" & " " & "</OPTION>"
                        For I= 1 To 12
                          If Len(I)=1 Then
                            Response.Write "<OPTION value='0"&I&"'>0"&I&"</OPTION>"
                          Else
                            Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
                          End If
                        Next
                        Response.Write "</SELECT>&nbsp;月"
                      %>
                      </td>
                      <td class="td02-c" width="20%" align="right"><input type="button" value="  授權到期通知  " name="from" class="addbutton" style="cursor:hand" onClick="ToDay_OnClick()"></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款方式：</td>
                      <td class="td02-c">
                      <%
                        SQL="Select Donate_Payment=CodeDesc From CASECODE Where CodeType='Payment' And CodeDesc Like '%授權書%' Order By Seq"
                        FName="Donate_Payment"
                        Listfield="Donate_Payment"
                        menusize="1"
                        BoundColumn=""
                        call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                      %>
                      </td>
                      <td class="td02-c" align="right">捐款用途：</td>
                      <td class="td02-c">
                      <%
                        SQL="Select Donate_Purpose=CodeDesc From CASECODE Where CodeType='Purpose' And CodeKind='D' Order By Seq"
                        FName="Donate_Purpose"
                        Listfield="Donate_Purpose"
                        menusize="1"
                        BoundColumn=""
                        call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                      %>
                      </td>
                      <td class="td02-c" align="right">信用卡有效月年：</td>
                      <td class="td02-c">
                      <%
                        Response.Write "<SELECT Name='Valid_Month' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        Response.Write "<OPTION>" & " " & "</OPTION>"
                        For I= 1 To 12
                          If Len(I)=1 Then
                            Response.Write "<OPTION value='0"&I&"'>0"&I&"</OPTION>"
                          Else
                            Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
                          End If
                        Next
                        Response.Write "</SELECT>&nbsp;月&nbsp;"
                        Response.Write "<SELECT Name='Valid_Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        Response.Write "<OPTION>" & " " & "</OPTION>"
                        For I= Year(Date()) To Year(Date())+10
                          Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
                        Next
                        Response.Write "</SELECT>&nbsp;年"
                      %>
                      </td>
                      <td align="right"><input type="button" value="信用卡到期通知" name="valid" class="addbutton" style="cursor:hand" onClick="Valid_OnClick()"></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">轉帳日期：</td>
                      <td class="td02-c"><%call Calendar("Donate_Date","")%>
                      </td>
                      <td class="td02-c" align="right">轉帳週期：</td>
                      <td class="td02-c" colspan="4">
                      	<SELECT Name='Donate_Period' size='1' style='font-size: 9pt; font-family: 新細明體'>
                      		<OPTION> </OPTION>
                      		<OPTION  value='單筆'>單筆</OPTION>
                      		<OPTION  value='月繳'>月繳</OPTION>
                      		<OPTION  value='隔月繳'>隔月繳</OPTION>
                      		<OPTION  value='季繳'>季繳</OPTION>
                      		<OPTION  value='半年繳'>半年繳</OPTION>
                      		<OPTION  value='年繳'>年繳</OPTION>
                      	</SELECT>
                      	&nbsp;&nbsp;&nbsp;&nbsp;狀態：
                      	<SELECT Name='Status' size='1' style='font-size: 9pt; font-family: 新細明體'>
                      		<OPTION  value='授權中'>授權中</OPTION>
                      		<OPTION  value='停止'>停止</OPTION>
                      	</SELECT>
                      	&nbsp;&nbsp;&nbsp;&nbsp;經手人：
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
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <input type="button" value=" 查 詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">	
                        <input type="button" value=" 列 表 " name="report" class="cbutton" style="cursor:hand" onClick="Report_OnClick()">
                        <input type="button" value=" 匯 出 " name="report" class="cbutton" style="cursor:hand" onClick="Export_OnClick()">
                      </td>
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td class="td02-c" colspan="7" width="100%"><iframe name="detail" src="pledge_list.asp?status=授權中" height="415" width="100%" frameborder="0" scrolling="auto"></iframe></td>
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
function ToDay_OnClick(){
  location.href='pledge_todate_list.asp';
}
function Valid_OnClick(){
  location.href='pledge_valid_list.asp';
}
function Query_OnClick(){
  if(document.form.Donate_Date.value!=''){
    <%call CheckDateJ("Donate_Date","轉帳日期")%>
  }
  document.form.target="detail"
  document.form.action.value="query"
  document.form.submit();
}
function Report_OnClick(){
  if(document.form.Donate_Date.value!=''){
    <%call CheckDateJ("Donate_Date","轉帳日期")%>
  }
  if(confirm('您是否確定要將查詢結果列表？')){
    document.form.target="report"
    document.form.action.value="report"
    window.open('','report','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=800,height=600')
    document.form.submit();
  }
}
function Export_OnClick(){
  if(document.form.Donate_Date.value!=''){
    <%call CheckDateJ("Donate_Date","轉帳日期")%>
  }
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
}
--></script>