<!--#include file="../include/dbfunctionJ.asp"-->
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>讀者資料維護</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <form name="form" method="post" action="member_list.asp" target="detail">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">讀者資料維護</td>
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
                      <td class="td02-c" width="65" align="right">讀者姓名：</td>
                      <td class="td02-c" width="535">
                      	<input type="text" name="Donor_Name" size="11" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                        &nbsp;
                      	讀者編號：
		                    <input type="text" name="Member_No" size="11" class="font9" onKeyUp="javascript:UCaseMemberNo();" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
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
                      <td class="td02-c" width="180">
                        <input type="button" value=" 查 詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">	
                        <input type="button" value=" 列 表 " name="report" class="cbutton" style="cursor:hand" onClick="Report_OnClick()">
                        <input type="button" value=" 匯 出 " name="export" class="cbutton" style="cursor:hand" onClick="Export_OnClick()">	
                      </td>
                    </tr>
                    <tr>
                      <!--<td class="td02-c" align="right">會員別：</td>-->
                      <td class="td02-c" colspan="2">
                        <%
                          SQL="Select Member_Type=CodeDesc From CASECODE Where codetype='MemberType' Order By Seq"
                          FName="Member_Type"
                          Listfield="Member_Type"
                          menusize="1"
                          BoundColumn=""
                          'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        &nbsp;
                         <!--狀態：-->
                        <%
                          SQL="Select Member_Status=CodeDesc From CASECODE Where codetype='MemberStatus' Order By Seq"
                          FName="Member_Status"
                          Listfield="Member_Status"
                          menusize="1"
                          BoundColumn=""
                          'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        &nbsp;
                      	文宣品：
                        
                        <input type="checkbox" name="IsSendEpaper" value="Y">電子報
                       
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">地址：</td>
                      <td class="td02-c" colspan="2">
                      	<table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="75"><input type="checkbox" name="IsAbroad" value="Y" OnClick='IsAbroad_OnClick()'>海外地址</td>
                            <td width="150" id="donor_address" style="display:block"><%call CodeArea ("form","City","","Area","","Y")%></td>
                            <td width="490"><input type="text" class="font9" name="Address" size="48" maxlength="80" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'></td>
                          </tr>
                      	</table>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" colspan="3" width="100%"><iframe name="detail" src="member_list.asp" height="415" width="100%" frameborder="0" scrolling="auto"></iframe></td>
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
function Query_OnClick(){
  document.form.target="detail"
  document.form.action.value="query"
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
  }else{
    donor_address.style.display='block';
  }
}
--></script>