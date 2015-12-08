<!--#include file="../include/dbfunctionJ.asp"-->
<%
  Donate_Date_End=DateAdd("D",-1,Cdate(Year(Date())&"/"&Month(Date())&"/1"))
  Donate_Date_Begin=Year(Donate_Date_End)&"/"&Month(Donate_Date_End)&"/1"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>會訊芳名錄</title>
  <link rel="stylesheet" type="text/css" href="../include/dms.css">
</head>
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="madazine_list_rpt.asp" target="main">
      <input type="hidden" name="action">
      <input type="hidden" name="Col" value="4">
      <input type="hidden" name="keyword" value=",">
      <input type="hidden" name="maxlen" value="7">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">會訊芳名錄</td>
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
                      <td class="td02-c" align="right" width="100">機構：</td>
                      <td class="td02-c" width="680">
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
                      <td class="td02-c" align="right">捐款日期：</td>
                      <td class="td02-c"><%call Calendar("Donate_Date_Begin",Donate_Date_Begin)%> ~ <%call Calendar("Donate_Date_End",Donate_Date_End)%></td>
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
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">人名顯示：</td>
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="Name_Type" id="Name_Type1" value="1" checked >匿名優先收據抬頭次之
                        <input type="radio" name="Name_Type" id="Name_Type2" value="2">匿名優先捐款人次之</font>&nbsp;
                        <input type="radio" name="Name_Type" id="Name_Type3" value="3">收據抬頭&nbsp;
                        <input type="radio" name="Name_Type" id="Name_Type4" value="4">捐款人</font>	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">多筆捐款：</td>
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="Amt_Type" id="Amt_Type1" value="1" checked >金額分計(100*2,300,500)
                        <input type="radio" name="Amt_Type" id="Amt_Type2" value="2">金額自動加總(1,000)</font>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款用途：</td>
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="Purpose_Type" id="Purpose_Type1" value="1" checked >依用途不同分開計算
                        <input type="radio" name="Purpose_Type" id="Purpose_Type2" value="2">合併計算</font>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">金額排序：</td>
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="Seq_Type" id="Seq_Type1" value="Asc" checked >由低至高
                        <input type="radio" name="Seq_Type" id="Seq_Type2" value="Desc">由高至低</font>
                      </td>
                    </tr>                    
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%"><button id='export' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Txt_OnClick()'><img src='../images/notepad.gif' width='14'><br>匯出</button></td>
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
function Report_OnClick(){
  if(document.form.Donate_Date_Begin.value!=''){
    <%call CheckDateJ("Donate_Date_Begin","捐款日期起")%>
  }
  if(document.form.Donate_Date_End.value!=''){
    <%call CheckDateJ("Donate_Date_End","捐款日期迄")%>
  }
  if(confirm('您是否確定要將查詢結果列表？')){
    document.form.target="report"
    document.form.action.value="report"
    window.open('','report','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600')
    document.form.submit();
  }
}

function Txt_OnClick(){
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='txt';
    document.form.submit();
  }
}
--></script>