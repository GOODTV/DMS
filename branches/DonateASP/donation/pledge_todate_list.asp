<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="pledge_todate"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="pledge_todate_rpt.asp" target="report">
      <input type="hidden" name="action">
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
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' Order By Comp_Label,Dept_Id"
                          Else
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id='"&Session("dept_id")&"' Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
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
                      <td class="td02-c" align="right" width="120">授權到期年月：</td>
                      <td class="td02-c" width="660">
                      <%
                        YY1=Cint(Year(Date()))
                        MM1=Cint(Month(Date()))+1
                        If MM1=13 Then
                          MM1=1
                          YY1=Cint(Year(Date()))+1
                        End If
                        
                        Response.Write "<SELECT Name='To_Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        Response.Write "<OPTION>" & " " & "</OPTION>"
                        For I= Year(Date()) To Year(Date())+10
                          If I=YY1 Then
                            Response.Write "<OPTION value='"&I&"' selected >"&I&"</OPTION>"
                          Else
                            Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
                          End If  
                        Next
                        Response.Write "</SELECT>&nbsp;年&nbsp;"
                        
                        Response.Write "<SELECT Name='To_Month' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                        Response.Write "<OPTION>" & " " & "</OPTION>"
                        For I= 1 To 12
                          If Len(I)=1 Then
                            If I=MM1 Then
                              Response.Write "<OPTION value='0"&I&"' selected >0"&I&"</OPTION>"
                            Else
                              Response.Write "<OPTION value='0"&I&"'>0"&I&"</OPTION>"
                            End If
                          Else
                            If I=MM1 Then
                              Response.Write "<OPTION value='"&I&"' selected >"&I&"</OPTION>"
                            Else
                              Response.Write "<OPTION value='"&I&"'>"&I&"</OPTION>"
                            End If  
                          End If
                        Next
                        Response.Write "</SELECT>&nbsp;月"
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">授權狀態：</td>
                      <td class="td02-c">
                      	<SELECT Name='Status' size='1' style='font-size: 9pt; font-family: 新細明體'>
                      		<OPTION  value=''></OPTION>
                      		<OPTION  value='授權中' selected >授權中</OPTION>
                      		<OPTION  value='停止'>停止</OPTION>
                      	</SELECT>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">授權類別：</td>
                      <td class="td02-c">
                      	<SELECT Name='Donate_Type' size='1' style='font-size: 9pt; font-family: 新細明體'>
                      		<OPTION  value=''></OPTION>
                      		<OPTION  value='長期捐款' selected >長期捐款</OPTION>
                      		<OPTION  value='單次捐款'>單次捐款</OPTION>
                      	</SELECT>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">授權捐款人：</td>
                      <td class="td02-c">
                      	<input type="text" name="Donor_Name" size="10" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Report_OnClick();}'>&nbsp;&nbsp;&nbsp;
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">授權編號：</td>
                      <td class="td02-c"><input type="text" name="Pledge_Id_Begin" size="10" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Report_OnClick();}'> ~ <input type="text" name="Pledge_Id_End" size="10" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Report_OnClick();}'></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">到期信函內容：</td>
                      <td class="td02-c">
                      <%
                        EmailMgr_Id=""
                        SQL1="Select * From EMAILMGR Where EmailMgr_Type='固定捐款到期通知' Order By Ser_No Desc"
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        If Not RS1.EOF Then
                          If RS1.Recordcount=1 Then EmailMgr_Id=RS1("Ser_No") 
                        End If
                        RS1.Close
                        Set RS1=Nothing
                        
                        SQL2=""
                        If Session("dept_id")="C001" Or Session("dept_id")="C002" Then
                          SQL2="Select * From EMAILMGR Where EmailMgr_Type='固定捐款到期通知' And EmailMgr_Subject Like '%北%' Order By Ser_No Desc"
                        ElseIf Session("dept_id")="C003" Or Session("dept_id")="C006" Then
                          SQL2="Select * From EMAILMGR Where EmailMgr_Type='固定捐款到期通知' And EmailMgr_Subject Like '%北%' Order By Ser_No Desc"
                        ElseIf Session("dept_id")="C004" Or Session("dept_id")="C007" Then
                          SQL2="Select * From EMAILMGR Where EmailMgr_Type='固定捐款到期通知' And EmailMgr_Subject Like '%中%' Order By Ser_No Desc"
                        ElseIf Session("dept_id")="C008" Or Session("dept_id")="C009" Then
                          SQL2="Select * From EMAILMGR Where EmailMgr_Type='固定捐款到期通知' And EmailMgr_Subject Like '%南%' Order By Ser_No Desc"
                        End If
                        If SQL2<>"" Then
                          Set RS2 = Server.CreateObject("ADODB.RecordSet")
                          RS2.Open SQL2,Conn,1,1
                          If Not RS2.EOF Then EmailMgr_Id=RS2("Ser_No")
                          RS2.Close
                          Set RS2=Nothing
                        End If
                        
                        SQL="Select EmailMgr_Id=Ser_No,EmailMgr_Subject From EMAILMGR Where EmailMgr_Type='固定捐款到期通知' Order By Ser_No Desc"
                        FName="EmailMgr_Id"
                        Listfield="EmailMgr_Subject"
                        menusize="1"
                        BoundColumn=EmailMgr_Id
                        call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                      %>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">名條內容：</td>
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="Print_Desc" id="Print_Desc1" value="1" checked >收據地址&nbsp;
                        <input type="radio" name="Print_Desc" id="Print_Desc2" value="2">通訊地址	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">名條格式：</td>
                      <td class="td02-c" colspan="3">
                      <%
                        SQL1="Select Label From DEPT Where Dept_Id='"&Session("dept_id")&"' "
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        If Not RS1.EOF Then Label=RS1("Label")
                        RS1.Close
                        Set RS1=Nothing
                        
                        SQL1="Select Seq,Label=CodeDesc From CASECODE Where CodeType='Label' Order By Seq"
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        Response.Write "<SELECT Name='Label' size='1' style='font-size: 9pt; font-family: 新細明體'>" 
                        Response.Write "<OPTION>" & " " & "</OPTION>"
                        While Not RS1.EOF
                          If Cstr(RS1("Seq"))=Cstr(Label) Then
                            Response.Write "<OPTION selected value='"&RS1("Seq")&"'>"&RS1("Label")&"</OPTION>"
                          Else
                            Response.Write "<OPTION value='"&RS1("Seq")&"'>"&RS1("Label")&"</OPTION>"
                          End If
                          RS1.MoveNext
                        Wend
                        Response.Write "</SELECT>"
                        RS1.Close
                        Set RS1=Nothing
                      %>	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">排序方式：</td>
                      <td class="td02-c" colspan="3">
                        <input type="radio" name="OrderBy_Type" id="OrderBy_Type1" value="1" checked >郵遞區號&nbsp;
                        <input type="radio" name="OrderBy_Type" id="OrderBy_Type2" value="2">捐款人編號&nbsp;
                        <input type="radio" name="OrderBy_Type" id="OrderBy_Type3" value="3">授權編號	
                      </td>
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
                        Response.Write "<button id='report' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Report_OnClick()'><img src='../images/print.gif' width='20' height='20'><br>到期函</button>&nbsp;"
                        Response.Write "<button id='address' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Address_OnClick()'><img src='../images/print.gif' width='20' height='20'><br>名條</button>&nbsp;"	
    			              Response.Write "<button id='view' style='position:relative;left:30;width:60;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='View_OnClick()'><img src='../images/search.gif' width='20'><br>樣本預覽</button>"	
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
<script language="JavaScript"><!--
function Window_OnLoad(){
  document.form.Donor_Name.focus();
}	
function Report_OnClick(){
  <%call CheckStringJ("EmailMgr_Id","到期信函內容")%>
  if(confirm('※請注意：地此若為空白到期信函將不會列印\n\n您是否確定要將查詢結果列印？')){
    document.form.target='report';
    document.form.action.value='report';
    window.open('','report','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600');
    document.form.submit();
  }
}
function Address_OnClick(){
  <%call CheckStringJ("Label","名條格式")%>
  if(confirm('您是否確定要列印地址名條？')){
    document.form.target='report';
    document.form.action.value='address';;
    window.open('','report','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600');
    document.form.submit();
  }
}
function View_OnClick(){
  <%call CheckStringJ("EmailMgr_Id","到期信函內容")%>
  window.open('pledge_todate_view.asp?id='+document.form.EmailMgr_Id.value+'','view','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600');
}
--></script>