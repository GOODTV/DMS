<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="change" Then
  SQL1="Select * From DEPT Where Dept_Id='"&Request("Dept_Id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  Session("sys_name")=RS1("sys_name")
  Session("dept_id")=RS1("dept_id")
  Session("dept_desc")=RS1("dept_desc")
  Session("comp_label")=RS1("comp_label")
  Session("comp_name")=RS1("comp_name")
  Session("comp_ShortName")=RS1("comp_ShortName")
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="切換成功，您已經重新以『"&Session("comp_ShortName")&"』管理者身份登入\n\n請您繼續使用本系統！"
  Response.Redirect "../sysmgr/main_d.asp"
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>切換作業部門</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">切換作業部門</td>
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
                      <td class="td02-c" width="30" align="right"> </td>
                      <td class="td02-c" width="750">
                      	目前負責部門：『&nbsp;<%=Session("comp_ShortName")%>&nbsp;』切換至&nbsp;--&gt
                      	<%
                      	  depttype=""
                      	  Ary_depttype=Split(Session("dept_type"),",")
                      	  For I=0 To UBound(Ary_depttype)
                      	    If Cstr(Session("dept_id"))<>Cstr(Trim(Ary_depttype(I))) Then
                      	      If depttype="" Then
                      	        depttype="'"&Cstr(Trim(Ary_depttype(I)))&"'"
                      	      Else
                      	        depttype=depttype&","&"'"&Cstr(Trim(Ary_depttype(I)))&"'"
                      	      End If
                      	    End If
                      	  Next
                      	  Row=1
                      	  SQL1="Select Dept_Id,Comp_ShortName From DEPT Where Dept_Id In ("&depttype&") Order By Comp_Label,Dept_Id"
                          Set RS1 = Server.CreateObject("ADODB.RecordSet")
                          RS1.Open SQL1,Conn,1,1
                          Response.Write "<select name='Dept_Id' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                          While Not RS1.EOF
                            If Row=1 Then
                              Response.Write "<option selected value='"&RS1("Dept_Id")&"'>"&RS1("Comp_ShortName")&"</OPTION>"
                            Else
                              Response.Write "<option value='"&RS1("Dept_Id")&"'>"&RS1("Comp_ShortName")&"</OPTION>"
                            End If
                            Row=Row+1
                            RS1.MoveNext
                          Wend
                          Response.Write "</select>"
                      	%>
                      	&nbsp;<input type="button" value=" 切換作業部門 " name="save" class="addbutton" style="cursor:hand" OnClick="Change_OnClick()">
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
function Change_OnClick(){
  <%call CheckStringJ("Dept_Id","切換作業部門")%>
  if(confirm('您是否確定要切換作業部門？')){
    document.form.action.value='change';
    document.form.submit();
  }
}
--></script>