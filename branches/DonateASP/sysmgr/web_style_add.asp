<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  SQL="WEB_STYLE"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS.Addnew
  RS("Style_Type")=request("Style_Type")
  RS("Style_Name")=request("Style_Name")
  RS("Style_Desc")=request("Style_Desc")
  RS("Style_Value")=request("Style_Value")
  RS("Style_Seq")=request("Style_Seq")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料新增成功 ！"
  SQL="Select @@IDENTITY as ser_no"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF then
    response.redirect "web_style_edit.asp?ser_no="&RS("ser_no")
  End If
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>系統參數設定【新增資料】</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="POST" action>
      <input type="hidden" name="action">
      <div align="center"><center>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
          <br>
          <tr>
            <td>
              <table width="800"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
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
                        <td class="table62-bg"></td>
                        <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">系統參數設定  【新增資料】</td>
                        <td class="table63-bg"></td>
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
                  <td class="table62-bg"></td>
                  <td valign="top">
                    <table width="100%" border="1" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#E1DFCE">
                      <tr>
                        <td width="10%"align="right"><font color="#000080">類別：</font></td>
                        <td width="12%">
                        <%
                          SQL="Select Distinct Style_Type From WEB_STYLE Where Style_Type<>'MD5'"
                          FName="Style_Type"
                          Listfield="Style_Type"
                          BoundColumn=""
                          menusize="1"
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        </td>
                        <td width="10%"align="right"><font color="#000080">控制項：</font></td>
                        <td width="20%"><input type="text" name="Style_Name" size="20" class="font9" maxlength="20"></td>
                        <td width="13%" align="right"><font color="#000080">控制項說明：</font></td>
                        <td width="35%"><input type="text" name="Style_Desc" size="40" class="font9" maxlength="100"></td>	
                      </tr>
                      <tr>
                        <td align="right"><font color="#000080">設定值：</font></td>
                        <td colspan="3"><input type="text" name="Style_Value" size="48" class="font9" maxlength="50"> </td>
                        <td align="right"><font color="#000080">排序：</font></td>
                        <td><input type="text" name="Style_Seq" size="8" class="font9" maxlength="3"></td>
                      </tr>                      
                      <tr>
                        <td width="99%" colspan="6" bgcolor="#C0C0C0" height="1"> </td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 class="font9" colspan="6" align="left"><%SaveButton%></td>
                      </tr>
                    </table>
                  </td>
                  <td class="table63-bg"></td>
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
      </center></div>
    </form>
  </center></div>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Save_OnClick(){
  <%call CheckStringJ("Style_Type","類別")%>
  <%call CheckStringJ("Style_Name","控制項")%>
  <%call ChecklenJ("Style_Name",20,"控制項")%>
  <%call ChecklenJ("Style_Desc",100,"控制項說明")%>
  <%call CheckStringJ("Style_Value","設定值")%>
  <%call ChecklenJ("Style_Value",50,"設定值")%>
  <%call CheckStringJ("Style_Seq","排序")%>
  <%call CheckNumberJ("Style_Seq","排序")%>
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  location.href='web_style.asp';
}
--></script>