<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From WEB_STYLE Where ser_no='"&request("ser_no")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("Style_Type")=request("Style_Type")
  RS("Style_Name")=request("Style_Name")
  RS("Style_Desc")=request("Style_Desc")
  RS("Style_Value")=request("Style_Value")
  RS("Style_Seq")=request("Style_Seq")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 ！"
  Response.Redirect "web_style.asp"
End If 

If request("action")="delete" Then
  SQL="Delete From WEB_STYLE Where ser_no='"&request("ser_no")&"' "
  Set RS=Conn.Execute(SQL)    
  session("errnumber")=1
  session("msg")="資料刪除成功 !"
  Response.Redirect "web_style.asp"
End If

SQL="Select * From WEB_STYLE Where Ser_No='"&request("ser_no")&"'"
call QuerySQL(SQL,RS)
Item="web_style"
Subject=RS("Style_Name")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>系統參數設定【修改資料】</title>
</head>
<body>
  <p><div align="center"><center>
    <form name="form" method="POST" action>
      <input type="hidden" name="action">
      <input type="hidden" name="Qa_SerNo" value="<%=request("qa_serno")%>">
      <div align="center"><center>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
          <tr>
            <td>
              <table width="800"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
                <tr>
                  <td width="5%"></td>
                  <td width="95%">
  		              <table width="40%"  border="0" cellspacing="0" cellpadding="0">
		                  <br>
		                  <tr>
                        <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gIf" width="10" height="10"></td>
                        <td class="table61-bg"><img src="../images/table06_02.gIf" width="1" height="10"></td>
                        <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gIf" width="10" height="10"></td>
                      </tr>
                      <tr>
                        <td class="table62-bg"></td>
                        <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">系統參數設定  【修改資料】</td>
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
	            <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
                <tr>
                  <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gIf" width="10" height="10"></td>
                  <td class="table61-bg"><img src="../images/table06_02.gIf" width="1" height="10"></td>
                  <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gIf" width="10" height="10"></td>
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
                          BoundColumn=RS("Style_Type")
                          menusize="1"
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        </td>
                        <td width="10%"align="right"><font color="#000080">控制項：</font></td>
                        <td width="20%"><input type="text" name="Style_Name" size="20" class="font9" maxlength="20" value="<%=RS("Style_Name")%>"></td>
                        <td width="13%" align="right"><font color="#000080">控制項說明：</font></td>
                        <td width="35%"><input type="text" name="Style_Desc" size="40" class="font9" maxlength="100" value="<%=RS("Style_Desc")%>"></td>	
                      </tr>
                      <tr>
                        <td align="right"><font color="#000080">設定值：</font></td>
                        <td colspan="3"><input type="text" name="Style_Value" size="48" class="font9" maxlength="50" value="<%=RS("Style_Value")%>"> </td>
                        <td align="right"><font color="#000080">排序：</font></td>
                        <td><input type="text" name="Style_Seq" size="8" class="font9" maxlength="3" value="<%=RS("Style_Seq")%>"></td>
                      </tr>
                      <tr>
                        <td width="99%" colspan="6" bgcolor="#C0C0C0" height="1"> </td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 class="font9" colspan="6" align="left">
              	          <button id='Update' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Update_OnClick()'> <img src='../images/update.gif' width='19' height='20' align='absmiddle'> 修改</button>
              	          <button id='Delete' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Delete_OnClick()'> <img src='../images/delete.gif' width='19' height='20' align='absmiddle'>刪除</button>
              	          <button id='Cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Cancel_OnClick()'> <img src='../images/icon6.gIf' width='20' height='15' align='absmiddle'> 離開</button>
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td class="table63-bg"></td>
                </tr>
                <tr>
                  <td style="background-color:#EEEEE3"><img src="../images/table06_06.gIf" width="10" height="10"></td>
                  <td class="table64-bg"><img src="../images/table06_07.gIf" width="1" height="10"></td>
                  <td style="background-color:#EEEEE3"><img src="../images/table06_08.gIf" width="10" height="10"></td>
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
function Update_OnClick(){
  <%call CheckStringJ("Style_Type","類別")%>
  <%call CheckStringJ("Style_Name","控制項")%>
  <%call ChecklenJ("Style_Name",20,"控制項")%>
  <%call ChecklenJ("Style_Desc",100,"控制項說明")%>
  <%call CheckStringJ("Style_Value","設定值")%>
  <%call ChecklenJ("Style_Value",50,"設定值")%>
  <%call CheckStringJ("Style_Seq","排序")%>
  <%call CheckNumberJ("Style_Seq","排序")%>
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  location.href='web_style.asp';
}
--></script>