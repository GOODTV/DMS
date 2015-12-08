<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From CONTENT Where Ser_No='"&request("ser_no")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("Content_Desc")=request("Content_Desc")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 ！"
End If
SQL="Select * From CONTENT Where Ser_No='"&request("ser_no")&"'"
call QuerySQL(SQL,RS)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>郵件內容管理 【修改資料】</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
      <table border="0" width="815" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
	        <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25" colspan="2"><font color="#FFFFFF"><b>&nbsp;郵件內容管理 </b> <font size="2">【修改資料】</font></font></td>
        </tr>
         <tr>
	        <td width="100%" class="font11" valign="top">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">
   	            <tr>
   		            <td>
                    <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" bordercolor="#E1DFCE" cellpadding="2">
                      <tr> 
                        <td noWrap align="right" width="10%"><font color="#000080">類別：</font></td>
                        <td valign="center" colspan="5" width="90%">
                        	<input class=font9t size=41 name="Content_SubType" maxlength="80" value="<%=RS("Content_SubType")%>" readonly >
                        </td>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2" colspan="5"></td>
                      </tr>
                      <tr>
                        <td align="right"><font color="#000080">郵件內容：</td>
                        <td valign="center" colspan="5">
                        <!--#include file="../fckeditor/fckeditor.asp" -->
                        <%
                          Dim oFCKeditor
			                    Set oFCKeditor = New FCKeditor
		                      oFCKeditor.BasePath	= "../fckeditor/"
			                    oFCKeditor.Value	= RS("Content_Desc")
		                      oFCKeditor.Create "Content_Desc"
		                    %>
                        </td>
                      </tr>
                      <tr>
                        <td width="100%" colspan="6">
                          <button id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Update_OnClick()'> <img src='../images/update.gIf' width='20' height='20' align='absmiddle'> 修改</button>&nbsp;
                          <button id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Cancel_OnClick()'> <img src='../images/icon6.gIf' width='20' height='15' align='absmiddle'> 離開</button>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </center></div>
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
function Update_OnClick(){
  <%call SubmitJ("update")%>
}
function Cancel_OnClick(){
  location.href='donor_thanks_list.asp';
}
--></script>