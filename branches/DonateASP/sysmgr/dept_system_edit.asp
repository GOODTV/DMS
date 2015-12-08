<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From DEPT Where Dept_Id='"&request("dept_id")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("Meta_Title")=request("Meta_Title")
  RS("Meta_Author")=request("Meta_Author")
  RS("Meta_Copyright")=request("Meta_Copyright")
  RS("Meta_Designer")=request("Meta_Designer")
  RS("Meta_Keywords")=request("Meta_Keywords")
  RS("Meta_Description")=request("Meta_Description")
  RS("Meta_Abstract")=request("Meta_Abstract")
  If request("Meta_Foot")<>"" Then
    RS("Meta_Foot")=request("Meta_Foot")
  Else
    RS("Meta_Foot")=""
  End If  
  If request("Meta_Revisit_After")<>"" Then
    RS("Meta_Revisit_After")=request("Meta_Revisit_After")
  Else
    RS("Meta_Revisit_After")=null
  End If
  RS("Meta_Distribution")=request("Meta_Distribution")
  RS("Meta_Robots")=request("Meta_Robots")
  RS("MaxFileSize")=request("MaxFileSize")
  RS("MaxUploadSize")=request("MaxUploadSize")
  RS("Log_Keep")=request("Log_Keep")
  RS("Hit_Keep")=request("Hit_Keep")
  RS("Client_IP")=request("Client_IP")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 !"
End If

SQL="Select * From DEPT Where DEPT_ID='"&request("dept_id")&"'"
call QuerySQL(SQL,RS) 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>機構部門管理</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
  <p><div align="center"><center>
    <form name="form" action="" method="post">
      <input type="hidden" name="Dept_Id" value="<%=request("dept_id")%>">
      <input type="hidden" name="action">
      <div align="center"><center>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
          <tr>
            <td>
              <table width="780"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
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
                        <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">參數設定</td>
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
                  <td height="445" valign="top" class="table62-bg">&nbsp;</td>
                  <td height="445" valign="top">              
                    <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                      <tr>
                        <td class="td02-c" width="100%" valign="top" colspan="3">
                          <table border="1" cellpadding="2" style="border-collapse: collapse" width="100%" height="25" cellspacing="1">
                            <tr>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">機構資訊設定</a></td>
                              <td class="button2-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_system_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool"><font color="#800000"><b>網站參數設定</b></font></a></td>
                              <%If Session("user_id")="npois" And (Request.ServerVariables("REMOTE_HOST")="60.250.147.33" Or Request.ServerVariables("REMOTE_HOST")="127.0.0.1") Then%>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_donate_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">捐款參數設定</a></td>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_content_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">捐款機制說明</a></td>
                              <%Else%>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_content_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">捐款機制說明</a></td>
                              <%End If%>
                            </tr>
                          </table>
                        </td>
                      </tr>
                      <tr>
                        <td valign="top">&nbsp;</td>
                      </tr>
                      <tr>
                        <td class="td02-c" width="100%">
                          <table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">                        
                            <tr>
                              <td width="12%" align="right"><font color="#000080">機構簡稱<span lang="en-us">:</span></font></td>
                              <td width="88%"><input type="text" name="Comp_ShortName" size="35" class="font9t" readonly maxlength="30" value="<%=RS("Comp_ShortName")%>"></td>
                            </tr>
                            <tr>
                              <td width="12%" align="right"><font color="#000080">標題<span lang="en-us">:</span></font></td>
                              <td width="88%">
                                <input type="text" name="Meta_Title" size="35" class="font9" maxlength="100" value="<%=RS("Meta_Title")%>">
                                &nbsp;&nbsp;
                                <font color="#000080">作&nbsp;&nbsp;&nbsp;&nbsp;者<span lang="en-us">:</span></font>
                                <input type="text" name="Meta_Author" size="35" class="font9" maxlength="80" value="<%=RS("Meta_Author")%>">
                              </td>
                            </tr>
                            <tr>
                              <td width="12%" align="right"><font color="#000080">版權<span lang="en-us">:</span></font></td>
                              <td width="88%">
                                <input type="text" name="Meta_Copyright" size="35" class="font9" maxlength="80" value="<%=RS("Meta_Copyright")%>">
                                &nbsp;&nbsp;
                                <font color="#000080">設計者<span lang="en-us">:</span></font>
                                <input type="text" name="Meta_Designer" size="35" class="font9" maxlength="80" value="<%=RS("Meta_Designer")%>">
                              </td>
                            </tr>
                            <tr>
                              <td width="12%" align="right"><font color="#000080">關鍵字<span lang="en-us">:</span></font></td>
                              <td width="88%"><input type="text" name="Meta_Keywords" size="80" class="font9" maxlength="400" value="<%=RS("Meta_Keywords")%>"></td>
                            </tr>
                            <tr>
                              <td width="12%" align="right"><font color="#000080">描述<span lang="en-us">:</span></font></td>
                              <td width="88%"><input type="text" name="Meta_Description" size="80" class="font9" maxlength="400" value="<%=RS("Meta_Description")%>"></td>
                            </tr>
                            <tr>
                              <td width="12%" align="right"><font color="#000080">摘要<span lang="en-us">:</span></font></td>
                              <td width="88%"><input type="text" name="Meta_Abstract" size="80" class="font9" maxlength="400" value="<%=RS("Meta_Abstract")%>"></td>
                            </tr>
                            <tr>
                              <td width="12%" align="right"><font color="#000080">頁尾關鍵字<span lang="en-us">:</span></font></td>
                              <td width="88%"><textarea rows="5" name="Meta_Foot" cols="78" class="font9"><%=RS("Meta_Foot")%></textarea></td>
                            </tr>
                            <tr>
                              <td width="12%" align="right"><font color="#000080">重訪天數<span lang="en-us">:</span></font></td>
                              <td width="88%">
                                <input type="text" name="Meta_Revisit_After" size="8" class="font9" maxlength="2" value="<%=RS("Meta_Revisit_After")%>">
                                &nbsp;&nbsp;
                                <font color="#000080">分配<span lang="en-us">:</span>
                                <select size="1" name="Meta_Distribution">
                                  <option value=""> </option>
		                              <option value="Global" <%If RS("Meta_Distribution")="Global" Then%>selected<%End If%> >Global</option>
		                              <option value="Local" <%If RS("Meta_Distribution")="Local" Then%>selected<%End If%>>Local</option>
		                              <option value="IU" <%If RS("Meta_Distribution")="IU" Then%>selected<%End If%> >Intern Use</option>
		                            </select>
		                            &nbsp;&nbsp;
                                <font color="#000080">機器人響導<span lang="en-us">:</span>
                                <select size="1" name="Meta_Robots">
                                  <option value=""> </option>
		                              <option value="All" <%If RS("Meta_Robots")="All" Then%>selected<%End If%> >All</option>
		                              <option value="None" <%If RS("Meta_Robots")="None" Then%>selected<%End If%>>None</option>
		                              <option value="Index" <%If RS("Meta_Robots")="Index" Then%>selected<%End If%>>Index</option>
		                            </select>
                              </td>
                            </tr>
                            <%If Session("user_id")="npois" Then%>
                            <tr>
                              <td width="12%" align="right"><font color="#000080">圖檔上傳上限<span lang="en-us">:</span></font></td>
                              <td width="88%">
                                <input type="text" name="MaxFileSize" size="5" class="font9" maxlength="2" value="<%=RS("MaxFileSize")%>"><font color="#000080"> (M)</font>
                                &nbsp;&nbsp;
                                <font color="#000080">影音 / 文件上傳上限<span lang="en-us">:</span></font>
                                <input type="text" name="MaxUploadSize" size="5" class="font9" maxlength="2" value="<%=RS("MaxUploadSize")%>"><font color="#000080"> (M)</font>
                                &nbsp;&nbsp;
                                <font color="#000080">Log保留<span lang="en-us">:</span></font>
                                <input type="text" name="Log_Keep" size="5" class="font9" maxlength="2" value="<%=RS("Log_Keep")%>"><font color="#000080"> (天)</font>
                                &nbsp;&nbsp;
                                <font color="#000080">流量保留<span lang="en-us">:</span></font>
                                <input type="text" name="Hit_Keep" size="5" class="font9" maxlength="2" value="<%=RS("Hit_Keep")%>"><font color="#000080"> (天)</font>
                              </td>
                            </tr>
                            <%Else%>
                            <input type="hidden" name="MaxFileSize" value="<%=RS("MaxFileSize")%>">
                            <input type="hidden" name="MaxUploadSize" value="<%=RS("MaxUploadSize")%>">
                            <input type="hidden" name="Log_Keep" value="<%=RS("Log_Keep")%>">
                            <input type="hidden" name="Hit_Keep" value="<%=RS("Hit_Keep")%>">
                            <%End If%>
                            <tr>
                              <td width="12%" align="right"><font color="#000080">許可登入I P<span lang="en-us">:</span></font></td>
                              <td width="88%"><input type="text" name="Client_IP" size="60" class="font9" maxlength="200" value="<%=RS("Client_IP")%>"><font color="#ff0000">&nbsp;(&nbsp;二個以上請用&nbsp;<b>;</b>&nbsp;區隔&nbsp;)</font></td>
                            </tr>                                                        
                          </table>
                        </td>
                      </tr>                      
                      <tr>
                        <td class="td02-c" width="100%" align="center"><button id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Update_OnClick()'> <img src='../images/update.gIf' width='20' height='20' align='absmiddle'> 修改</button>&nbsp;<button id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Cancel_OnClick()'> <img src='../images/icon6.gIf' width='20' height='15' align='absmiddle'> 離開</button>&nbsp;</td>
                      </tr>
                    </table>
                  </td>
                  <td height="445" class="table63-bg">&nbsp;</td>
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
function Update_OnClick(){
  <%call ChecklenJ("Meta_Title",100,"標題")%>
  <%call ChecklenJ("Meta_Author",80,"作者")%>
  <%call ChecklenJ("Meta_Copyright",80,"版權")%>
  <%call ChecklenJ("Meta_Designer",80,"設計者")%>
  <%call ChecklenJ("Meta_Keywords",400,"關鍵字")%>
  <%call ChecklenJ("Meta_Description",400,"描述")%>
  <%call ChecklenJ("Meta_Abstract",400,"摘要")%>
  if(document.form.Meta_Revisit_After.value!=''){
    <%call CheckNumberJ("Meta_Revisit_After","重訪天數")%>
  }
  <%call CheckStringJ("MaxFileSize","圖檔上傳上限")%>
  <%call CheckNumberJ("MaxFileSize","圖檔上傳上限")%>
  <%call CheckStringJ("MaxUploadSize","影音 / 文件上傳上限")%>
  <%call CheckNumberJ("MaxUploadSize","影音 / 文件上傳上限")%>
  <%call CheckStringJ("Log_Keep","Log保留")%>
  <%call CheckNumberJ("Log_Keep","Log保留")%>
  <%call CheckStringJ("Hit_Keep","流量保留")%>
  <%call CheckNumberJ("Hit_Keep","流量保留")%> 
  <%call ChecklenJ("Client_IP",200,"許可登入I P")%> 
  <%call SubmitJ("update")%>
}

function Cancel_OnClick(){
  location.href="dept.asp"
}
--></Script>