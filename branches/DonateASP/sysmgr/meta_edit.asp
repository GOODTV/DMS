<!--#include file="../include/dbfunction.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From DEPT Where DEPT_ID='"&request("DEPT_ID")&"'"
  Set RS2 = Server.CreateObject("ADODB.RecordSet")
  RS2.Open SQL,Conn,1,3
  RS2("META_TITLE")=request("META_TITLE")
  RS2("META_AUTHOR")=request("META_AUTHOR")
  RS2("META_COPYRIGHT")=request("META_COPYRIGHT")
  RS2("META_KEYWORDS")=request("META_KEYWORDS")
  RS2("META_DESCRIPTION")=request("META_DESCRIPTION")
  RS2("META_ABSTRACT")=request("META_ABSTRACT")
  RS2("META_REVISIT_AFTER")=request("META_REVISIT_AFTER")
  If request("META_LANGUAGE")<>"" Then
    RS2("META_LANGUAGE")=request("META_LANGUAGE")
  Else
    RS2("META_LANGUAGE")="utf-8"
  End If
  If request("META_DISTRIBUTION")<>"" Then
    RS2("META_DISTRIBUTION")=request("META_DISTRIBUTION")
  Else
    RS2("META_DISTRIBUTION")="Global"
  End If
  If request("META_ROBOTS")<>"" Then
    RS2("META_ROBOTS")=request("META_ROBOTS")
  Else
    RS2("META_ROBOTS")="All"
  End If  
  RS2.Update
  RS2.Close
  Set RS2=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 !"
  response.redirect "dept.asp"  
End If

If request("action")="delete" Then
  SQL=" Delete From DEPT Where DEPT_ID='"&request("DEPT_ID")&"' "
  Set RS1=Conn.Execute(SQL)
  session("errnumber")=1
  session("msg")="資料刪除成功 !"
  Response.Redirect "dept.asp"
End If

SQL="Select * From DEPT Where DEPT_ID='"&request("DEPT_ID")&"'"
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
    <input type="hidden" name="action">
    <input type="hidden" name="DEPT_ID" value="<%=request("DEPT_ID")%>"> 
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
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">機構參數設定</td>
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
                          <td class="button2-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_edit.asp?DEPT_ID=<%=RS("DEPT_ID")%>" class="tool">機構參數設定</a></td>
                          <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="meta_edit.asp?DEPT_ID=<%=RS("DEPT_ID")%>" class="tool"><font color="#800000"><b>Meta參數設定護</b></font></font></a></td>
                          <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="ecpay_edit.asp?DEPT_ID=<%=RS("DEPT_ID")%>" class="tool">捐款參數設定</a></td>
                          <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="donate_edit.asp?DEPT_ID=<%=RS("DEPT_ID")%>" class="tool">捐款機制說明</a></td>
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
                          <td align="right"><font color="#000080">標題<span lang="en-us">:</span></font></td>
                          <td><input type="text" name="META_TITLE" size="80" class="font9" maxlength="80" value=<%=RS("META_TITLE")%>></td>
                        </tr>
                        <tr>
                          <td align="right"><font color="#000080">作者<span lang="en-us">:</span></font></td>
                          <td><input type="text" name="META_AUTHOR" size="80" class="font9" maxlength="80" value="<%=RS("META_AUTHOR")%>"></td>
                        </tr>
                        <tr>
                          <td align="right"><font color="#000080">版權<span lang="en-us">:</span></font></td>
                          <td><input type="text" name="META_COPYRIGHT" size="80" class="font9" maxlength="80" value=<%=RS("META_COPYRIGHT")%>></td>
                        </tr>                        
                        <tr>
                          <td align="right"><font color="#000080">關鍵字<span lang="en-us">:</span></font></td>
                          <td><input type="text" name="META_KEYWORDS" size="80" class="font9" maxlength="400" value=<%=RS("META_KEYWORDS")%>></td>
                        </tr>
                        <tr>
                          <td align="right"><font color="#000080">描述<span lang="en-us">:</span></font></td>
                          <td><input type="text" name="META_DESCRIPTION" size="80" class="font9" maxlength="400" value="<%=RS("META_DESCRIPTION")%>"></td>
                        </tr>
                        <tr>
                          <td align="right"><font color="#000080">摘要<span lang="en-us">:</span></font></td>
                          <td><input type="text" name="META_ABSTRACT" size="80" class="font9" maxlength="400" value=<%=RS("META_ABSTRACT")%>></td>
                        </tr>
                        <tr>
                          <td align="right"><font color="#000080">重訪天數<span lang="en-us">:</span></font></td>
                          <td>
                            <input type="text" name="META_REVISIT_AFTER" size="5" class="font9" maxlength="2" value="<%=RS("META_REVISIT_AFTER")%>">&nbsp;&nbsp;&nbsp;&nbsp;
                            <font color="#000080">語言<span lang="en-us">:</span></font>
                            <input type="radio" name="META_LANGUAGE" value="utf-8" <%If RS("META_LANGUAGE")<>"zh-CN" Then%>checked<%End If%> >utf-8&nbsp;
                            <input type="radio" name="META_LANGUAGE" value="zh-CN" <%If RS("META_LANGUAGE")="zh-CN" Then%>checked<%End If%>>zh-CN&nbsp;&nbsp;&nbsp;&nbsp;
                            <font color="#000080">分配<span lang="en-us">:</span></font>
                            <input type="radio" name="META_DISTRIBUTION" value="Global" <%If RS("META_DISTRIBUTION")="Global" Then%>checked<%End If%> >Global&nbsp;
                            <input type="radio" name="META_DISTRIBUTION" value="Local" <%If RS("META_DISTRIBUTION")="Local" Then%>checked<%End If%>>Local&nbsp;
                            <input type="radio" name="META_DISTRIBUTION" value="IU" <%If RS("META_DISTRIBUTION")="IU" Then%>checked<%End If%>>Intern Use&nbsp;&nbsp;&nbsp;&nbsp;
                          </td>
                        </tr>
                        <tr>
                          <td align="right"><font color="#000080">機器人導向<span lang="en-us">:</span></font></td>
                          <td>
                            <input type="radio" name="META_ROBOTS" value="All" <%If RS("META_ROBOTS")="All" Then%>checked<%End If%> >All&nbsp;
                            <input type="radio" name="META_ROBOTS" value="None" <%If RS("META_ROBOTS")="None" Then%>checked<%End If%>>None&nbsp;
                            <input type="radio" name="META_ROBOTS" value="Index" <%If RS("META_ROBOTS")="Index" Then%>checked<%End If%>>Index&nbsp;
                            <input type="radio" name="META_ROBOTS" value="No Index" <%If RS("META_ROBOTS")="No Index" Then%>checked<%End If%>>No Index&nbsp;
                            <input type="radio" name="META_ROBOTS" value="Follow" <%If RS("META_ROBOTS")="Follow" Then%>checked<%End If%>>Follow&nbsp;
                            <input type="radio" name="META_ROBOTS" value="No Follow" <%If RS("META_ROBOTS")="No Follow" Then%>checked<%End If%>>No Follow
                          </td>
                        </tr>                                                                                                                         
                        <tr> 
                          <td height="1" colspan="2"><hr></td>
                        </tr>                                    
                      </table>
                    </td>
                  </tr>
                  <tr>
                    <td class="td02-c" width="100%" align="center"><button id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg'> <img src='../images/update.gif' width='19' height='20' align='absmiddle'> 修改</button><button id='query' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg'> <img src='../images/search.gif' width='19' height='20' align='absmiddle'>查詢</button><button id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg'> <img src='../images/document.gif' width='19' height='20' align='absmiddle'>上一頁</button></td>
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
  <%message%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<Script Language="VBScript"><!--
sub window_OnLoad
      
end sub

sub Update_OnClick
  <%call Checklen("META_TITLE",80,"標題")%>
  <%call Checklen("META_AUTHOR",80,"作者")%>
  <%call Checklen("META_COPYRIGHT",80,"版權")%>
  <%call Checklen("META_KEYWORDS",400,"關鍵字")%>
  <%call Checklen("META_DESCRIPTION",400,"描述")%>
  <%call Checklen("META_ABSTRACT",400,"摘要")%>
  If form.META_REVISIT_AFTER.value="" Then
    form.META_REVISIT_AFTER.value="30"
  End If
  <%call CheckNumber("META_REVISIT_AFTER","重訪天數")%>
  ConfirmUpdate()
end sub

sub query_OnClick
  location.href="dept.asp"
end sub
--></Script>