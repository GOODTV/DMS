<!--#include file="../include/dbfunction.asp"-->
<%
if request("dept_id")<>"" then
  dept_id=request("dept_id")
else
  dept_id=session("dept_id")
end if     
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>系統LOG查詢</title>
</head>
<body class=gray>
  <form method="POST" action="syslog_qry_d.asp" name="form" target="detail">
    <div align="center"><center>
      <table border="1" width="95%" cellspacing="1" style="border-collapse: collapse" cellpadding="2">
        <tr>
          <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25"><font color="#FFFFFF">系統LOG查詢</font></td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" width="100%">
              <tr>
                <td width="16%" class="font9A" align="right">Log 日期<font color="#000080">：</font></td>                                                                 
                <td width="29%">
                  <input type=text name=date1 size=10 class=font9>
                  <a href onclick=cal19.select(document.form.date1,'date1','yyyy/MM/dd');>
                  <img border=0 src=../images/date.gif width=16 height=14></a>
                   ~ 
                  <input type=text name=date2 size=10 class=font9>
                  <a href onclick=cal19.select(document.form.date2,'date2','yyyy/MM/dd');>
                  <img border=0 src=../images/date.gif width=16 height=14></a>
                </td>
                <td width="11%" align="right">機構<font color="#000080">：</font></td>
                <td width="43%" colspan="2">
                <%sql="select dept_id,comp_shortname from dept where '"&session("user_group")&"'='系統管理員' or dept_id='"&session("dept_id")&"'"
                  FName="dept_id"                                                                                                                                                                                                                                                              
                  Listfield="comp_shortname"                                                                                                                                                                                          
                  menusize="1"                                               
                  BoundColumn=request("dept_id")                                                                                                                                                                  
                  call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                %>         
          　    </td>
              </tr>
              <!--#include file="../include/calendar2.asp"-->
              <tr>
                <td width="16%" class="font9A" align="right">Log 類別<font color="#000080">：</font></td>                                                                 
                <td width="29%">
                  <select size="1" name="log_type" class="font9">
                    <option value=""> </option>
                    <option value="系統登入">系統登入</option>
                    <option value="更改密碼">更改密碼</option>
                    <option value="使用者異動">使用者異動</option>
                    <option value="刪除資料">刪除資料</option>
                  </select>
                </td>
                <td width="11%" align="right">帳號<font color="#000080">：</font></td>
                <td width="16%">
                <%SQL="select userid=user_id,username=user_id+' '+user_name from userfile where dept_id='"&dept_id&"' order by user_id"  
                  FName="userid"                                                                   
                  Listfield="username"                                                                     
                  menusize="1"        
                  BoundColumn=""       
                  call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                %>
          　    </td>
                <td width="27%"><input type="button" value=" 查 詢 " name="query" class="cbutton"></td>
              </tr>
              <tr>       
                <td width="100%" colspan="5"><iframe name="detail" height="440" width="100%" frameborder="0" scrolling="auto"></iframe></td>       
              </tr>         
            </table>        
          </td>        
        </tr>       
      </table>      
    </center></div>       
  </form> 
</body>
</html>        
<!--#include file="../include/dbclose.asp"-->        
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--                              
sub Query_OnClick 
  form.submit                              
end sub

sub dept_id_OnChange
  location.href="syslog_qry.asp?dept_id="&form.dept_id.value
end sub                    
--></script>