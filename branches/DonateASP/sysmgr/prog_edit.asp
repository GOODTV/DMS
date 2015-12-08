<!--#include file="../include/dbfunction.asp"-->
<%
if request("action") = "update" then
   sql="update prog set prog_desc='"&request("prog_desc")&"',prog_url='"&request("prog_url")&"',"_
+"      prog_seq='"&request("prog_seq")&"',prog_type='"&request("prog_type")&"',security='"&request("security")&"'"_
+"      where prog_id = '"&request("prog_id")&"' "
   call UpdateSQL (SQL)
   Response.Redirect("prog.asp")
end if

if request("action") = "delete" then
   sql="delete from prog where prog_id = '"&request("prog_id")&"'"_
+"      delete from prog_menu where progid='"&request("prog_id")&"'"_
+"      delete from user_prog where prog_id='"&request("prog_id")&"'"   
   call UpdateSQL (SQL)
   Response.Redirect("prog.asp")
end if

  sql="select * from prog where prog_id='"&request("prog_id")&"'"
  call QuerySQL (SQL,RS)
  session("prog_id")=request("prog_id")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>程式管理--修改資料</title>
</head>
<body>
  <form name="form" action="" method="post">
    <input type="hidden" name="action"> 
    <input type="hidden" name="AddButton" value="prog_add.asp">
    <input type="hidden" name="ProgType" value="<%=rs("prog_type")%>">
    <table border="0" width="80%" cellspacing="1" cellpadding="0" align="center">
    <tr>
	    <td width="100%" class="font11" colspan="2" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"  height="25">程式管理【修改資料】</td>
    </tr>
    <tr>
      <td width="100%" colspan="2">
        <div align="center"><center>
          <table width="100%" border=1 cellspacing="1" style="border-collapse: collapse" bordercolor="#6487DC" cellpadding="2">
            <tr> 
              <td noWrap align=right width="18%" height=22 class="td1"><font color="#FF0000">* </font>程式代號：</td>
              <td width="80%" height=22 class="td2"> <input class=font9 size=20 readonly name="prog_id" value="<%=RS("prog_id")%>"></td>
            </tr>
            <tr> 
              <td align=right width="18%" height=22 bgcolor="#FFFFCC"><font color="#FF0000">*</font><font color="#000080"><span lang="zh-tw">程式名稱</span>：</font></td>
              <td width="80%" height=22 bgcolor="#FFFCE1"> <input class=font9 size=40 name=prog_desc maxlength="40" value="<%=rs("prog_desc")%>"></td>
            </tr>
            <tr> 
              <td align=right width="18%" bgcolor=#FFFFCC height=22 class="td1"><font color="#000080"><span lang="zh-tw">連結URL</span>：</font></td>
              <td valign=center width="80%" bgcolor=#FFFCE1 height=22> <input class=font9 size=40 name=prog_url maxlength="100" value="<%=rs("prog_url")%>"></td>
            </tr>
            <tr> 
              <td align=right width="18%" bgcolor=#FFFFCC height=22 class="td1"><font color="#000080"><span lang="zh-tw">排序</span>：</font></td>
              <td valign=center width="80%" bgcolor=#FFFCE1 height=22><input type="text" name="prog_seq" size="4" class="font9" maxlength="3" value="<%=rs("prog_seq")%>"></td>
            </tr>
            <tr> 
              <td align=right width="18%" height=22 bgcolor="#FFFFCC" class="td1">程式類別<font color="#000080">：</font></td>
              <td valign=center width="80%" height=22 bgcolor="#FFFCE1"> 
                <select size="1" name="prog_type" class="font9A">
                  <option selected value="ASP">ASP</option>
                  <option value="QRY">QRY</option>
                  <option value="RPT">RPT</option>
                </select>
              </td>
            </tr>
            <tr> 
              <td align=right width="18%" bgcolor=#FFFFCC height=22 class="td1"><span lang="zh-tw">安全</span><font color="#000080"><span lang="zh-tw">控管</span>：</font></td>
              <td valign=center width="80%" height=22 bgcolor="#FFFCE1"><input type="text" name="security" size="4" class="font9" maxlength="1" value="<%=rs("security")%>"></td>
            </tr>
            <tr>
              <td width="100%" bgcolor=#FFFCE1 height=22 class="font9" colspan="2"><%EditButton()%></td>
            </tr>
          </table>
        </center></div>
      <tr>
        <td width="50%" align="right"><font color="#800000"><span style="font-size: 11pt">帳號授權</span></font> </td>
        <td width="50%" align="center"><input type="button" value="帳號設定" name="setup" class="cbutton"></td>
      </tr>
      <tr>
        <td width="100%" align="center" colspan="2">
        <% 
          SQL="select ser_no,機構=user_prog.dept_id,程式群組=menu_id,帳號=user_prog.User_ID,姓名=user_name,帳號類別=user_group from user_prog"_
          +"     join userfile on user_prog.dept_id+user_prog.user_id=userfile.dept_id+userfile.user_id"_
          +"     where prog_id='"&request("prog_id")&"' order by user_prog.dept_id,menu_id,user_prog.User_ID" 
          HLink="prog_setup_delete.asp?ser_no=" 
          LinkParam="ser_no"
          LinkTarget="_self"
          call DataGrid(SQL,HLink,LinkParam,LinkTarget) 
        %> 
        </td>
      </tr>
    </td>
  </tr>
</table>
<%message%>           
</form>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--

sub window_OnLoad
  form.prog_type.value=form.progtype.value
end sub

sub Update_OnClick
  <%call CheckString("prog_desc","程式名稱")%>
  <%call Checklen("prog_desc",40,"程式名稱")%>
  ConfirmUpdate()
end sub

sub setup_OnClick
  prog_desc=form.prog_desc.value
  location.href="prog_setup.asp?prog_id=<%=request("prog_id")%>"&"&prog_desc="&prog_desc
end sub                                             

sub cancel_OnClick
  history.back
end sub

sub query_OnClick
  location.href="prog.asp"
end sub  

--></script>