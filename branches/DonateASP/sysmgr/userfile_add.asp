<!--#include file="../include/dbfunction.asp"-->
<%
if request("action")="save" then
   sql="select * from userfile where user_id='"&request("userid")&"' and dept_id='"&request("dept_id")&"'"
   Set RS1 = Server.CreateObject("ADODB.RecordSet")
   RS1.Open SQL,Conn,1,1
   if Not RS1.EOF then
      session("errnumber")=1
      session("msg")="資料重覆 ! "&request("userid")&" 使用者已存在"
   else
      sys_date=date()
      sql="select password_day from dept where dept_id='"&session("dept_id")&"'"
   	  call QuerySQL(SQL,RS)
      if not RS.EOF then
         if RS("password_day") > 0 then
            expire_date=DateAdd("d",RS("password_day"),date())   
         else
            expire_date="2099/12/31"  
         end if    
      end if
      sql="insert into userfile values ('"&request("userid")&"','"&request("password")&"','"&request("user_name")&"','"&request("usergroup")&"',"_
+"        '"&request("dept_id")&"','"&request("supervisor")&"','"&expire_date&"','','"&sys_date&"','','"&request("BranchType")&"','"&request("email")&"','"&request("dept_type")&"')"_
+"         insert into user_menu"_
+"           select '"&request("userid")&"',menu_id,'"&request("dept_id")&"' from group_menu where user_group='"&request("usergroup")&"'"
      Set RS=Conn.Execute(SQL)
      sql="select * from group_menu where user_group='"&request("usergroup")&"'"
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL,Conn,1,1
      While Not RS1.EOF
        sql1="insert into user_prog"_
+"              select '"&request("userid")&"',menu_id,prog_id,user_right,'"&request("dept_id")&"' from group_prog where user_group='"&request("usergroup")&"' and menu_id='"&RS1("menu_id")&"'"
        Set RS=Conn.Execute(SQL1) 
        RS1.MoveNext
      Wend
      RS1.Close
      Response.Redirect "userfile_edit.asp?uid="&request("dept_id")+request("userid")
   end if
end if   
      
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>帳號管理--修改資料</title>
</head>
<BODY>
<div align="center">
<table border="0" width="780" cellspacing="1" cellpadding="0">
  <tr>
	    <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"  height="25">
        帳號管理【新增資料】</td>
  </tr>
  <tr>
    <td width="100%">
                  <form name="form" action="" method="post">
                  <input type="hidden" name="action"> 
                    <div align="center">
                      <center>
                    <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" bordercolor="#6487DC" cellpadding="2">
                      <tr> 
                        <td noWrap align=right width="20%" bgcolor=#FFFFCC height=22>
                        <font color="#FF0000">* </font> 
                        <font color="#000099">部門名稱</font><font color="#000080">：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22>
                         <%sql="select dept_id,dept_desc from dept order by dept_id"
                          FName="dept_id"                                                                                                                                                                                                                                                              
                          Listfield="dept_desc"                                                                                                                                                                                          
                          menusize="1"                                               
                          BoundColumn=""                                                                                                                                                                  
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                         %> 
                          	　</td>
                      </tr>
                      <tr> 
                        <td noWrap align=right width="20%" bgcolor=#FFFFCC height=22>
                        <font color="#FF0000">*</font> <font color="#000080">
                        <span lang="zh-tw">帳號</span>：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          	<input class=font9 size=12 
                        name="userid" maxlength="10"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" height=22 bgcolor="#FFFFCC">
                        <font color="#FF0000">*</font>
                        <font color="#000080"><span lang="zh-tw">密碼</span>：</font></td>
                        <td width="80%" height=22 bgcolor="#FFFCE1"> 
                          	<input class=font9 
                        size=12 name=password maxlength="10"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" bgcolor=#FFFFCC height=22 class="td1">
                        <font color="#FF0000">*</font>
                        <font color="#000080">帳號名稱：</font></td>
                        <td valign=center width="80%" bgcolor=#FFFCE1 height=22> 
                          	<input class=font9 
                        size=12 name=user_name maxlength="10"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" height=22 class="td1">
                        <font color="#FF0000">*</font> 帳號群組：</td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1"> 
                        <%SQL="select usergroup=user_group from groupfile where"_
+"                             user_group <>'系統管理員' and '"&session("user_group")&"' <> '系統管理員' or '"&session("user_group")&"' = '系統管理員'"  
                          FName="usergroup"
                          Listfield="usergroup"  
                          BoundColumn=""  
                          menusize="1"  
                         call OptionList (SQL,FName,Listfield,BoundColumn,menusize)%>
                          　</td>
                      </tr>
                      <%
                        SQL1="Select Dept_Id,Comp_ShortName From DEPT Order By Comp_Label,Dept_Id"
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        If RS1.Recordcount=1 Then
                          Response.Write "<input type='hidden' name='dept_type' value='"&RS1("Dept_Id")&"'>"   
                        Else
                          Response.Write "<tr>"
                          Response.Write "  <td align='right'class='td1'><font color='#FF0000'>*</font>負責部門：</td>"
                          Response.Write "  <td valign='center' bgcolor='#FFFCE1'>"
                          Row=1
                          While Not RS1.EOF
                            Response.Write "<input type='checkbox' name='dept_type' id='dept_type"&Row&"' value='"&RS1("Dept_Id")&"' checked >"&RS1("Comp_ShortName")&"&nbsp;"
                            Row=Row+1
                            RS1.MoveNext
                          Wend	
                          Response.Write "  </td>"
                          Response.Write "</tr>"
                        End If
                        RS1.Close
                        Set RS1=Nothing
                      %>               
                      <tr> 
                        <td align=right width="20%" height=22 class="td1">
                        Email：</td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1"> 
                          	<input class=font9 
                        size=30 name=email maxlength="40"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" height=22 class="td1">
                          檔案管理：</td>
                        <td valign=center width="80%" height=22 bgcolor="#FFFCE1"> 
                          <input type="checkbox" name="supervisor" value="Y">上傳新檔案&nbsp; </td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 height=22 class="font9" colspan="2">
                         <%SaveButton%>　
                        </td>
                        </table>
                      </center>
                    </div>
                  </form>
      </td>
  </tr>
</table>

</div>

<%message%>           

</BODY>
</HTML>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--

sub save_OnClick
  <%call CheckString("dept_id","部門名稱")%>
  <%call CheckString("userid","帳號")%>
  <%call CheckString("password","密碼")%>
  <%call CheckString("user_name","帳號名稱")%>
  <%call Checklen("user_name",10,"帳號名稱")%>
  <%call CheckString("usergroup","帳號群組")%>
  form.action.value="save"
  form.submit   
end sub

sub cancel_OnClick
  location.href="userfile.asp"
end sub

--></script>