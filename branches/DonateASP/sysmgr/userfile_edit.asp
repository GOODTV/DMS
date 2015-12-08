<!--#include file="../include/dbfunction.asp"-->
<%
if request("action") = "update" then
   sql="update userfile set user_name='"&request("user_name")&"',password='"&request("password")&"',dept_id='"&request("dept_id")&"',"_
+"      pwd_validDate='"&request("expire_date")&"',BranchType='"&request("BranchType")&"',email='"&request("email")&"',dept_type='"&request("dept_type")&"',supervisor='"&request("supervisor")&"' where dept_id+user_id='"&request("uid")&"'"
   call ExecSQL (SQL)
   if request("dept_id")<>request("deptid") then
       sql="update user_menu set dept_id='"&request("dept_id")&"' where dept_id+user_id='"&request("uid")&"'"_
+"          update user_prog set dept_id='"&request("dept_id")&"' where dept_id+user_id='"&request("uid")&"'"     
       Set RS=Conn.Execute(SQL)  
   end if
   Session("dept_type")=request("dept_type")
   Session("all_dept_type")="'"&Replace(request("dept_type"),", ","','")&"'"
   Response.Redirect("userfile.asp")
end if

if request("action") = "delete" then
   sql="delete from userfile where dept_id+user_id = '"&request("uid")&"'"_
+"      delete from user_menu where dept_id+user_id= '"&request("uid")&"'"_
+"      delete from user_prog where dept_id+user_id= '"&request("uid")&"'"  
   call ExecSQL (SQL)
   Response.Redirect("userfile.asp")
end if

if request("action") = "add" then
   sql="insert into user_menu values ('"&request("userid")&"','"&request("menuid2")&"','"&request("dept_id")&"')"_
+"      insert into user_prog"_
+"        select '"&request("userid")&"',menu_id,progid,'"&request("user_right")&"','"&request("dept_id")&"' from prog_menu where menu_id='"&request("menuid2")&"'"
   Set RS=Conn.Execute(SQL)
end if

if request("menuid")<>"" then
   menu_id=request("menuid")
else   
  SQL="select user_menu.menu_id from user_menu join menu on user_menu.menu_id=menu.menu_id where dept_id+user_id='"&request("uid")&"'"_
+"        and menu.menu_id not in ('首頁','登出') order by menu.menu_seq"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  if Not RS1.EOF then 
     menu_id=RS1("menu_id")
  end if   
  RS1.close  
end if  
sql="select userfile.*,comp_shortname from userfile join dept on userfile.dept_id=dept.dept_id where userfile.dept_id+user_id='"&request("uid")&"'"
call QuerySQL (SQL,RS)
  
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>程式管理--修改資料</title>
</head>
<BODY>
<form name="form" action="" method="post">
  <input type="hidden" name="action"> 
  <input type="hidden" name="deptid" value="<%=rs("dept_id")%>">
  <input type="hidden" name="AddButton" value="userfile_add.asp">
  <input type="hidden" name="uid" value="<%=request("uid")%>">
	<div align="center">
<table border="0" width="780" cellspacing="1" cellpadding="0">
  <tr>
	    <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"  height="25">
        帳號管理 【修改資料】</td>
  </tr>
  <tr>
    <td width="100%">
                    <div align="center">
                      <center>
                    <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" bordercolor="#6487DC" cellpadding="2">
                      <tr> 
                        <td noWrap align=right width="20%" height=22 class="td1">
						<font color="#000080">部門名稱：</font></td>
                        <td width="29%" height=22 class="td2"> 
                        <%sql="select dept_id,dept_desc from dept order by dept_id"
                          FName="dept_id"                                                                                                                                                                                                                                                              
                          Listfield="dept_desc"                                                                                                                                                                                          
                          menusize="1"                                               
                          BoundColumn=rs("dept_id")                                                                                                                                                                  
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                         %></td>
                        <td width="18%" height=22 class="td1" align="right"> 
                          <font color="#FF0000">*</font>
                          所屬分會<font color="#000080">：</font></td>
                        <td width="31%" height=22 class="td2"> 
                        <%SQL="Select BranchType=CodeDesc From CASECODE Where CodeType='BranchType' Order By Seq"
                          FName="BranchType"
                          Listfield="BranchType"
                          menusize="1"
                          BoundColumn=rs("BranchType")   
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>                          
                        </td>
                      </tr>
                      <tr> 
                        <td noWrap align=right width="20%" height=22 class="td1">帳號：</td>
                        <td width="29%" height=22 class="td2"> 
                          <input class=font9t size=12 readonly name="userid" value="<%=RS("user_id")%>"></td>
                        <td width="18%" height=22 class="td1" align="right"> 
                          <font color="#FF0000">*</font>
                        <font color="#000080"><span lang="zh-tw">密碼</span>：</font></td>
                        <td width="31%" height=22 class="td2"> 
                          <input class=font9 
                        size=12 name=password maxlength="10" value="<%=rs("password")%>"></td>
                      </tr>
                      <tr> 
                        <td align=right width="20%" height=22 class="td1">
                        <font color="#FF0000">*</font>
                        <font color="#000080">帳號名稱：</font></td>
                        <td width="29%" height=22 class="td2"> 
                          <input class=font9 
                        size=12 name=user_name maxlength="10" value="<%=rs("user_name")%>"></td>
                        <td width="18%" height=22 align="right" class="td1"> 
                          帳號群組<font color="#000080">：</font></td>
                        <td width="31%" height=22 class="td2"> 
                          <input class=font9t size=12 readonly name="user_group" value="<%=RS("user_group")%>"></td>
                      </tr>
                      <%
                        SQL1="Select Dept_Id,Comp_ShortName From DEPT Order By Comp_Label,Dept_Id"
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        If RS1.Recordcount=1 Then
                          Response.Write "<input type='hidden' name='dept_type' value='"&RS1("Dept_Id")&"'>"   
                        Else
                          Response.Write "<tr>"
                          Response.Write "  <td align='right'class='td1'><font color='#ff0000'>*</font>負責部門：</td>"
                          Response.Write "  <td class='td2' colspan='3'>"
                          Row=1
                          While Not RS1.EOF
                            If Instr(Cstr(RS("dept_type")),Cstr(RS1("Dept_Id")))>0 Then
                              Response.Write "<input type='checkbox' name='dept_type' id='dept_type"&Row&"' value='"&RS1("Dept_Id")&"' checked >"&RS1("Comp_ShortName")&"&nbsp;"
                            Else
                              Response.Write "<input type='checkbox' name='dept_type' id='dept_type"&Row&"' value='"&RS1("Dept_Id")&"'>"&RS1("Comp_ShortName")&"&nbsp;"
                            End If
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
                        <td align=right width="20%" height=22 class="td1">密碼有效期：</td>
                        <td width="29%" height=22 class="td2"><%call Calendar("expire_date",RS("pwd_validDate"))%></td>
                        <td width="18%" height=22 align="right" class="td1"><font color="#000080"><span lang="zh-tw">建檔日期</span>：</font></td>
                        <td width="31%" height=22 class="td2"><%call Calendar("create_date",RS("create_date"))%></td>
                      </tr>
                      <!--#include file="../include/calendar2.asp"-->
                      <tr> 
                        <td align=right width="20%" height=22 class="td1">
                        Email：</td>
                        <td width="29%" height=22 class="td2">
                          	<input class=font9 
                        size=30 name=email maxlength="40" value="<%=rs("email")%>"></td>
                        <td width="18%" height=22 align="right" class="td1">
                          檔案管理：</td>
                        <td width="31%" height=22 class="td2">
                          <input type="checkbox" name="supervisor" value="Y" <%=checkFun(rs("supervisor"),"Y")%>>上傳新檔案&nbsp; </td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 height=22 class="font9" colspan="4">
                         <%EditButton()%>
                        </td>
                        </table>
                      </center>
                    </div>
  <tr>
    <td width="100%" align="center" height="10">
      </td>
  </tr>
  <tr>
    <td width="100%" align="center" height="10">
      <div align="center">
        <center>
        <table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#6487DC" width="100%" id="AutoNumber1">
          <tr>
            <td width="50%" valign="top">
            <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2" height="49">
              <tr>
    <td width="50%" align="center" bgcolor="#808080" height="15">
      <font color="#FFFFFF" size="2">程式群組授權</font><span style="font-size: 11pt"><font color="#800000"> </font> </span>&nbsp;</td>
              </tr>
              <tr>
    <td width="50%" height="19">
     <%SQL="select menuid2=menu_id from menu where menu_id NOT IN (select menu_id from user_menu where dept_id+user_id='"&request("uid")&"')"_
+"            and (menu_id<>'系統管理' and '"&session("user_group")&"' <> '系統維護' or '"&session("user_group")&"' = '系統維護') order by menu_seq"
      FName="menuid2"
      Listfield="menuid2"
      menusize="1"
      BoundColumn=""
      call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
     %>
      <input type="button" value="群組設定" name="save" class="cbutton"></td>
              </tr>
              <tr>
    			<td width="50%" height="19">
    			程式授權<font color="#000080">：</font><input type="checkbox" name="user_right" checked value="A">新增 
                <input type="checkbox" name="user_right" checked value="R">查詢 
                <input type="checkbox" name="user_right" checked value="U">修改 
                <input type="checkbox" name="user_right" checked value="D">刪除
    			　</td>
              	</tr>
              <tr>
    <td width="50%" valign="top" height="12">
    <% 
      SQL="select ser_no,menuid=user_menu.menu_id,程式群組=user_menu.menu_id from user_menu join menu on user_menu.menu_id=menu.menu_id where dept_id+user_id='"&request("uid")&"'"_ 
+"         and menu.menu_id not in ('首頁','登出') order by menu.menu_seq" 
      HLink="userfile_edit.asp?uid="&request("uid")&"&menuid=" 
      LinkParam="menuid"
      LinkTarget="_self"
      DataLink="userfile_menu_delete.asp?ser_no="
      DataParam="ser_no"
      DataTarget="_self"
      if session("user_group")="系統管理員" Or session("user_group")="系統管理" then
	     call DataLinkGrid (SQL,HLink,LinkParam,LinkTarget,DataLink,DataParam,DataTarget) 
	  else
	     call ShowGrid (SQL)
	  end if      
%> </td>
              </tr>
            </table>
            </td>
            <td width="50%" valign="top">
            <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3">
              <tr>
    <td width="50%" align="center" bgcolor="#808080">
      <font color="#FFFFFF" size="2">程式授權</font></td>
              </tr>
              <tr>
    <td width="50%">
      <input type="text" name="menu_id" size="12" readonly class="font9" value="<%=menu_id%>"> <input type="button" value="程式設定" name="setup" class="cbutton"></td>
              </tr>
              <tr>
                <td width="100%">
    <% 
     SQL="select ser_no,程式代號=user_prog.prog_id,程式名稱=prog_desc,程式授權=user_right from user_prog inner join prog on user_prog.prog_id=prog.prog_id"_
+"        where dept_id+user_id='"&request("uid")&"' and menu_id='"&menu_id&"' order by prog.prog_seq" 
     HLink="userfile_prog_delete.asp?ser_no=" 
     LinkParam="ser_no"
     LinkTarget="_self"
     if session("user_group")="系統管理員" Or session("user_group")="系統管理" then
        call DataGrid(SQL,HLink,LinkParam,LinkTarget) 
     else
	     call ShowGrid (SQL)
	 end if  
    %> 
                　</td>
              </tr>
            </table>
            </td>
          </tr>
        </table>
        </center>
      </div>
      </td>
  </tr>
      </td>
  </tr>
</table>
	</div>
<%message%>           
</form>
</BODY>
</HTML>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--
sub window_OnLoad
  <%if session("user_group")<>"系統管理員" And session("user_group")<>"系統管理" then%>
       form.save.disabled=true
       form.setup.disabled=true
  <%end if%>     
end sub

sub Update_OnClick
  <%call CheckString("BranchType","所屬分會")%>
  <%call CheckString("password","密碼")%>
  <%call CheckString("user_name","帳號名稱")%>
  <%call Checklen("user_name",10,"帳號名稱")%>
  ConfirmUpdate()
end sub

sub setup_OnClick
  location.href="userfile_setup.asp?dept_id="&form.dept_id.value&"&user_id="&form.userid.value&"&uid=<%=request("uid")%>&menu_id=<%=menu_id%>"
end sub                                             

sub save_OnClick
  form.action.value="add"
  form.submit
end sub

sub cancel_OnClick
  location.href="userfile.asp"
end sub

sub query_OnClick
  location.href="userfile.asp"
end sub  

--></script>