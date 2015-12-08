<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL=" Select * From user_menu join menu On user_menu.menu_id=menu.menu_id " & _
    " Where dept_id='"&session("menu_dept_id")&"' And user_id='"&session("menu_user_id")&"' " & _
    " Order By menu.menu_seq"
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.Open SQL,Conn,1,1
RecordCount=RS.RecordCount
Total_Div=RS.RecordCount
If UBound(Split(Session("dept_type"),","))+1>1 Then Total_Div=Total_Div+1
I=1
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>XP_滑動選單</title>
  <style type=text/css>
    body  { background:#efefe9; margin:0px; font:normal 12px 細明體; Scrollbar-Face-Color:#EAEAD7; Scrollbar-Arrow-Color:#4B4B4B; Scrollbar-Base-Color:#484824; Scrollbar-Shadow-Color:#FFFFFF; Scrollbar-Dark-Shadow-Color:#FFFFFF; Scrollbar-Highlight-Color:#FFFFFF; Scrollbar-3D-Light-Color:#445FCA;}
    table  { border:0px; }
    td  { font:normal 12px 細明體; }
    img  { vertical-align:bottom; border:0px; }
    a  { font:normal 12px 細明體; color:#215DC6; text-decoration:none; }
    a:hover  { color:#428EFF }
    .sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
    .menu_title  { }
    .menu_title span  { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; }
    .menu_title2  { }
    .menu_title2 span  { position:relative; top:2px; left:8px; color:#428EFF; font-weight:bold; }
  </style>
</head>
<body onselectstart="return false;" ondragstart="return false;" oncontextmenu="return false;">
  <script>
    function startmenu(){
      <%for J = 1 to Total_Div%>
        document.getElementById('menu<%=J%>').style.display = "none";
      <%Next%>
    }
    function menufunc(menuId){
      if(document.getElementById(menuId).style.display == "none"){
        startmenu();
	      document.getElementById(menuId).style.display = "block";
      }else{
	      startmenu();
      }
    }
  </script>
  <table border="0" cellspacing="0" cellpadding="0" width="160">
    <tr>
      <td width="33" align="right" background="../images/xmenu01.gif" height="40"><img src="../images/xmenu02.gif" alt="" width="22" height="22" /></td>
      <%If Instr(Session("dept_type"),",")>0 Then%>
      <td width="129" class="explain" background="../images/xmenu01.gif" height="40">&nbsp;<%=session("user_name")%>-<%=Session("comp_ShortName")%></td>
      <%Else%>
      <td width="129" class="explain" background="../images/xmenu01.gif" height="40">&nbsp;<%=session("user_name")%></td>
      <%End If%>
    </tr>
    <tr>
      <td height="2" background="../images/xmenu03.gif" colspan="2"></td>
    </tr>
    <tr>
      <td colspan="2"><img src="../images/xmenu04.jpg" alt="" width="162" height="17" /></td>
    </tr>
  </table>
  <%
    While Not RS.Eof
      SQL=" Select * From user_prog join prog On user_prog.prog_id=prog.prog_id " & _
          " Where dept_id='"&Session("dept_id_login")&"' And user_id='"&session("user_id")&"' And menu_id='"&rs("menu_id")&"' " & _
          " Order By prog.prog_seq"
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL,Conn,1,1
      J=RS1.RecordCount*20+20
      If I = 1 Then
        displaydata=""
      Else
        displaydata="none"
      End If
  %> 
  <table border="0" cellspacing="0" cellpadding="0" width="160" >
    <tr style="cursor:hand;">
      <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=../images/item_show.gif><a href="javascript:menufunc('menu<%=I%>')"><span class='menu_title'><%=rs("menu_id")%></span></a></td>
    </tr>
    <tr>
      <td>
	      <div class=sec_menu style="width:160;height:0;filter:alpha(Opacity=100);overflow:hidden;display:none;" id=menu<%=i%>>
	        <table cellpadding=0 cellspacing=0 align=center width=145 style="position:relative;top:10px;">
	          <%
	            While Not RS1.Eof
	          %>
	          <tr>
	            <td height=20><a href="../<%=RS1("prog_url")%>" target="main" onfocus=this.blur();><img src=../images/book.gif width="16" height="16">&nbsp;<%=RS1("prog_desc")%></a></td>
	          </tr>
            <%
                RS1.MoveNext
              Wend
            %>
            <tr><td height="15px"></td></tr>
          </table>
        </div>
      </td>
    </tr>
  </table>
  <%
      I=I+1
      RS.MoveNext
    Wend
  %>
  <%If UBound(Split(Session("dept_type"),","))+1>1 Then%>
  <table border="0" cellspacing="0" cellpadding="0" width="160" >
    <tr style="cursor:hand;">
      <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=../images/item_show.gif><a href="javascript:menufunc('menu<%=I%>')"><span class='menu_title'>部門作業切換</span></a></td>
    </tr>
    <tr>
      <td>
	      <div class=sec_menu style="width:160;height:0;filter:alpha(Opacity=100);overflow:hidden;display:none;" id=menu<%=i%>>
	        <table cellpadding=0 cellspacing=0 align=center width=145 style="position:relative;top:10px;">
	          <tr>
	            <td height=20><a href="../donation/dept_change.asp" target="main" onfocus=this.blur();><img src=../images/book.gif width="16" height="16">&nbsp;切換作業部門</a></td>
	          </tr>
            <tr><td height="15px"></td></tr>
          </table>
        </div>
      </td>
    </tr>
  </table>
  <%End If%>  
  <br><center><%=Request.ServerVariables("REMOTE_ADDR")%></center>
</body>
</html>
