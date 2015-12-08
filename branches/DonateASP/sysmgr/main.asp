<%
If request("action")="change" Then
  SQL1="Select * From DEPT Where Dept_Id='"&Request("Dept_Id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  Session("sys_name")=RS1("sys_name")
  Session("dept_id")=RS1("dept_id")
  Session("dept_desc")=RS1("dept_desc")
  Session("comp_label")=RS1("comp_label")
  Session("comp_name")=RS1("comp_name")
  Session("comp_ShortName")=RS1("comp_ShortName")
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="切換成功，您已經重新以『"&Session("comp_ShortName")&"』管理者身份登入\n\n請您繼續使用本系統！"
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title><%=session("sys_name")%></title>
</head>
<frameset framespacing="0" border="0" frameborder="0" rows="80,*">
  <frame name="top" scrolling="no" src="main_top.asp">
  <frameset cols="180,*" framespacing="0" frameborder="NO" border="0">
    <frame name="left" src="xmenu.asp" frameborder="no" scrolling="auto" noresize id="left">
    <frame name="main" src="main_d.asp" frameborder="no" scrolling="auto">
  </frameset>
  <noframes>
    <body>
      <p>此網頁使用框架，但是您的瀏覽器並不支援。</p>
    </body>
  </noframes>
</frameset>
</html>