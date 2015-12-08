<%
LogoUrl="../images/logo.gif"
LogoUrl_C="../images/case.gif"
LogoUrl_D="../images/logo.png"
LogoUrl_M="../images/member.gif"
Set objFS = Server.CreateObject("Scripting.FileSystemObject")
If objFS.FileExists(Server.MapPath(LogoUrl_C)) Then LogoUrl=LogoUrl_C
If objFS.FileExists(Server.MapPath(LogoUrl_D)) Then LogoUrl=LogoUrl_D
If objFS.FileExists(Server.MapPath(LogoUrl_M)) Then LogoUrl=LogoUrl_M
Set objFS = Nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta http-equiv="refresh" CONTENT=300>	
  <title><%=session("sys_name")%></title>
  <style type="text/css">
  <!--
    body {
      margin:0px;
      background-repeat: repeat-x;
      background-color: #ffffff;
      background-image: url(../images/background_image.jpg);
    }
    .system{
	    font-size: 16px;
	    color: #FFFFFF;
	    font-weight: bold;
	    font-family: Arial, Helvetica, sans-serif;
	    text-decoration: none;
	    line-height: 22px;
    }
    .menu {
	    font-size: 10pt;
	    font-weight: bold;
	    color: #FFFFFF;
	    font-family: "Verdana", "Arial", "Helvetica", "sans-serif";
	    text-decoration: none;
    }
    a.menu:link {
	    font-family: "Verdana", "Arial", "Helvetica", "sans-serif";
	    font-size: 10pt;
	    line-height: 19pt;
	    color: #FFFFFF;
	    text-decoration: none;
    }
    a.menu:visited {
	    font-family: "Verdana", "Arial", "Helvetica", "sans-serif";
	    font-size: 10pt;
	    line-height: 19pt;
	    color: #FFFFFF;
	    text-decoration: none;
    }
    a.menu:hover {
	    font-family: "Verdana", "Arial", "Helvetica", "sans-serif";
	    font-size: 10pt;
	    line-height: 19pt;
	    color: #A70500;
	    text-decoration: none;
    }
    a.menu:active {
	    font-family: "Verdana", "Arial", "Helvetica", "sans-serif";
	    font-size: 10pt;
	    line-height: 19pt;
	    color: #FFFFFF;
	    text-decoration: none;
    }	 
  -->
  </style>
	<base target="_self">
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onselectstart="return false;" ondragstart="return false;" oncontextmenu="return false;">
  <table width="1000" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
    <tr>
      <td width="626"><img src="<%=LogoUrl%>" height="59"></td>
      <td valign="bottom">
        <table width="271" border="0" align="right" cellpadding="0" cellspacing="0">
          <tr>
            <td width="85">
              <table width="85" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="7"><img src="../images/maintop01.gif" width="7" height="26" /></td>
                  <td width="73" background="../images/maintop02.gif"><div align="center"><a href="../index.asp" class="menu" target="_top">網站首頁</a></div></td>
                  <td width="5"><div align="right"><img src="../images/maintop03.gif" width="5" height="26" /></div></td>
                </tr>
              </table>
            </td>
            <td width="4">&nbsp;</td>
            <td width="85">
              <table width="85" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="7"><img src="../images/maintop01.gif" width="7" height="26" /></td>
                  <td width="73" background="../images/maintop02.gif"><div align="center"><a href="../sysmgr/main.asp" class="menu" target="_top">管理訊息</a></div></td>
                  <td width="5"><div align="right"><img src="../images/maintop03.gif" width="5" height="26" /></div></td>
                </tr>
              </table>
            </td>
            <td width="2">&nbsp;</td>
            <td>
              <table width="85" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="7"><img src="../images/maintop01.gif" width="7" height="26" /></td>
                  <td width="73" background="../images/maintop02.gif"><div align="center">
                  	<a href="JavaScript:if(confirm('是否確定要登出 ?')){window.location.href='../admin/index.asp';}" class="menu" target="_top">管理登出</a></div></td>
                  <td width="5"><div align="right"><img src="../images/maintop03.gif" width="5" height="26" /></div></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <table width="1000" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
    <tr>
      <td height="25" background="../images/maintop04.jpg">
        <table width="1000" border="0" cellspacing="0" cellpadding="0" background="../images/maintop04.jpg">
          <tr>
            <td width="169">&nbsp;</td>
            <td width="831" class="system"><%=session("sys_name")%></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>