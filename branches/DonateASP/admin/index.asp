<%
Sub Message()
  If session("errnumber")=0 then
    Response.Write "<center>"&session("msg")&"</center>"
  Else
    Response.Write "<script Language='JavaScript'>alert('"&session("msg")&"')</script>"
  End If
  session("msg")=""
  session("errnumber")=0
End Sub
If Cint(Session("LoginErrTime"))>=6 Or Session("User_Lock")="Y" Or Session("User_Valid")="Y" Or Session("Check_IP")="Y" Then Response.End
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta name="GOOGLEBOT" content="NOARCHIVE">
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <title>GOODTV捐款管理系統</title>
  <style type="text/css"><!--
    body {
 	    background-image: url();
	    background-repeat: repeat-x;
	    background-color: #D2E1F0;
    }
    .system{
	    font-size: 16px;
	    color: #FFFFFF;
	    font-weight: bold;
	    font-family: Arial, Helvetica, sans-serif;
	    text-decoration: none;
	    line-height: 22px;
    }
    .login{
	    font-size: 16px;
	    color: #001c58;
	    font-weight: bold;
	    font-family: Arial, Helvetica, sans-serif;
	    text-decoration: none;
	    line-height: 18px;
    }
    .keyin {
	    border: 1px solid #CCCCCC;
	    background-color: #FFFFFF;
	    letter-spacing: normal;
	    vertical-align: baseline;
	    word-spacing: normal;
	    height: auto;
	    width: auto;
	    margin: 1px;
	    padding: 1px;
	    font-size: 10pt;
	    color: #666666;
	    text-indent: 2px;
	    font-family: "Verdana", "Arial", "Helvetica", "sans-serif";
	    text-decoration: none;
    }
    .contact {
	    font-family: "Verdana", "Arial", "Helvetica", "sans-serif";
	    font-size: 10pt;
	    line-height: 12pt;
	    color: #0f83c2;
	    text-decoration: none;

    }
    a.contact:visited {
	    font-family: "Verdana", "Arial", "Helvetica", "sans-serif";
	    font-size: 10pt;
	    line-height: 12pt;
	    color: #0F83C2;
	    text-decoration: none;
    }
    a.contact:hover {
	    font-family: "Verdana", "Arial", "Helvetica", "sans-serif";
	    font-size: 10pt;
	    line-height: 12pt;
	    color: #A70500;
	    text-decoration: none;
    }
    a.contact:active {
	    font-family: "Verdana", "Arial", "Helvetica", "sans-serif";
	    font-size: 10pt;
	    line-height: 12pt;
	    color: #666666;
	    text-decoration: none;
    }  
  --></style>
</head>
<body OnLoad="Window_OnLoad()" onselectstart="return false;" ondragstart="return false;" oncontextmenu="return false;">
  <form name="form" method="POST" action="login.asp">
    <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="110" background="../images/login_bg_01.gif">
        	<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#0066CC">
            <tr>
              <td width="50" height="60" bgcolor="#FFFFFF" align="right"> </td>
              <td width="750" height="60" bgcolor="#FFFFFF"><img src="../images/donation.gif"></td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td height="24" background="../images/login_bg_02.gif" bgcolor="#FFFFFF">
        	<table width="800" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="80">&nbsp;</td>
              <td width="720"><span class="system">捐款管理系統</span></td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td height="155" background="../images/login_bg_03.gif" bgcolor="#FFFFFF">
        	<table width="800" border="0" cellspacing="2" cellpadding="0">
            <tr>
              <td colspan="4"> </td>
            </tr>
            <tr>
              <td width="180"> </td>
              <td width="65" class="login">帳號 :</td>
              <td width="180"><input name="user_id" type="text" class="keyin" size="20" maxlength="20"></td>
              <td width="375"> </td>
            </tr>
            <tr>
              <td> </td>
              <td class="login">密碼 :</td>
              <td><input name="password" type="password" class="keyin" size="20" maxlength="20"></td>
              <td align="left"><input type="submit" value=" 登 入 " name="action" style="font-family: 新細明體; font-size: 9pt"></td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF">
          <table width="800" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="600"><img src="../images/login_bg_04.gif" width="600" height="64"></td>
              <td width="60" valign="top" background="../images/login_bg_05.gif">
              	<table width="60" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="40"> </td>
                  </tr>
                  <tr>
                    <td align="center"><a href="#" class="contact">技術支援</a></td>
                  </tr>
                </table>
              </td>
              <td width="60" valign="top" background="../images/login_bg_06.gif">
              	<table width="60" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="40"> </td>
                  </tr>
                  <tr>
                     <td align="center"><a href="#" class="contact">聯絡我們</a></td>
                  </tr>
                </table>
              </td>
              <td width="60" valign="top" background="../images/login_bg_07.gif">
              	<table width="60" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td height="40"> </td>
                  </tr>
                  <tr>
                    <td align="center"><a href="#" class="contact">版權宣告</a></td>
                  </tr>
                </table>
              </td>
              <td><img src="../images/login_bg_08.gif" width="32" height="64"></td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td><img src="../images/logo_corp.gif"></td>
      </tr>
    </table>
  </form>
  <%Message()%>
</body>
</html>
<script language="JavaScript"><!--
  function Window_OnLoad(){
    document.form.user_id.focus();
  }
--></script>