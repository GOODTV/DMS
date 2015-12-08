<!--#include file="../include/dbfunction.asp"-->
<%
  sql="select * from userfile where user_id='"&session("user_id")&"'"
  call QuerySQL (SQL,RS)
  
Sub Edit2Button()
    Response.Write "<button id='update' style='position:relative;left:20;width:45;height:40;font-size:9pt'> <img src='../images/update.gif' width='19' height='20'><br>修改</button>"
End Sub 
 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta name="Microsoft Border" content="none">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">	
  <title>使用者密碼變更</title>
</head>
<body>
<form method="POST" action="password_update.asp" name="form" target="main">
  <input type="hidden" name="action" value>
  <input type="hidden" name="pass" value="<%=rs("password")%>">
  <input type="hidden" name="oldpass1" value="<%=rs("old_password1")%>">
  <input type="hidden" name="tmenu" value="<%=session("menu_id")%>">
  <input type="hidden" name="helpdesk" value="helpdesk/password_update.htm">
  <div align="center">
    <center>
  <table border="1" width="50%" cellspacing="0" style="border-collapse: collapse">
    <tr>
      <td width="100%" class="font11"
      style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25">使用者密碼變更</td>
    </tr>
    <tr>
      <td width="100%" class="font9" style="color: rgb(255,255,255)" bgcolor="#C0C0C0"><table
      border="0" width="100%">
        <tr>
          <td width="24%" class="font9A" align="right">使用者代號 :</td>                                                          
          <td width="76%"><input type="text" name="user_id" size="20" readonly class="font9t" value="<%=rs("user_id")%>"></td>
        </tr>
        <tr>
          <td width="24%" class="font9A" align="right">使用者名稱 :</td>                                                         
          <td width="76%"><input type="text" name="user_name" size="20" readonly class="font9t"  
          value="<%=rs("user_name")%>"></td>
        </tr>
        <tr>
          <td width="24%" class="font9A" align="right">舊密碼 :</td>                                  
          <td width="76%">
          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="10" --><input type="password" name="old_password" size="20" class="font9"  
          maxlength="10"></td>                
        </tr>                
        <tr>                
          <td width="24%" class="font9A" align="right">新密碼 :</td>                                  
          <td width="76%">
          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="10" --><input type="password" name="password" size="20" class="font9"  
          maxlength="10"></td>
        </tr>
        <tr>                
          <td width="24%" class="font9A" align="right">確認密碼 :</td>                                  
          <td width="76%">
          <!--webbot bot="Validation" b-value-required="TRUE" i-maximum-length="10" --><input type="password" name="cpassword" size="20" class="font9"  
          maxlength="10"></td>
        </tr>
        <tr>    
          <td width="100%" class="font9A" align="right" colspan="2">
            <hr>
          </td>     
        </tr>
        <tr>
          <td width="100%" colspan="2"><%Edit2Button()%>
          </td>
        </tr>
      </table> 
      </td> 
    </tr> 
  </table> 
    </center>
  </div>
</form> 
</body> 
</html> 
<!--#include file="../include/dbclose.asp"--> 
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!-- 
sub window_OnLoad 
  form.old_password.focus
end sub

sub update_OnClick
   if Len(form.password.value) < 4 then
      Msgbox "密碼位數要4碼以上"
      form.password.focus
      exit sub
   end if    
  if form.pass.value<>form.old_password.value then
     MsgBox "舊密碼不正確, 請重新輸入 !"
     form.old_password.focus
     exit sub
  end if 
  if form.password.value=form.old_password.value then
     Msgbox "新密碼不可與舊密碼一樣 !"
     form.password.value=""
     form.cpassword.value=""
     form.password.focus
     exit sub
  end if
  if form.password.value<>form.cpassword.value then
     Msgbox "密碼並未確認正確 ! 請確定您所輸入的新密碼與確認密碼是一致的"
     form.password.value=""
     form.cpassword.value=""
     form.password.focus
     exit sub
  end if
  ConfirmUpdate() 
end sub 
 
--></script>