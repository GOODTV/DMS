<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head> 
 <body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
	<form name="Form" method="post" action="pledge_import_act.asp" enctype="multipart/form-data">
	  <input type="hidden" name="act">
	  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
		<tr><td class="td02-c">信用卡批次授權(一般)回覆檔匯入：</td></tr>
		<tr>
		  <td class="td02-c"><input name="AttachFile" size="48" type="file" class="font9">
		  </td>
		</tr>
		<tr><td align="center"><input type="button" value=" 匯入 " name="import" class="addbutton" style="cursor:hand" onClick="Import_OnClick(this.form)"></td></tr>
	  </table>
	</form>
	</center></div></p>
  </body>
</html>
<script language="JavaScript"><!--
function Window_OnLoad(){
  document.form.AttachFile.focus();
}
function Import_OnClick(f){
  if(confirm('您是否確定要匯入TXT授權資料？')){
  f.act.value="import"
  //alert (f.act.value);return;
  //alert (document.f.AttachFile2.value);return;
  f.submit();
  }
}
--></script>