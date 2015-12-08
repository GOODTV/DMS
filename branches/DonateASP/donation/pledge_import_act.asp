<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head> 
<body>   
<%
Dim objUpload,UploadFile,objUploadedFile
SQL1="delete from dbo.PLEDGE_SEND_RETURN where Pledge_Id >0"
Set RS1=Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,3
Set RS1=Nothing

Set objUpload = Server.CreateObject("Dundas.Upload.2")
On Error Resume Next
objUpload.MaxFileSize = 3000000
objUpload.UseVirtualDir = True
objUpload.UseUniqueNames = False
objUpload.Save "/upload"

For Each objUploadedFile in objUpload.Files
    If objUploadedFile.Size>0 Then
      SourcePath = objUploadedFile.Originalpath
      SourceName = objUpload.GetFileName(SourcePath)
      UploadName = objUpload.GetFileName(objUploadedFile.path)
      UploadSize = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
      ExtName = objUpload.GetFileExt(objUploadedFile.Originalpath)
	End If
Next
UploadFile="../upload/"&UploadName&""
'response.write UploadFile
Set FS=Server.CreateObject("Scripting.FileSystemObject")
  Set Txt=FS.OpenTextFile(Server.MapPath(UploadFile))
  While Not Txt.atEndOfStream
	Line=Txt.ReadLine
	If Instr(Line,"T")<> 1 Then
		'response.write Line&"<br>"
		Account_No=Trim(mid(Line,20,19))
		'response.write "<br>"&Account_No&"<br>"
		Donate_Amt=Trim0(mid(Line,46,10))
		'response.write "<br>"&Donate_Amt&"<br>"
		Return_Status=Trim(mid(Line,58,2))
		'response.write "<br>"&Return_Status&"<br>"
		Pledge_Id=Trim0(mid(Line,77,8))
		'response.write "<br>"&Pledge_Id&"<br>"
		
		SQL2="select * from dbo.PLEDGE_SEND_RETURN"
		Set RS2=Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,3
		RS2.Addnew
		RS2("Pledge_Id")=Pledge_Id
		RS2("Account_No")=Account_No
		If Return_Status = "00" then
			RS2("Return_Status")="Y"
		Else
			RS2("Return_Status")="N"
		End If
		RS2("Donate_Amt")=Donate_Amt
		RS2.Update
		RS2.Close
		Set RS2=Nothing
	End If
  Wend	
Set objUpload = Nothing
SQL_ChkPledge="Select Count(*) AS Count,SUM(Donate_Amt) AS Sum_Amt From dbo.PLEDGE_SEND_RETURN Where Return_Status='Y'"
Set RS_Chk = Server.CreateObject("ADODB.RecordSet")
RS_Chk.Open SQL_ChkPledge,Conn,1,1
Count=RS_Chk("Count")
Sum_Amt=RS_Chk("Sum_Amt")
%>
<div align="center">
<form name="form" method="post" action="">
<input type="hidden" name="Count" value="<%=Count%>">
<input type="hidden" name="Sum_Amt" value="<%=Sum_Amt%>">
<input onClick="Close_OnClick()" value="關閉視窗" type="button">
</form>
</div>
</body>
</html>
<%
Function Trim0(s)
    Dim re
    Set re = New RegExp
    re.Pattern = "^0+"
    Trim0 = re.Replace(s, "")
End Function
%>
<script language="JavaScript"><!--
var count= document.form.Count.value;
var Sum_Amt= document.form.Sum_Amt.value;
alert("回覆檔資料匯入成功，共"+count+"筆，金額"+Sum_Amt+"元");
function Close_OnClick(){
window.opener.document.form.PledgeId_check.value ="Y";
window.close();
}
--></script>