<%Response.ContentType="text/html; charset=utf-8"%>
<!--#include file="../include/dbfunction.asp"-->

<%
   FilePath="../upload/"
   call FileUpLoad(FilePath,FName,Fsize,objUpload)
   sql="update SEAL set Seal_TitleImg='"&FName&"' where Ser_No='"&objUpload.form("Ser_No")&"'"
   Set RS1=Conn.Execute(SQL)   
   session("errnumber")=1
   session("msg")=objUpload.form("SealSubject")&"\n電子印鑑圖檔上傳成功 !"
   response.redirect "seal_edit.asp?ser_no="&objUpload.form("Ser_No")  
%>