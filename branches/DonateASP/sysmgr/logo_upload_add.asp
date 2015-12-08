<%Response.ContentType="text/html; charset=utf-8"%>
<!--#include file="../include/dbfunction.asp"-->
<!--#include file ="../include/upload.inc"-->

<%
   FilePath="../image/"
   call UpLoad(FilePath,FName,Fsize,upfile)
   sql="update dept set logo='"&FName&"' where dept_id='"&upfile.form("dept_id")&"'"
   Set RS1=Conn.Execute(SQL)   
   session("errnumber")=1
   session("msg")="檔案上傳成功 !"
   response.redirect "dept_edit.asp?dept_id="&upfile.form("dept_id")  
%>