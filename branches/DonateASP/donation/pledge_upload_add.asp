<%Response.ContentType="text/html; charset=utf-8"%>
<!--#include file="../include/dbfunctionJ.asp"-->
<%
  FilePath="../upload/"
  call DundasUpLoad(FilePath,UploadName,UploadSize,objUpload,request("MaxFileSize"))
  If UploadName<>"" Then 
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("Object_ID")
    RS("Ap_Name")=objUpload.form("Ap_Name")
    RS("Attach_Type")="doc"
    RS("Upload_FileName")=objUpload.form("Upload_FileName")
    RS("Upload_FileURL")=UploadName
    RS("Upload_FileSize")=UploadSize
    RS("Upload_FileURL_Old")=""
    RS("Upload_FileSize_Old")="0"
    RS.Update
    RS.Close
    Set RS=Nothing    
  End If
  session("errnumber")=1
  session("msg")="轉帳授權書上傳成功 ！"
  response.redirect "pledge_upload.asp"             
%>