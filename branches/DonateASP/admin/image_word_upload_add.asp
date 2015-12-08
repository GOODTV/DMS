<%Response.ContentType="text/html; charset=utf-8"%>
<!--#include file="../include/dbfunctionJ.asp"-->
<!--#include file="../include/class.img.asp"-->
<%
  FilePath="../upload/"
  call DundasUpLoad(FilePath,UploadName,UploadSize,objUpload,request("MaxFileSize"))
  FileURL=UploadName
  FileSize=UploadSize
  FileURL_Old=""
  FileSize_Old=0
  
  SQL="UPLOAD"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS.Addnew
  RS("Object_ID")=objUpload.form("ser_no")
  RS("Ap_Name")=objUpload.form("item")
  RS("Attach_Type")=objUpload.form("type")
  RS("Upload_FileName")=objUpload.form("filename")
  RS("Upload_FileURL")=FileURL
  RS("Upload_FileSize")=FileSize
  RS("Upload_FileURL_Old")=FileURL_Old
  RS("Upload_FileSize_Old")=FileSize_Old
  RS.Update
  RS.Close
  Set RS=Nothing
        
  session("errnumber")=1
  If objUpload.form("item")="signup" Then
    session("msg")="報名簡章上傳成功 ！"
    response.redirect "activity_edit.asp?code_id="&objUpload.form("code_id")&"&ser_no="&objUpload.form("ser_no")
  ElseIf objUpload.form("item")="epaper" Then
    session("msg")="附加檔案上傳成功 ！"
    response.redirect "epaper_content_edit.asp?ser_no="&objUpload.form("ser_no")    
  Else
    session("msg")="附加檔案上傳成功 ！"
    response.redirect objUpload.form("item")&"_edit.asp?code_id="&objUpload.form("code_id")&"&ser_no="&objUpload.form("ser_no")
  End If
%>