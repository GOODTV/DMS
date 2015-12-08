<%Response.ContentType="text/html; charset=utf-8"%>
<!--#include file="../include/dbfunctionJ.asp"-->
<!--#include file="../include/class.img.asp"-->
<%
  FilePath="../upload/"
  call DundasUpLoad(FilePath,UploadName,UploadSize,objUpload,request("MaxFileSize"))
  If objUpload.form("resize")="Y" Then
    '製造縮圖
    Set img = new IMAGE
    srcFile = FilePath & UploadName
    dstFile = FilePath & "S_" & UploadName
    img.transImage srcFile, dstFile, objUpload.form("img_width"), ""
    Set img = Nothing
    
    '新圖檔寬高
    Set ILIB = server.createobject("Overpower.ImageLib")
    ILIB.PictureSize Server.MapPath(FilePath&UploadName),width,height
    Set ILIB = Nothing
    
    '新圖檔Size 
    FileURL="S_" & UploadName
    Set objFS = Server.CreateObject("Scripting.FileSystemObject")     
    FileSize=FormatNumber((objFS.GetFile(Server.MapPath(FilePath&FileURL)).Size)/1024,2)
    FileURL_Old=UploadName
    FileSize_Old=UploadSize
    Set objFS = Nothing
  Else
    FileURL=UploadName
    FileSize=UploadSize
    FileURL_Old=""
    FileSize_Old=0
  End If

  SQL="UPLOAD"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS.Addnew
  RS("Object_ID")=objUpload.form("ser_no")
  RS("Ap_Name")=objUpload.form("item")
  RS("Attach_Type")="img"
  RS("Upload_FileName")=""
  RS("Upload_FileURL")=FileURL
  RS("Upload_FileSize")=FileSize
  RS("Upload_FileURL_Old")=FileURL_Old
  RS("Upload_FileSize_Old")=FileSize_Old
  RS.Update
  RS.Close
  Set RS=Nothing
    
  session("errnumber")=1
  session("msg")="上傳圖檔成功 ！"
  If objUpload.form("item")="epaper" Then
    response.redirect "epaper_content_edit.asp?ser_no="&objUpload.form("ser_no")  
  Else   
    response.redirect objUpload.form("item")&"_edit.asp?code_id="&objUpload.form("code_id")&"&ser_no="&objUpload.form("ser_no")
  End If
%>