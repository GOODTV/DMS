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
    If objUpload.form("img_height")<>"" Then
      img.transImage srcFile, dstFile,"", objUpload.form("img_height")
    Else
      img.transImage srcFile, dstFile, objUpload.form("img_width"),""
    End If
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
  
  If objUpload.form("item")="album_media" Then
    SQL="Select Album_TitleImg From ALBUM_PHOTO Where Ser_No='"&objUpload.form("ser_no")&"'"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3 
    RS("Album_TitleImg")=FileURL
    RS.Update
    RS.Close
    Set RS=Nothing
  
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")&"_title"
    RS("Attach_Type")="image"
    RS("Upload_FileName")=""
    RS("Upload_FileURL")=FileURL
    RS("Upload_FileSize")=FileSize
    RS("Upload_FileURL_Old")=FileURL_Old
    RS("Upload_FileSize_Old")=FileSize_Old
    RS.Update
    RS.Close
    Set RS=Nothing
    
    session("errnumber")=1
    session("msg")="上傳標題圖檔成功 ！"
    response.redirect objUpload.form("item")&"_edit.asp?code_id="&objUpload.form("code_id")&"&album_id="&objUpload.form("album_id")&"&ser_no="&objUpload.form("ser_no")
  ElseIf objUpload.form("item")="epaper" Then
    SQL="Select ePaper_TitleImg From EPAPER_CONTENT Where Ser_No='"&objUpload.form("ser_no")&"'"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3 
    RS("ePaper_TitleImg")=FileURL
    RS.Update
    RS.Close
    Set RS=Nothing
  
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")&"_title"
    RS("Attach_Type")="image"
    RS("Upload_FileName")=""
    RS("Upload_FileURL")=FileURL
    RS("Upload_FileSize")=FileSize
    RS("Upload_FileURL_Old")=FileURL_Old
    RS("Upload_FileSize_Old")=FileSize_Old
    RS.Update
    RS.Close
    Set RS=Nothing
    
    session("errnumber")=1
    session("msg")="上傳標題圖檔成功 ！"
    response.redirect objUpload.form("item")&"_content_edit.asp?code_id="&objUpload.form("code_id")&"&album_id="&objUpload.form("album_id")&"&ser_no="&objUpload.form("ser_no")
  Else
    SQL="Select " & objUpload.form("item") & "_TitleImg From " & objUpload.form("item") & " Where Ser_No='"&objUpload.form("ser_no")&"'"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3 
    RS("" & objUpload.form("item") & "_TitleImg")=FileURL
    RS.Update
    RS.Close
    Set RS=Nothing
  
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")&"_title"
    RS("Attach_Type")="image"
    RS("Upload_FileName")=""
    RS("Upload_FileURL")=FileURL
    RS("Upload_FileSize")=FileSize
    RS("Upload_FileURL_Old")=FileURL_Old
    RS("Upload_FileSize_Old")=FileSize_Old
    RS.Update
    RS.Close
    Set RS=Nothing
    
    session("errnumber")=1
    session("msg")="上傳標題圖檔成功 ！"
    response.redirect objUpload.form("item")&"_edit.asp?code_id="&objUpload.form("code_id")&"&ser_no="&objUpload.form("ser_no")  
  End If
%>