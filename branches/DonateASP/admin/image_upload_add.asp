﻿<%Response.ContentType="text/html; charset=utf-8"%>
<!--#include file="../include/dbfunctionJ.asp"-->
<!--#include file="../include/class.img.asp"-->
<%
  If request("album_id")<>"" Then
    FilePath="../upload/album_"&request("album_id")&"/"
    datafile="../upload/album_"&request("album_id")
    Set fso = CreateObject("Scripting.FileSystemObject")
    foldername=Server.MapPath(datafile)
    If fso.FolderExists(foldername)=false Then
      Set fld = fso.CreateFolder(foldername)
    End If
  Else
    FilePath="../upload/"
  End If
  call DundasUpLoad(FilePath,UploadName,UploadSize,objUpload,request("MaxFileSize"))
  If objUpload.form("resize")="Y" And objUpload.form("type")="image" And (objUpload.form("img_height")<>"" Or objUpload.form("img_width")<>"") Then
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
  
  If objUpload.form("item")="ads" Or objUpload.form("item")="links" Then
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
  ElseIf objUpload.form("item")="album_photo" Then
    SQL="ALBUM_PHOTO"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Album_Id")=objUpload.form("ser_no")
    RS("Album_Date")=objUpload.form("albumdate")
    RS("Album_Desc")=objUpload.form("filename")
    RS("Album_URL")=FileURL
    RS("Album_FileSize")=FileSize
    RS("Album_URL_Old")=FileURL_Old
    RS("Album_FileSize_Old")=FileSize_Old
    RS("Album_Location")=""
    RS("Album_Type")="image"
    RS("Album_Seq")=objUpload.form("albumseq")
    RS.Update
    RS.Close
    Set RS=Nothing

    SQL="Select @@IDENTITY As Ser_No"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,1
    If Not RS.EOF then
      Object_ID=RS("Ser_No")
    End If
    RS.Close
    Set RS=Nothing 
        
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=Object_ID
    RS("Ap_Name")=objUpload.form("item")
    RS("Attach_Type")="image"
    RS("Upload_FileName")=""
    RS("Upload_FileURL")=FileURL
    RS("Upload_FileSize")=FileSize
    RS("Upload_FileURL_Old")=FileURL_Old
    RS("Upload_FileSize_Old")=FileSize_Old
    RS.Update
    RS.Close
    Set RS=Nothing
  Else
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
  End If
        
  session("errnumber")=1
  If objUpload.form("item")="ads" Or objUpload.form("item")="links" Then
    session("msg")="標題檔案上傳成功 ！"
    response.redirect objUpload.form("item")&"_edit.asp?code_id="&objUpload.form("code_id")&"&ser_no="&objUpload.form("ser_no")
  ElseIf objUpload.form("item")="album_photo" Then
    session("msg")="相片檔案上傳成功 ！"
    response.redirect objUpload.form("item")&".asp?code_id="&objUpload.form("code_id")&"&ser_no="&objUpload.form("ser_no")
  ElseIf objUpload.form("item")="signup" Then
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