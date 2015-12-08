<!--#include file="../include/dbfunctionJ.asp"-->
<%
  If request("item")="album_media" Then
    SQL="Select Album_TitleImg From ALBUM_PHOTO Where ser_no='"&request("ser_no")&"'"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    TitleImg= RS("Album_TitleImg")
    RS("Album_TitleImg")=""
  ElseIf request("epaper")="album_media" Then
    SQL="Select ePaper_TitleImg From EPAPER_CONTENT Where ser_no='"&request("ser_no")&"'"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    TitleImg= RS("ePaper_TitleImg")
    RS("ePaper_TitleImg")=""       
  Else
    SQL="Select " & request("item") & "_TitleImg From " & request("item") & " Where ser_no='"&request("ser_no")&"'"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    TitleImg= RS("" & request("item") & "_TitleImg")
    RS("" & request("item") & "_TitleImg")=""
  End If
  RS.Update
  RS.Close
  Set RS=Nothing
  
  SQL="Select * From UPLOAD Where Object_ID='"&request("ser_no")&"' And Ap_Name='"&request("item")&"_title'"
  Call QuerySQL(SQL,RS)
  If Not RS.EOF then
    Ser_ID=RS("Ser_No")
    Object_ID=RS("Object_ID")
    Upload_FileURL=RS("Upload_FileURL")
    Upload_FileURL_Old=RS("Upload_FileURL_Old")
  End If
  RS.Close
  Set RS=Nothing
  SQL="Delete From UPLOAD Where ser_no='"&Ser_ID&"'"
  Set RS=conn.Execute(SQL)
  If Upload_FileURL<>"" Then
    DataFile="../upload/"&Upload_FileURL&""
    Set objFS = Server.CreateObject("Scripting.FileSystemObject")
    If objFS.FileExists(Server.MapPath(DataFile)) Then objFS.DeleteFile Server.Mappath(DataFile)  
    Set objFS = Nothing 
  End If
  If Upload_FileURL_Old<>"" Then
    DataFile_Old="../upload/"&Upload_FileURL_Old&""
    Set objFS = Server.CreateObject("Scripting.FileSystemObject")
    If objFS.FileExists(Server.MapPath(DataFile_Old)) Then objFS.DeleteFile Server.Mappath(DataFile_Old)    
    Set objFS = Nothing
  End If
  session("errnumber")=1
  If request("item")="ads" or request("item")="links" Then
    session("msg")="標題檔案刪除成功 ！"
  Else  
    session("msg")="標題圖檔刪除成功 ！"
  End If
  
  If request("item")="album_media" Then
    response.redirect request("item")&"_edit.asp?code_id="&request("code_id")&"&album_id="&request("album_id")&"&ser_no="&request("ser_no")
  ElseIf request("item")="epaper" Then
    response.redirect "epaper_content_edit.asp?ser_no="&request("ser_no")
  Else
    response.redirect request("item")&"_edit.asp?code_id="&request("code_id")&"&ser_no="&request("ser_no")
  End If
%>