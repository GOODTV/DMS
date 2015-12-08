<!--#include file="../include/dbfunctionJ.asp"-->
<% 
  SQL="Select * From UPLOAD Where ser_no='"&request("ser_no")&"'"
  Call QuerySQL(SQL,RS)
  If Not RS.EOF then
    Object_ID=RS("Object_ID")
    Upload_FileURL=RS("Upload_FileURL")
    Upload_FileURL_Old=RS("Upload_FileURL_Old")
  End If
  RS.Close
  Set RS=Nothing
  SQL="Delete From UPLOAD Where ser_no="&request("ser_no")&""
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
  session("msg")="圖檔刪除成功 ！"
  response.redirect "image_left_list.asp?code_id="&request("code_id")&"&item="&request("item")&"&ser_no="&Object_ID
%>