<!--#include file="../include/dbfunctionJ.asp"-->
<%
  SQL="Select * From UPLOAD Where ser_no='"&request("ser_no")&"'"
  Call QuerySQL(SQL,RS)
  If Not RS.EOF then
    Object_ID=RS("Object_ID")
    Attach_Type=RS("Attach_Type")
    Upload_FileURL=RS("Upload_FileURL")
    Upload_FileURL_Old=RS("Upload_FileURL_Old")
  End If
  RS.Close
  Set RS=Nothing
  SQL="Delete From UPLOAD Where ser_no="&request("ser_no")&""
  Set RS=conn.Execute(SQL)
  If Attach_Type<>"YouTube" And Attach_Type<>"無名小站" And Attach_Type<>"WMV" Then
    DataFile="../upload/"&Upload_FileURL&""
    Set objFS = Server.CreateObject("Scripting.FileSystemObject")
    If objFS.FileExists(Server.MapPath(DataFile)) Then objFS.DeleteFile Server.Mappath(DataFile)
    If Upload_FileURL_Old<>"" Then
      DataFile_Old="../upload/"&Upload_FileURL_Old&""
      If objFS.FileExists(Server.MapPath(DataFile_Old)) Then objFS.DeleteFile Server.Mappath(DataFile_Old)
    End If
    Set objFS = Nothing
  End If
  session("errnumber")=1
  session("msg")="檔案刪除成功 !"
  If request("item")="epaper" Then
    response.redirect "epaper_content_edit.asp?ser_no="&Object_ID
  Else
    response.redirect request("item")&"_edit.asp?code_id="&request("code_id")&"&ser_no="&Object_ID
  End If
%>