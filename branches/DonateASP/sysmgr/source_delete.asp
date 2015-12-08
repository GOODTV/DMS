<!--#include file="../include/dbfunction.asp"-->
<%
SQL="Select * From SOURCE Where Ser_No='"&request("Ser_No")&"'"
Call QuerySQL(SQL,RS)
If Not RS.EOF Then
  Source_RootType=RS("Source_RootType")
  Source_FileURL=RS("Source_FileURL")
End If
RS.close
Set RS=nothing
'刪除檔案
SQL="Delete From SOURCE Where Ser_No='"&request("Ser_No")&"'"
call ExecSQL(SQL)
'刪除source
If Source_FileURL<>"" Then
  If Source_RootType<>"" Then
    DataFile="../"&Source_RootType&"/"&Source_FileURL&""
  Else
    DataFile="../"&Source_FileURL&""
  End IF
  Set objFS = Server.CreateObject("Scripting.FileSystemObject")
  If objFS.FileExists(Server.MapPath(DataFile)) Then objFS.DeleteFile Server.Mappath(DataFile)
  Set objFS = Nothing 
End If
session("errnumber")=1
session("msg")="檔案刪除成功！"
response.redirect "source_list.asp"
%>