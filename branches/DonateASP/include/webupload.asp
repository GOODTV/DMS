<%
session.Timeout=600
Server.ScriptTimeout=600
set conn=server.createobject("ADODB.Connection")
conn.connectiontimeout=600
conn.commandtimeout=600
conn.Provider="sqloledb"
conn.open "server=(local);uid=sa;pwd=kc1997;database=NPOEWB3"

Function DundasUpLoad(FilePath,UploadName1,UploadSize1,objUpload,MaxFileSize)
  set objUpload = Server.CreateObject("Dundas.Upload")
  on error resume next
  objUpload.MaxFileCount = 1
  objUpload.MaxFileSize = MaxFileSize*1048576
  objUpload.MaxUploadSize = MaxFileSize*1048576
  objUpload.UseUniqueNames = True
  objUpload.UseVirtualDir = True
  objUpload.Save FilePath
  
  For Each objUploadedFile in objUpload.Files
    If objUploadedFile.Size>0 Then
      SourcePath1 = objUploadedFile.Originalpath
      SourceName1 = objUpload.GetFileName(objUploadedFile.Originalpath)
      UploadName1 = objUpload.GetFileName(objUploadedFile.path)
      UploadSize1 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
      ExtName1 = objUpload.GetFileExt(objUploadedFile.Originalpath)
    End If
  Next
End Function
%>