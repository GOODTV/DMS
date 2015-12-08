<%Response.ContentType="text/html; charset=Big5"%>
<!--#include file="../include/dbfunction.asp"-->
<%
  FilePath="../upload/temp"
  SysPath=session("syspath")
  
  Set fs = CreateObject ("Scripting.FileSystemObject") 
  If Not fs.FolderExists (SysPath&"\upload\temp") Then
    fs.CreateFolder(SysPath&"\upload\temp")
  End If
  Set fs = Nothing  
  
  call SourceUpLoad(FilePath,SysPath,UploadName1,UploadSize1,UploadName2,UploadSize2,UploadName3,UploadSize3,UploadName4,UploadSize4,UploadName5,UploadSize5,UploadName6,UploadSize6,UploadName7,UploadSize7,UploadName8,UploadSize8,UploadName9,UploadSize9,UploadName10,UploadSize10,objUpload,request("MaxFileSize"))
  If UploadName1<>"" Then 
    SQL="SOURCE"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Source_RootType")=objUpload.form("Source_RootType")
    RS("Source_FileURL")=UploadName1
    RS("Source_FileSize")=UploadSize1
    RS("Source_RegDate")=now()
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
  If UploadName2<>"" Then 
    SQL="SOURCE"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Source_RootType")=objUpload.form("Source_RootType")
    RS("Source_FileURL")=UploadName2
    RS("Source_FileSize")=UploadSize2
    RS("Source_RegDate")=now()
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
  If UploadName3<>"" Then 
    SQL="SOURCE"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Source_RootType")=objUpload.form("Source_RootType")
    RS("Source_FileURL")=UploadName3
    RS("Source_FileSize")=UploadSize3
    RS("Source_RegDate")=now()
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
  If UploadName4<>"" Then 
    SQL="SOURCE"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Source_RootType")=objUpload.form("Source_RootType")
    RS("Source_FileURL")=UploadName4
    RS("Source_FileSize")=UploadSize4
    RS("Source_RegDate")=now()
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
  If UploadName5<>"" Then 
    SQL="SOURCE"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Source_RootType")=objUpload.form("Source_RootType")
    RS("Source_FileURL")=UploadName5
    RS("Source_FileSize")=UploadSize5
    RS("Source_RegDate")=now()
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
  If UploadName6<>"" Then 
    SQL="SOURCE"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Source_RootType")=objUpload.form("Source_RootType")
    RS("Source_FileURL")=UploadName6
    RS("Source_FileSize")=UploadSize6
    RS("Source_RegDate")=now()
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
  If UploadName7<>"" Then 
    SQL="SOURCE"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Source_RootType")=objUpload.form("Source_RootType")
    RS("Source_FileURL")=UploadName7
    RS("Source_FileSize")=UploadSize7
    RS("Source_RegDate")=now()
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
  If UploadName8<>"" Then 
    SQL="SOURCE"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Source_RootType")=objUpload.form("Source_RootType")
    RS("Source_FileURL")=UploadName8
    RS("Source_FileSize")=UploadSize8
    RS("Source_RegDate")=now()
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
  If UploadName9<>"" Then 
    SQL="SOURCE"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Source_RootType")=objUpload.form("Source_RootType")
    RS("Source_FileURL")=UploadName9
    RS("Source_FileSize")=UploadSize9
    RS("Source_RegDate")=now()
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
  If UploadName10<>"" Then 
    SQL="SOURCE"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Source_RootType")=objUpload.form("Source_RootType")
    RS("Source_FileURL")=UploadName10
    RS("Source_FileSize")=UploadSize10
    RS("Source_RegDate")=now()
    RS.Update
    RS.Close
    Set RS=Nothing
  End If          
  session("errnumber")=1
  session("msg")="檔案上傳成功 !"
  response.redirect "source.asp"
%>