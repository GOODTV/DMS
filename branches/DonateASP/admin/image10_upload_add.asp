<%Response.ContentType="text/html; charset=utf-8"%>
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
  call DundasUpLoad10(FilePath,UploadName1,UploadSize1,ExtName1,UploadName2,UploadSize2,ExtName2,UploadName3,UploadSize3,ExtName3,UploadName4,UploadSize4,ExtName4,UploadName5,UploadSize5,ExtName5,UploadName6,UploadSize6,ExtName6,UploadName7,UploadSize7,ExtName7,UploadName8,UploadSize8,ExtName8,UploadName9,UploadSize9,ExtName9,UploadName10,UploadSize10,ExtName10,objUpload,request("MaxFileSize"))
  If objUpload.form("resize")="Y" And objUpload.form("type")="image" And (objUpload.form("img_height")<>"" Or objUpload.form("img_width")<>"") Then
    If UploadName1<>"" Then
      '製造縮圖
      Set img = new IMAGE
      srcFile = FilePath & UploadName1
      dstFile = FilePath & "S_" & UploadName1
      If objUpload.form("img_height")<>"" Then
        img.transImage srcFile, dstFile,"", objUpload.form("img_height")
      Else
        img.transImage srcFile, dstFile, objUpload.form("img_width"),""
      End If
      Set img = Nothing
    
      '新圖檔寬高
      Set ILIB = server.createobject("Overpower.ImageLib")
      ILIB.PictureSize Server.MapPath(FilePath&UploadName1),width,height
      Set ILIB = Nothing
    
      '新圖檔Size 
      FileURL1="S_" & UploadName1
      Set objFS = Server.CreateObject("Scripting.FileSystemObject")
      FileSize1=FormatNumber((objFS.GetFile(Server.MapPath(FilePath&FileURL1)).Size)/1024,2)
      FileURL1_Old=UploadName1
      FileSize1_Old=UploadSize1
      Set objFS = Nothing
    Else
      FileURL1=UploadName1
      FileSize1=UploadSize1
      FileURL1_Old=""
      FileSize1_Old=0
    End If
    
    If UploadName2<>"" Then
      '製造縮圖
      Set img = new IMAGE
      srcFile = FilePath & UploadName2
      dstFile = FilePath & "S_" & UploadName2
      If objUpload.form("img_height")<>"" Then
        img.transImage srcFile, dstFile,"", objUpload.form("img_height")
      Else
        img.transImage srcFile, dstFile, objUpload.form("img_width"),""
      End If
      Set img = Nothing
    
      '新圖檔寬高
      Set ILIB = server.createobject("Overpower.ImageLib")
      ILIB.PictureSize Server.MapPath(FilePath&UploadName2),width,height
      Set ILIB = Nothing
    
      '新圖檔Size 
      FileURL2="S_" & UploadName2
      Set objFS = Server.CreateObject("Scripting.FileSystemObject")
      FileSize2=FormatNumber((objFS.GetFile(Server.MapPath(FilePath&FileURL2)).Size)/1024,2)
      FileURL2_Old=UploadName2
      FileSize2_Old=UploadSize2
      Set objFS = Nothing
    Else
      FileURL2=UploadName2
      FileSize2=UploadSize2
      FileURL2_Old=""
      FileSize2_Old=0
    End If
    
    If UploadName3<>"" Then
      '製造縮圖
      Set img = new IMAGE
      srcFile = FilePath & UploadName3
      dstFile = FilePath & "S_" & UploadName3
      If objUpload.form("img_height")<>"" Then
        img.transImage srcFile, dstFile,"", objUpload.form("img_height")
      Else
        img.transImage srcFile, dstFile, objUpload.form("img_width"),""
      End If
      Set img = Nothing
    
      '新圖檔寬高
      Set ILIB = server.createobject("Overpower.ImageLib")
      ILIB.PictureSize Server.MapPath(FilePath&UploadName3),width,height
      Set ILIB = Nothing
    
      '新圖檔Size 
      FileURL3="S_" & UploadName3
      Set objFS = Server.CreateObject("Scripting.FileSystemObject")
      FileSize3=FormatNumber((objFS.GetFile(Server.MapPath(FilePath&FileURL3)).Size)/1024,2)
      FileURL3_Old=UploadName3
      FileSize3_Old=UploadSize3
      Set objFS = Nothing
    Else
      FileURL3=UploadName3
      FileSize3=UploadSize3
      FileURL3_Old=""
      FileSize3_Old=0
    End If
    
    If UploadName4<>"" Then
      '製造縮圖
      Set img = new IMAGE
      srcFile = FilePath & UploadName4
      dstFile = FilePath & "S_" & UploadName4
      If objUpload.form("img_height")<>"" Then
        img.transImage srcFile, dstFile,"", objUpload.form("img_height")
      Else
        img.transImage srcFile, dstFile, objUpload.form("img_width"),""
      End If
      Set img = Nothing
    
      '新圖檔寬高
      Set ILIB = server.createobject("Overpower.ImageLib")
      ILIB.PictureSize Server.MapPath(FilePath&UploadName4),width,height
      Set ILIB = Nothing
    
      '新圖檔Size 
      FileURL4="S_" & UploadName4
      Set objFS = Server.CreateObject("Scripting.FileSystemObject")
      FileSize4=FormatNumber((objFS.GetFile(Server.MapPath(FilePath&FileURL4)).Size)/1024,2)
      FileURL4_Old=UploadName4
      FileSize4_Old=UploadSize4
      Set objFS = Nothing
    Else
      FileURL4=UploadName4
      FileSize4=UploadSize4
      FileURL4_Old=""
      FileSize4_Old=0
    End If
    
    If UploadName5<>"" Then
      '製造縮圖
      Set img = new IMAGE
      srcFile = FilePath & UploadName5
      dstFile = FilePath & "S_" & UploadName5
      If objUpload.form("img_height")<>"" Then
        img.transImage srcFile, dstFile,"", objUpload.form("img_height")
      Else
        img.transImage srcFile, dstFile, objUpload.form("img_width"),""
      End If
      Set img = Nothing
    
      '新圖檔寬高
      Set ILIB = server.createobject("Overpower.ImageLib")
      ILIB.PictureSize Server.MapPath(FilePath&UploadName5),width,height
      Set ILIB = Nothing
    
      '新圖檔Size 
      FileURL5="S_" & UploadName5
      Set objFS = Server.CreateObject("Scripting.FileSystemObject")
      FileSize5=FormatNumber((objFS.GetFile(Server.MapPath(FilePath&FileURL5)).Size)/1024,2)
      FileURL5_Old=UploadName5
      FileSize5_Old=UploadSize5
      Set objFS = Nothing
    Else
      FileURL5=UploadName5
      FileSize5=UploadSize5
      FileURL5_Old=""
      FileSize5_Old=0
    End If
    
    If UploadName6<>"" Then
      '製造縮圖
      Set img = new IMAGE
      srcFile = FilePath & UploadName6
      dstFile = FilePath & "S_" & UploadName6
      If objUpload.form("img_height")<>"" Then
        img.transImage srcFile, dstFile,"", objUpload.form("img_height")
      Else
        img.transImage srcFile, dstFile, objUpload.form("img_width"),""
      End If
      Set img = Nothing
    
      '新圖檔寬高
      Set ILIB = server.createobject("Overpower.ImageLib")
      ILIB.PictureSize Server.MapPath(FilePath&UploadName6),width,height
      Set ILIB = Nothing
    
      '新圖檔Size 
      FileURL6="S_" & UploadName6
      Set objFS = Server.CreateObject("Scripting.FileSystemObject")
      FileSize6=FormatNumber((objFS.GetFile(Server.MapPath(FilePath&FileURL6)).Size)/1024,2)
      FileURL6_Old=UploadName6
      FileSize6_Old=UploadSize6
      Set objFS = Nothing
    Else
      FileURL6=UploadName6
      FileSize6=UploadSize6
      FileURL6_Old=""
      FileSize6_Old=0
    End If
    
    If UploadName7<>"" Then
      '製造縮圖
      Set img = new IMAGE
      srcFile = FilePath & UploadName7
      dstFile = FilePath & "S_" & UploadName7
      If objUpload.form("img_height")<>"" Then
        img.transImage srcFile, dstFile,"", objUpload.form("img_height")
      Else
        img.transImage srcFile, dstFile, objUpload.form("img_width"),""
      End If
      Set img = Nothing
    
      '新圖檔寬高
      Set ILIB = server.createobject("Overpower.ImageLib")
      ILIB.PictureSize Server.MapPath(FilePath&UploadName7),width,height
      Set ILIB = Nothing
    
      '新圖檔Size 
      FileURL7="S_" & UploadName7
      Set objFS = Server.CreateObject("Scripting.FileSystemObject")
      FileSize7=FormatNumber((objFS.GetFile(Server.MapPath(FilePath&FileURL7)).Size)/1024,2)
      FileURL7_Old=UploadName7
      FileSize7_Old=UploadSize7
      Set objFS = Nothing
    Else
      FileURL7=UploadName7
      FileSize7=UploadSize7
      FileURL7_Old=""
      FileSize7_Old=0
    End If
    
    If UploadName8<>"" Then
      '製造縮圖
      Set img = new IMAGE
      srcFile = FilePath & UploadName8
      dstFile = FilePath & "S_" & UploadName8
      If objUpload.form("img_height")<>"" Then
        img.transImage srcFile, dstFile,"", objUpload.form("img_height")
      Else
        img.transImage srcFile, dstFile, objUpload.form("img_width"),""
      End If
      Set img = Nothing
    
      '新圖檔寬高
      Set ILIB = server.createobject("Overpower.ImageLib")
      ILIB.PictureSize Server.MapPath(FilePath&UploadName8),width,height
      Set ILIB = Nothing
    
      '新圖檔Size 
      FileURL8="S_" & UploadName8
      Set objFS = Server.CreateObject("Scripting.FileSystemObject")
      FileSize8=FormatNumber((objFS.GetFile(Server.MapPath(FilePath&FileURL8)).Size)/1024,2)
      FileURL8_Old=UploadName8
      FileSize8_Old=UploadSize8
      Set objFS = Nothing
    Else
      FileURL8=UploadName8
      FileSize8=UploadSize8
      FileURL8_Old=""
      FileSize8_Old=0
    End If
    
    If UploadName9<>"" Then
      '製造縮圖
      Set img = new IMAGE
      srcFile = FilePath & UploadName9
      dstFile = FilePath & "S_" & UploadName9
      If objUpload.form("img_height")<>"" Then
        img.transImage srcFile, dstFile,"", objUpload.form("img_height")
      Else
        img.transImage srcFile, dstFile, objUpload.form("img_width"),""
      End If
      Set img = Nothing
    
      '新圖檔寬高
      Set ILIB = server.createobject("Overpower.ImageLib")
      ILIB.PictureSize Server.MapPath(FilePath&UploadName9),width,height
      Set ILIB = Nothing
    
      '新圖檔Size 
      FileURL9="S_" & UploadName9
      Set objFS = Server.CreateObject("Scripting.FileSystemObject")
      FileSize9=FormatNumber((objFS.GetFile(Server.MapPath(FilePath&FileURL9)).Size)/1024,2)
      FileURL9_Old=UploadName9
      FileSize9_Old=UploadSize9
      Set objFS = Nothing
    Else
      FileURL9=UploadName9
      FileSize9=UploadSize9
      FileURL9_Old=""
      FileSize9_Old=0
    End If
   
    If UploadName10<>"" Then
      '製造縮圖
      Set img = new IMAGE
      srcFile = FilePath & UploadName10
      dstFile = FilePath & "S_" & UploadName10
      If objUpload.form("img_height")<>"" Then
        img.transImage srcFile, dstFile,"", objUpload.form("img_height")
      Else
        img.transImage srcFile, dstFile, objUpload.form("img_width"),""
      End If
      Set img = Nothing
    
      '新圖檔寬高
      Set ILIB = server.createobject("Overpower.ImageLib")
      ILIB.PictureSize Server.MapPath(FilePath&UploadName10),width,height
      Set ILIB = Nothing
    
      '新圖檔Size 
      FileURL10="S_" & UploadName10
      Set objFS = Server.CreateObject("Scripting.FileSystemObject")
      FileSize10=FormatNumber((objFS.GetFile(Server.MapPath(FilePath&FileURL10)).Size)/1024,2)
      FileURL10_Old=UploadName10
      FileSize10_Old=UploadSize10
      Set objFS = Nothing
    Else
      FileURL10=UploadName10
      FileSize10=UploadSize10
      FileURL10_Old=""
      FileSize10_Old=0
    End If                  
  Else
    FileURL1=UploadName1
    FileSize1=UploadSize1
    FileURL1_Old=""
    FileSize1_Old=0
    
    FileURL2=UploadName2
    FileSize2=UploadSize2
    FileURL2_Old=""
    FileSize2_Old=0
    
    FileURL3=UploadName3
    FileSize3=UploadSize3
    FileURL3_Old=""
    FileSize3_Old=0
    
    FileURL4=UploadName4
    FileSize4=UploadSize4
    FileURL4_Old=""
    FileSize4_Old=0
    
    FileURL5=UploadName5
    FileSize5=UploadSize5
    FileURL5_Old=""
    FileSize5_Old=0
    
    FileURL6=UploadName6
    FileSize6=UploadSize6
    FileURL6_Old=""
    FileSize6_Old=0
    
    FileURL7=UploadName7
    FileSize7=UploadSize7
    FileURL7_Old=""
    FileSize7_Old=0
    
    FileURL8=UploadName8
    FileSize8=UploadSize8
    FileURL8_Old=""
    FileSize8_Old=0
    
    FileURL9=UploadName9
    FileSize9=UploadSize9
    FileURL9_Old=""
    FileSize9_Old=0
    
    FileURL10=UploadName10
    FileSize10=UploadSize10
    FileURL10_Old=""
    FileSize10_Old=0     
  End If
  
  If FileURL1<>"" Then
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")
    RS("Attach_Type")=objUpload.form("type")
    RS("Upload_FileName")=objUpload.form("filename1")
    RS("Upload_FileURL")=FileURL1
    RS("Upload_FileSize")=FileSize1
    RS("Upload_FileURL_Old")=FileURL1_Old
    RS("Upload_FileSize_Old")=FileSize1_Old
    RS.Update
    RS.Close
    Set RS=Nothing
  End If

  If FileURL2<>"" Then
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")
    RS("Attach_Type")=objUpload.form("type")
    RS("Upload_FileName")=objUpload.form("filename2")
    RS("Upload_FileURL")=FileURL2
    RS("Upload_FileSize")=FileSize2
    RS("Upload_FileURL_Old")=FileURL2_Old
    RS("Upload_FileSize_Old")=FileSize2_Old
    RS.Update
    RS.Close
    Set RS=Nothing
  End If

  If FileURL3<>"" Then
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")
    RS("Attach_Type")=objUpload.form("type")
    RS("Upload_FileName")=objUpload.form("filename3")
    RS("Upload_FileURL")=FileURL3
    RS("Upload_FileSize")=FileSize3
    RS("Upload_FileURL_Old")=FileURL3_Old
    RS("Upload_FileSize_Old")=FileSize3_Old
    RS.Update
    RS.Close
    Set RS=Nothing
  End If

  If FileURL4<>"" Then
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")
    RS("Attach_Type")=objUpload.form("type")
    RS("Upload_FileName")=objUpload.form("filename4")
    RS("Upload_FileURL")=FileURL4
    RS("Upload_FileSize")=FileSize4
    RS("Upload_FileURL_Old")=FileURL4_Old
    RS("Upload_FileSize_Old")=FileSize4_Old
    RS.Update
    RS.Close
    Set RS=Nothing
  End If

  If FileURL5<>"" Then
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")
    RS("Attach_Type")=objUpload.form("type")
    RS("Upload_FileName")=objUpload.form("filename5")
    RS("Upload_FileURL")=FileURL5
    RS("Upload_FileSize")=FileSize5
    RS("Upload_FileURL_Old")=FileURL5_Old
    RS("Upload_FileSize_Old")=FileSize5_Old
    RS.Update
    RS.Close
    Set RS=Nothing
  End If  

  If FileURL6<>"" Then
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")
    RS("Attach_Type")=objUpload.form("type")
    RS("Upload_FileName")=objUpload.form("filename6")
    RS("Upload_FileURL")=FileURL6
    RS("Upload_FileSize")=FileSize6
    RS("Upload_FileURL_Old")=FileURL6_Old
    RS("Upload_FileSize_Old")=FileSize6_Old
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
  
  If FileURL7<>"" Then
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")
    RS("Attach_Type")=objUpload.form("type")
    RS("Upload_FileName")=objUpload.form("filename7")
    RS("Upload_FileURL")=FileURL7
    RS("Upload_FileSize")=FileSize7
    RS("Upload_FileURL_Old")=FileURL7_Old
    RS("Upload_FileSize_Old")=FileSize7_Old
    RS.Update
    RS.Close
    Set RS=Nothing
  End If

  If FileURL8<>"" Then
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")
    RS("Attach_Type")=objUpload.form("type")
    RS("Upload_FileName")=objUpload.form("filename8")
    RS("Upload_FileURL")=FileURL8
    RS("Upload_FileSize")=FileSize8
    RS("Upload_FileURL_Old")=FileURL8_Old
    RS("Upload_FileSize_Old")=FileSize8_Old
    RS.Update
    RS.Close
    Set RS=Nothing
  End If  

  If FileURL9<>"" Then
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")
    RS("Attach_Type")=objUpload.form("type")
    RS("Upload_FileName")=objUpload.form("filename9")
    RS("Upload_FileURL")=FileURL9
    RS("Upload_FileSize")=FileSize9
    RS("Upload_FileURL_Old")=FileURL9_Old
    RS("Upload_FileSize_Old")=FileSize9_Old
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
  
  If FileURL10<>"" Then
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=objUpload.form("ser_no")
    RS("Ap_Name")=objUpload.form("item")
    RS("Attach_Type")=objUpload.form("type")
    RS("Upload_FileName")=objUpload.form("filename10")
    RS("Upload_FileURL")=FileURL10
    RS("Upload_FileSize")=FileSize10
    RS("Upload_FileURL_Old")=FileURL10_Old
    RS("Upload_FileSize_Old")=FileSize10_Old
    RS.Update
    RS.Close
    Set RS=Nothing
  End If
                
  session("errnumber")=1
  session("msg")="附加檔案上傳成功 ！"
  If objUpload.form("item")="epaper" Then
    response.redirect "epaper_content_edit.asp?ser_no="&objUpload.form("ser_no")  
  Else
    response.redirect objUpload.form("item")&"_edit.asp?code_id="&objUpload.form("code_id")&"&ser_no="&objUpload.form("ser_no")
  End If
%>