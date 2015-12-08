
<!--#include file="../global.asp"-->
<%
response.buffer =true
If instr(Mid(Request.ServerVariables("URL"),InstrRev(Request.ServerVariables("URL"),"/")+1),"sending.asp")=0 Then response.expires=-1
If session("user_id")="" Or Session("dept_id")="" Or Session("dept_id_login")=""  Or Session("all_dept_type")="" Then
  session.abandon
  response.redirect("../sysmgr/timeout.asp")
End If
session.Timeout=600
Server.ScriptTimeout=600
set conn=server.createobject("ADODB.Connection")
conn.connectiontimeout=600
conn.commandtimeout=600
conn.Provider="sqloledb"
conn.open "server="&session("server")&";uid="&session("uid")&";pwd="&session("pwd")&";database="&session("database")
  
Sub Warning (msg)
  Response.Write "<script Language='JavaScript'> alert('" & msg & "')"
  Response.Write "</script>"
End Sub 

Sub Message()
  If session("errnumber")=0 then
    Response.Write "<center>"&session("msg")&"</center>"
  Else
    Response.Write "<script Language='JavaScript'> alert('" & session("msg")&"')</script>"
  End If
  session("msg")=""
  session("errnumber")=0
End Sub

Function QuerySQL (SQL,RS)
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,1
End Function

Function AddSQL (SQL)
  On Error Resume Next
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,3
  Conn.Close
  Select case err.number
    case 0 session("msg")="訊息 : 新增資料成功 !"
    case -2147217900 session("msg")="錯誤訊息 : 資料已經存在, 無法新增 !"
    case -2147217873 session("msg")="錯誤訊息 : 資料已經存在, 無法新增 !"
    case else
      session("msg")="錯誤 : " & err.number & "  " & err.description
  End Select
  session("errnumber")=err.number
End Function

Function UpdateSQL (SQL)
  On Error Resume Next
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,conn,1,3
  Conn.Close
  Select case err.number
    case 0 session("msg")="訊息 : 異動資料成功 !"
    case -2147217873 session("msg")="錯誤訊息 : 資料已經存在, 無法新增 !"
    case else
      session("msg")="錯誤 : " & err.number & "  " & err.description
  End Select
  session("errnumber")=err.number
End Function

Function ExecSQL (SQL)
  On Error Resume Next
  Set RS=Conn.Execute(SQL)
  Select case err.number
    case 0 session("msg")="訊息 : 異動資料成功 !"
    case -2147217900 session("msg")="錯誤訊息 : 資料已經存在, 無法新增 !"
    case -2147217873 session("msg")="錯誤訊息 : 資料已經存在, 無法新增 !"
    case else
      session("msg")="錯誤 : " & err.number & "  " & err.description
  End Select
  session("errnumber")=err.number
End Function

Function CheckStringJ (FName,ListField)
  Response.Write "if(document.form." & FName & ".value==''){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位不可為空白！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function ChecklenJ(FName,CLength,ListField)'檢測中英文夾雜字串實際長度
  Response.Write "var cnt=0;" & Chr(13) & Chr(10) 
  Response.Write "var sName=document.form." & FName & ".value;" & Chr(13) & Chr(10)
  Response.Write "for(var i=0;i<sName.length;i++ ){" & Chr(13) & Chr(10)
  Response.Write "  if(escape(sName.charAt(i)).length >= 4) cnt+=2;" & Chr(13) & Chr(10)   
  Response.Write "  else cnt++;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "if(cnt>" & CLength & "){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位長度超過限制！');" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End function

Function CheckDateJ (FName,ListField)
  Response.Write "if(document.form." & FName & ".value.indexOf(""/"")==-1||document.form." & FName & ".value.indexOf(""/"")==1||document.form." & FName & ".value.indexOf(""/"")==document.form." & FName & ".value.length){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位格式錯誤！(西元年/月/日)');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "Ary_Date=document.form." & FName & ".value.split(""/"");" & Chr(13) & Chr(10)
  Response.Write "if(Ary_Date.length!=3) return false;" & Chr(13) & Chr(10)
  Response.Write "for(i=0;i<3;i++){" & Chr(13) & Chr(10)
  Response.Write "  if(isNaN(Number(Ary_Date[i]))==true){" & Chr(13) & Chr(10)
  Response.Write "    alert('"& ListField & "  欄位格式錯誤！(西元年/月/日)');" & Chr(13) & Chr(10)
  Response.Write "    document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "    return false;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "if(Ary_Date[0].length!=4){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位格式錯誤！(西元年/月/日)');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return false;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "if(parseInt(Number(Ary_Date[0]))<1000){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位格式錯誤！(西元年/月/日)');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return false;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "if(parseInt(Number(Ary_Date[1]))<1||parseInt(Number(Ary_Date[1]))>12){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位格式錯誤！(西元年/月/日)');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return false;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "var YYYY=parseInt(Number(Ary_Date[0]))" & Chr(13) & Chr(10)
  Response.Write "var MM=parseInt(Number(Ary_Date[1]))" & Chr(13) & Chr(10)
  Response.Write "var DD=parseInt(Number(Ary_Date[2]))" & Chr(13) & Chr(10)
  Response.Write "if(MM==1||MM==3||MM==5||MM==7||MM==8||MM==10||MM==12){" & Chr(13) & Chr(10)
  Response.Write "  if(DD<1||DD>31) return false;" & Chr(13) & Chr(10)
  Response.Write "}else if(MM==4||MM==6||MM==9||MM==11){" & Chr(13) & Chr(10)
  Response.Write "  if(DD<1||DD>30){" & Chr(13) & Chr(10)
  Response.Write "    alert('"& ListField & "  欄位格式錯誤！(西元年/月/日)');" & Chr(13) & Chr(10)
  Response.Write "    document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "    return false;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "}else if(MM==2){" & Chr(13) & Chr(10)
  Response.Write "  var leapyear=false;" & Chr(13) & Chr(10)
  Response.Write "  if(parseInt(YYYY)%4==0){" & Chr(13) & Chr(10)
  Response.Write "    if(parseInt(YYYY)%100==0){" & Chr(13) & Chr(10)
  Response.Write "      if(parseInt(YYYY)%400==0) leapyear=true;" & Chr(13) & Chr(10)
  Response.Write "    }else{" & Chr(13) & Chr(10)
  Response.Write "      leapyear=true;" & Chr(13) & Chr(10)
  Response.Write "    }" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "  if(leapyear){" & Chr(13) & Chr(10)
  Response.Write "    if(DD<1||DD>29){" & Chr(13) & Chr(10)
  Response.Write "      alert('"& ListField & "  欄位格式錯誤！(西元年/月/日)');" & Chr(13) & Chr(10)
  Response.Write "      document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "      return false;" & Chr(13) & Chr(10)
  Response.Write "    }" & Chr(13) & Chr(10)
  Response.Write "  }else{" & Chr(13) & Chr(10)
  Response.Write "    if(DD<1||DD>28){" & Chr(13) & Chr(10)
  Response.Write "      alert('"& ListField & "  欄位格式錯誤！(西元年/月/日)');" & Chr(13) & Chr(10)
  Response.Write "      document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "      return false;" & Chr(13) & Chr(10)
  Response.Write "    }" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "}else{" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位格式錯誤！(西元年/月/日)');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return false;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function


Function CheckNumberJ (FName,ListField)
  Response.Write "if(isNaN(Number(document.form." & FName & ".value))==true){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位必須為數字！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function CheckIntJ (FName,ListField)
  Response.Write "if(document.form." & FName & ".value.indexOf(""-"")!=-1&&document.form." & FName & ".value.indexOf(""."")!=-1){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位必須為大於0的整數！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function CheckMinNumberJ (FName,ListField,MinNum)  
  Response.Write "if(Number(document.form." & FName & ".value)<" & MinNum & "){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  必須大於" & MinNum & "元！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function CheckMaxNumberJ (FName,ListField,MaxNum)  
  Response.Write "if(Number(document.form." & FName & ".value)>" & MaxNum & "){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  必須小於" & MaxNum & "元！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function CheckEmailJ (FName,ListField)
  Response.Write "if(document.form." & FName & ".value.indexOf(""@"")==-1||document.form." & FName & ".value.indexOf(""."")==-1){" & Chr(13) & Chr(10)
  Response.Write "  alert('您輸入的"& ListField & "格式錯誤！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "if(!(/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(document.form." & FName & ".value))){" & Chr(13) & Chr(10)
  Response.Write "  alert('您輸入的"& ListField & "格式錯誤！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "var keyword=document.form." & FName & ".value.toLowerCase();" & Chr(13) & Chr(10)
  Response.Write "var AryKey = new Array('`','~','!','#','$','%','^','&','*','(',')','+','=','[',']','{','}','\\','|',';',':','\'','""','<','>',',','?','/');" & Chr(13) & Chr(10)
  Response.Write "for(var i=0;i<=AryKey.length-1;i++){" & Chr(13) & Chr(10)
  Response.Write "  if(keyword.indexOf(AryKey[i])!=-1){" & Chr(13) & Chr(10)
  Response.Write "    alert('您輸入的"& ListField & "請勿使用特殊字元！');" & Chr(13) & Chr(10)
  Response.Write "    document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "    return;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)    
  Response.Write "var keyword=document.form." & FName & ".value.toLowerCase();" & Chr(13) & Chr(10)
  'disabled by Samuel on 2014/1/22 will resume on 3.0 .net version
  'Response.Write "var AryKey = new Array('1=1','--','./','.asp','.css','.ini','/*','*/','.@','@.','@@','@C','@T','@P','acunetix','alert','and','begin','cast','char','chr','count','create','css','cursor','deallocate','declare','delete','dir','drop','end','exec','execute','fetch','from','iframe','insert','into','js','kill','left','master','mid','nchar','ntext','nvarchar','open','right','script','set','select','src','sys','sysobjects','syscolumns','table','text','truncate','update','varchar','while','xtype','%2527','address.tst');" & Chr(13) & Chr(10)
  'Response.Write "for(var i=0;i<=AryKey.length-1;i++){" & Chr(13) & Chr(10)
  'Response.Write "  if(keyword.indexOf(AryKey[i])!=-1){" & Chr(13) & Chr(10)
  'Response.Write "    alert('您輸入的"& ListField & "中請勿使用保留字元！');" & Chr(13) & Chr(10)
  'Response.Write "    document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  'Response.Write "    return;" & Chr(13) & Chr(10)
  'Response.Write "  }" & Chr(13) & Chr(10)
  'Response.Write "}" & Chr(13) & Chr(10)
End Function

Function CheckStringCheckBoxJ (FName,ListField)
  Response.Write "var check_checbox=false;" & Chr(13) & Chr(10)
  Response.Write "for(i=0;i<document.form." & FName & ".length;i++){" & Chr(13) & Chr(10)
  Response.Write "  if(document.form." & FName & "[i].checked){" & Chr(13) & Chr(10)
  Response.Write "    check_checbox=true;" & Chr(13) & Chr(10)
  Response.Write "    i=document.form." & FName & ".length;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "if(check_checbox==false){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位不可為空白！');" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function ChecklenCheckBoxJ(FName,CLength,ListField)'檢測CheckBox中英文夾雜字串實際長度
  Response.Write "var cnt=0;" & Chr(13) & Chr(10)
  Response.Write "var k=0;" & Chr(13) & Chr(10)
  Response.Write "for(i=0;i<document.form." & FName & ".length;i++){" & Chr(13) & Chr(10)
  Response.Write "  if(document.form." & FName & "[i].checked){" & Chr(13) & Chr(10)
  Response.Write "    for(var j=0;j<document.form." & FName & "[i].value.length;j++ ){" & Chr(13) & Chr(10)
  Response.Write "      if(escape(document.form." & FName & "[i].value.charAt(j)).length >= 4) cnt+=2;" & Chr(13) & Chr(10)
  Response.Write "      else cnt++;" & Chr(13) & Chr(10)
  Response.Write "    }" & Chr(13) & Chr(10)
  Response.Write "    k++;" & Chr(13) & Chr(10)
  Response.Write "    if(k>1) cnt++;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "if(cnt>" & CLength & "){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位長度超過限制！');" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End function

Function CheckKeyWordJ (FName,ListField)
  Response.Write "var content=document.form." & FName & ".value.toLowerCase();" & Chr(13) & Chr(10)
  Response.Write "var AryKey = new Array('script','iframe','a href','url','drop','create','delete','table','""','1=1','--',';');" & Chr(13) & Chr(10)  
  Response.Write "for(var i=0;i<=AryKey.length-1;i++){" & Chr(13) & Chr(10)
  Response.Write "  if(content.indexOf(AryKey[i])!=-1){" & Chr(13) & Chr(10)
  Response.Write "    alert('"& ListField & "  欄位請勿輸入 '+AryKey[i]+'  保留字元！');" & Chr(13) & Chr(10)
  Response.Write "    return;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function CheckExtJ (FName,Ext,ListField)
  Response.Write "if(document.form." & FName & ".value!==''){" & Chr(13) & Chr(10)
  Response.Write "  if(document.form." & FName & ".value.length>4){" & Chr(13) & Chr(10)
  Response.Write "    if(document.form." & FName & ".value.indexOf('.')>-1){" & Chr(13) & Chr(10)
  Response.Write "      var ext_check=false;" & Chr(13) & Chr(10) 
  Response.Write "      var filename=document.form." & FName & ".value.toLowerCase();" & Chr(13) & Chr(10)
  Response.Write "      var ext=filename.substr(filename.lastIndexOf('.')+1,filename.length);" & Chr(13) & Chr(10)
  If Ext="img" Then
    Response.Write "      var StrExt = 'jpg,gif,bmp';" & Chr(13) & Chr(10)
  ElseIf Ext="flash" Then
    Response.Write "      var StrExt = 'swf';" & Chr(13) & Chr(10)  
  ElseIf Ext="doc" Then
    Response.Write "      var StrExt = 'doc,pdf,xls';" & Chr(13) & Chr(10)
  ElseIf Ext="wmv" Then
    Response.Write "      var StrExt = 'wmv,asf,avi,mpg';" & Chr(13) & Chr(10)
  ElseIf Ext="flv" Then
    Response.Write "      var StrExt = 'flv';" & Chr(13) & Chr(10)    
  End If
  Response.Write "      var AryExt = StrExt.split(',');" & Chr(13) & Chr(10)
  Response.Write "      var AllExt = '';" & Chr(13) & Chr(10) 
  Response.Write "      for(var i=0;i<=AryExt.length-1;i++){" & Chr(13) & Chr(10)
  Response.Write "        if(AllExt==''){" & Chr(13) & Chr(10)
  Response.Write "          AllExt='.'+AryExt[i]+' ';" & Chr(13) & Chr(10)
  Response.Write "        }else{" & Chr(13) & Chr(10)
  Response.Write "          AllExt=AllExt+' 或 .'+AryExt[i]+' ';" & Chr(13) & Chr(10)
  Response.Write "        }" & Chr(13) & Chr(10)
  Response.Write "        if(AryExt[i]==ext){" & Chr(13) & Chr(10)
  Response.Write "          ext_check=true;" & Chr(13) & Chr(10)
  Response.Write "        }" & Chr(13) & Chr(10)        
  Response.Write "      }" & Chr(13) & Chr(10)
  Response.Write "      if(ext_check==false){" & Chr(13) & Chr(10)
  Response.Write "        alert('"& ListField & "  檔案名稱錯誤，副檔名必須為\n'+AllExt);" & Chr(13) & Chr(10)
  Response.Write "        return;" & Chr(13) & Chr(10)
  Response.Write "      }" & Chr(13) & Chr(10)
  Response.Write "    }else{" & Chr(13) & Chr(10)
  Response.Write "      alert('"& ListField & "  檔案名稱錯誤！');" & Chr(13) & Chr(10)
  Response.Write "      return;" & Chr(13) & Chr(10)
  Response.Write "    } " & Chr(13) & Chr(10)
  Response.Write "  }else{" & Chr(13) & Chr(10)
  Response.Write "    alert('"& ListField & "  檔案名稱錯誤！');" & Chr(13) & Chr(10)
  Response.Write "    return;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10) 
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function DiffDateJ (FName1,FName2)
  Response.Write "var begindate=new Date(document.form." & FName1 & ".value)" & Chr(13) & Chr(10)
  Response.Write "var enddate=new Date(document.form." & FName2 & ".value)" & Chr(13) & Chr(10)
  Response.Write "var diffdate=(Date.parse(begindate.toString())-Date.parse(enddate.toString()))/(1000*60*60*24)" & Chr(13) & Chr(10)
  Response.Write "if(parseInt(diffdate)>0){" & Chr(13) & Chr(10)
  Response.Write "  alert('刊登起日不可大於迄日！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName1 & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)  
  Response.Write "}" & Chr(13) & Chr(10)
End Function 

Function DiffDateMsgJ (FName1,FName2,Msg)
  Response.Write "var begindate=new Date(document.form." & FName1 & ".value)" & Chr(13) & Chr(10)
  Response.Write "var enddate=new Date(document.form." & FName2 & ".value)" & Chr(13) & Chr(10)
  Response.Write "var diffdate=(Date.parse(begindate.toString())-Date.parse(enddate.toString()))/(1000*60*60*24)" & Chr(13) & Chr(10)
  Response.Write "if(parseInt(diffdate)>0){" & Chr(13) & Chr(10)
  Response.Write "  alert('"&Msg&"起日不可大於迄日！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName1 & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)  
  Response.Write "}" & Chr(13) & Chr(10)
End Function 

Function CheckVerifyCodeJ (FName,CLength,ListField)
  Response.Write "if(document.form." & FName & ".value==''){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位不可為空白！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "if(isNaN(Number(document.form." & FName & ".value))==true){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位必須為數字！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "if(document.form." & FName & ".value.indexOf(""-"")!=-1&&document.form." & FName & ".value.indexOf(""."")!=-1){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位必須為大於0的整數！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
  Response.Write "if(document.form." & FName & ".value.length!=" & CLength & "){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  欄位必須為" & CLength & "個數字！');" & Chr(13) & Chr(10)
  Response.Write "  document.form." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)      
End Function

Function CheckImage_Width_HeightJ (FName,image_width,image_height)
  Response.Write "if(document.form." & FName & ".value!=''){" & Chr(13) & Chr(10)
  Response.Write "  var image = new Image();" & Chr(13) & Chr(10)
  Response.Write "  image.src = document.form."&FName&".value;" & Chr(13) & Chr(10)
  Response.Write "  var iwidth=image.width;" & Chr(13) & Chr(10)
  Response.Write "  var iheight=image.height;" & Chr(13) & Chr(10)
  Response.Write "  document.form." & image_width & ".value=iwidth;" & Chr(13) & Chr(10)
  Response.Write "  document.form." & image_height & ".value=iheight; " & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10) 
End Function

Function CheckIDJ(FName) '檢核身份證/統編
	Response.Write "val = document.form."&FName&".value;" & Chr(13) & Chr(10) 
	Response.Write "if (val=='') return true;" & Chr(13) & Chr(10)
	Response.Write "if(val.length != 8 && val.length != 10){" & Chr(13) & Chr(10) 
	Response.Write "	alert('身份證/統編長度為8碼或10碼！');" & Chr(13) & Chr(10) 
	Response.Write "	return false;" & Chr(13) & Chr(10) 
	Response.Write "}" & Chr(13) & Chr(10) 
	Response.Write "else if(val.length == 10){" & Chr(13) & Chr(10)
	Response.Write "	tab = 'ABCDEFGHJKLMNPQRSTUVXYWZIO'" & Chr(13) & Chr(10) 
	Response.Write "	A1 = new Array (1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,3,3,3,3,3,3 );" & Chr(13) & Chr(10) 
	Response.Write "	A2 = new Array (0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5 );" & Chr(13) & Chr(10) 
	Response.Write "	Mx = new Array (9,8,7,6,5,4,3,2,1,1);" & Chr(13) & Chr(10) 
	Response.Write "	i = tab.indexOf( val.charAt(0) );" & Chr(13) & Chr(10) 	
	Response.Write "	if ( i == -1 ) {alert('身份證字號輸入錯誤！'); return false;}" & Chr(13) & Chr(10) 
	Response.Write " 	sum = A1[i] + A2[i]*9;" & Chr(13) & Chr(10) 	
	Response.Write "	for ( i=1; i<10; i++ ) {" & Chr(13) & Chr(10) 
	Response.Write "		v = parseInt( val.charAt(i) );" & Chr(13) & Chr(10) 
	Response.Write "		if ( isNaN(v) ) {alert('身份證字號輸入錯誤！'); return false;}" & Chr(13) & Chr(10) 
	Response.Write "		sum = sum + v * Mx[i];" & Chr(13) & Chr(10) 
	Response.Write "	}" & Chr(13) & Chr(10) 
	Response.Write "	if ( sum % 10 != 0 ) {alert('身份證字號輸入錯誤！'); return false;}" & Chr(13) & Chr(10)
	Response.Write "	else return true;" & Chr(13) & Chr(10)
	Response.Write "}" & Chr(13) & Chr(10) 		
	Response.Write "else{" & Chr(13) & Chr(10) 	
	Response.Write "	for (i = 0; i < 8; i++){" & Chr(13) & Chr(10) 		
	Response.Write "		c = val.charAt(i);" & Chr(13) & Chr(10) 		
	Response.Write "		if ('0123456789'.indexOf(c) == -1) {alert('統編輸入錯誤！'); return false;}" & Chr(13) & Chr(10) 			 		
	Response.Write "	}" & Chr(13) & Chr(10) 		
	Response.Write "	d1 = parseInt(val.charAt(0));" & Chr(13) & Chr(10)
	Response.Write "	d2 = parseInt(val.charAt(1));" & Chr(13) & Chr(10) 		
	Response.Write "	d3 = parseInt(val.charAt(2));" & Chr(13) & Chr(10) 		
	Response.Write "	d4 = parseInt(val.charAt(3));" & Chr(13) & Chr(10) 		
	Response.Write "	d5 = parseInt(val.charAt(4));" & Chr(13) & Chr(10) 		
	Response.Write "	d6 = parseInt(val.charAt(5));" & Chr(13) & Chr(10) 		
	Response.Write "	d7 = parseInt(val.charAt(6));" & Chr(13) & Chr(10) 		
	Response.Write "	cd8 = parseInt(val.charAt(7));" & Chr(13) & Chr(10) 		
	Response.Write "	c1 = d1;" & Chr(13) & Chr(10) 		
	Response.Write "	c2 = d3;" & Chr(13) & Chr(10) 	
	Response.Write "	c3 = d5;" & Chr(13) & Chr(10) 	
	Response.Write "	c4 = cd8;" & Chr(13) & Chr(10) 	
	Response.Write "	a1 = parseInt((d2 * 2) / 10);" & Chr(13) & Chr(10) 	
	Response.Write "	b1 = (d2 * 2) % 10;" & Chr(13) & Chr(10) 	
	Response.Write "	a2 = parseInt((d4 * 2) / 10);" & Chr(13) & Chr(10) 	
	Response.Write "	b2 = (d4 * 2) % 10;" & Chr(13) & Chr(10) 	
	Response.Write "	a3 = parseInt((d6 * 2) / 10);" & Chr(13) & Chr(10) 	
	Response.Write "	b3 = (d6 * 2) % 10;" & Chr(13) & Chr(10) 	
	Response.Write "	a4 = parseInt((d7 * 4) / 10);" & Chr(13) & Chr(10) 	
	Response.Write "	b4 = (d7 * 4) % 10;" & Chr(13) & Chr(10) 	
	Response.Write "	a5 = parseInt((a4 + b4) / 10);" & Chr(13) & Chr(10) 	
	Response.Write "	b5 = (a4 + b4) % 10;" & Chr(13) & Chr(10) 	
	Response.Write "	if((a1 + b1 + c1 + a2 + b2 + c2 + a3 + b3 + c3 + a4 + b4 + c4) % 10 == 0) return true;" & Chr(13) & Chr(10) 	
	Response.Write "	if(d7 = 7){" & Chr(13) & Chr(10) 
	Response.Write "		if((a1 + b1 + c1 + a2 + b2 + c2 + a3 + b3 + c3 + a5 + c4) % 10 == 0) return true;" & Chr(13) & Chr(10) 		
	Response.Write "	}" & Chr(13) & Chr(10) 
	Response.Write "	alert('統編輸入錯誤！');" & Chr(13) & Chr(10) 
	Response.Write "	return false;" & Chr(13) & Chr(10) 
	Response.Write "}" & Chr(13) & Chr(10) 			
End Function

'檢核信用卡卡別(VISA,MasterCard,JCB,AE)/卡號
Function CheckCreditCardNumJ(CardType,AccountFName1,AccountFName2,AccountFName3,AccountFName4) 	
	Response.Write "var CardType = document.form."&CardType&".value;" & Chr(13) & Chr(10) 	
	Response.Write "if(CardType=='銀聯' || CardType=='聯合信用卡')" & Chr(13) & Chr(10) 	
	Response.Write "	return true;" & Chr(13) & Chr(10) 	
	Response.Write "var num = document.form."&AccountFName1&".value;" & Chr(13) & Chr(10) 
	Response.Write "		num += document.form."&AccountFName2&".value;" & Chr(13) & Chr(10) 
	Response.Write "		num += document.form."&AccountFName3&".value;" & Chr(13) & Chr(10) 
	Response.Write "		num += document.form."&AccountFName4&".value;" & Chr(13) & Chr(10) 
	Response.Write "var len = num.length;" & Chr(13) & Chr(10) 
	Response.Write "var sum=0;" & Chr(13) & Chr(10)
	Response.Write "var mul=1;" & Chr(13) & Chr(10)
	Response.Write "for(i=(len-1); i >= 0 ;--i){" & Chr(13) & Chr(10)	
	Response.Write "	numAtI = parseInt(num.charAt(i));" & Chr(13) & Chr(10)	
	Response.Write "	tempSum = numAtI * mul;" & Chr(13) & Chr(10)	
	Response.Write "	if(tempSum > 9){" & Chr(13) & Chr(10)	
	Response.Write "		first = tempSum % 10;" & Chr(13) & Chr(10)	
	Response.Write "		tempSum =  1 +  first;" & Chr(13) & Chr(10)	
	Response.Write "	}" & Chr(13) & Chr(10)	
	Response.Write "	sum += tempSum;" & Chr(13) & Chr(10)	
	Response.Write "	if (mul == 1) mul++;" & Chr(13) & Chr(10)	
	Response.Write "	else mul--;" & Chr(13) & Chr(10)	
	Response.Write "}" & Chr(13) & Chr(10)		
	Response.Write "if(sum % 10){" & Chr(13) & Chr(10)	
	Response.Write "	if (len == 15){" & Chr(13) & Chr(10)	
	Response.Write "		txtcard = CheckFor15(num,0);" & Chr(13) & Chr(10)	
	Response.Write "		if ( txtcard == 'invalid' ){" & Chr(13) & Chr(10)	
	Response.Write "			alert('此卡號無效!!');" & Chr(13) & Chr(10)	
	Response.Write "			return false;" & Chr(13) & Chr(10)	
	Response.Write "		}" & Chr(13) & Chr(10)	
	Response.Write "	}" & Chr(13) & Chr(10)	
	Response.Write "	else{" & Chr(13) & Chr(10)	
	Response.Write "		alert('此卡號無效!!');" & Chr(13) & Chr(10)	
	Response.Write "		return false;" & Chr(13) & Chr(10)	
	Response.Write "	}" & Chr(13) & Chr(10)	
	Response.Write "}" & Chr(13) & Chr(10)			
	Response.Write "	switch(len){" & Chr(13) & Chr(10)	
	Response.Write "	case 13:" & Chr(13) & Chr(10)	
	Response.Write "		txtcard = CheckFor13(num);" & Chr(13) & Chr(10)	
	Response.Write "		break;" & Chr(13) & Chr(10)	
	Response.Write "	case 14:" & Chr(13) & Chr(10)	
	Response.Write "		txtcard = CheckFor14(num);" & Chr(13) & Chr(10)	
	Response.Write "		break;" & Chr(13) & Chr(10)	
	Response.Write "		case 15:" & Chr(13) & Chr(10)	
	Response.Write "			txtcard = CheckFor15(num,1);" & Chr(13) & Chr(10)	
	Response.Write "			break;" & Chr(13) & Chr(10)	
	Response.Write "	case 16:" & Chr(13) & Chr(10)	
	Response.Write "		txtcard = CheckFor16(num);" & Chr(13) & Chr(10)	
	Response.Write "		break;" & Chr(13) & Chr(10)	
	Response.Write "		default:" & Chr(13) & Chr(10)	
	Response.Write "			txtcard = 'invalid';" & Chr(13) & Chr(10)	
	Response.Write "	}" & Chr(13) & Chr(10)		
	Response.Write "	if(txtcard == 'invalid'){" & Chr(13) & Chr(10)	
	Response.Write "		alert('此卡號無效!!');" & Chr(13) & Chr(10)	
	Response.Write "		return false;" & Chr(13) & Chr(10)	
	Response.Write "	}" & Chr(13) & Chr(10)	
	Response.Write "	else{" & Chr(13) & Chr(10)	
	Response.Write "		if(CardType==txtcard){" & Chr(13) & Chr(10)	
'	Response.Write "			alert('此卡別及卡號正確!!');" & Chr(13) & Chr(10)
	Response.Write "			return true;}" & Chr(13) & Chr(10)
	Response.Write "		else{" & Chr(13) & Chr(10)	
	Response.Write "			alert('此卡號正確!!但卡片應該是：'+txtcard);" & Chr(13) & Chr(10)
	Response.Write "			return false;" & Chr(13) & Chr(10)	
	Response.Write "		}" & Chr(13) & Chr(10)
	Response.Write "	}" & Chr(13) & Chr(10)		
	
	'Checks for Visa
	Response.Write "function CheckFor13(number){" & Chr(13) & Chr(10)		
	Response.Write "	if(parseInt(number.charAt(0)) == 4 )" & Chr(13) & Chr(10)			
	Response.Write "		return 'VISA';" & Chr(13) & Chr(10)			
	Response.Write "	return 'invalid';" & Chr(13) & Chr(10)			
	Response.Write "}" & Chr(13) & Chr(10)
	
	'Check for Diners Club
	Response.Write "function CheckFor14(number){" & Chr(13) & Chr(10)	
	Response.Write "	if( parseInt(number.charAt(0)) != 3 )" & Chr(13) & Chr(10)			
	Response.Write "		return 'invalid';" & Chr(13) & Chr(10)			
	Response.Write "	if( (parseInt(number.charAt(1)) == 6) || (parseInt(number.charAt(1)) == 8) )" & Chr(13) & Chr(10)			
	Response.Write "		return 'Diners Club/Carte Blanche';" & Chr(13) & Chr(10)		
	Response.Write "	if( !(parseInt(number.charAt(1)) ))" & Chr(13) & Chr(10)			
	Response.Write "		if( (parseInt(number.charAt(2)) >= 0) && (parseInt(number.charAt(2)) <= 5))" & Chr(13) & Chr(10)		
	Response.Write "			return 'Diners Club/Carte Blanche';" & Chr(13) & Chr(10)	
	Response.Write "	return 'invalid';" & Chr(13) & Chr(10)			
	Response.Write "}" & Chr(13) & Chr(10)
	
	'Check for American Express, enRoute, JCB
	Response.Write "function CheckFor15(number, chec){" & Chr(13) & Chr(10)	
	Response.Write "	if ( (number.charAt(0) == 3) && ( (number.charAt(1) == 4)||(number.charAt(1) == 7) ) && chec)" & Chr(13) & Chr(10)	
	Response.Write "		return 'AE';" & Chr(13) & Chr(10)
	Response.Write "	var FirstFour =  parseInt(number.charAt(0))*1000;" & Chr(13) & Chr(10)
	Response.Write "			FirstFour += parseInt(number.charAt(1))*100;" & Chr(13) & Chr(10)
	Response.Write "			FirstFour += parseInt(number.charAt(2))*10;" & Chr(13) & Chr(10)
	Response.Write "			FirstFour += parseInt(number.charAt(3));" & Chr(13) & Chr(10)
	Response.Write "	if( (FirstFour == 2014) || (FirstFour == 2149) )" & Chr(13) & Chr(10)
	Response.Write "		return 'enRoute';" & Chr(13) & Chr(10)
	Response.Write "	if( ((FirstFour == 2131) || (FirstFour == 1800)) && chec )" & Chr(13) & Chr(10)
	Response.Write "		return 'JCB';" & Chr(13) & Chr(10)	
	Response.Write "return 'invalid';" & Chr(13) & Chr(10)
	Response.Write "}" & Chr(13) & Chr(10)

	'Check for Visa, MasterCard, Discover or JCB
	Response.Write "function CheckFor16(number){" & Chr(13) & Chr(10)		
	Response.Write "	var a = parseInt(number.charAt(0));" & Chr(13) & Chr(10)	
	Response.Write "	var b = parseInt(number.charAt(1));" & Chr(13) & Chr(10)	
	Response.Write "	switch (a){" & Chr(13) & Chr(10)	
	Response.Write "		case 5:" & Chr(13) & Chr(10)	
	Response.Write "			if((b > 0)&& (b < 6))" & Chr(13) & Chr(10)	
	Response.Write "				return 'MASTER';" & Chr(13) & Chr(10)	
	Response.Write "			else" & Chr(13) & Chr(10)	
	Response.Write "				return 'invalid';" & Chr(13) & Chr(10)	
	Response.Write "			break;" & Chr(13) & Chr(10)	
	Response.Write "		case 4:" & Chr(13) & Chr(10)	
	Response.Write "			return 'VISA';" & Chr(13) & Chr(10)	
	Response.Write "			break;" & Chr(13) & Chr(10)	
	Response.Write "		case 6:" & Chr(13) & Chr(10)	
	Response.Write "			if ( (b == 0) && ( parseInt(number.charAt(2)) == 1) && ( parseInt(number.charAt(3)) == 1))" & Chr(13) & Chr(10)	
	Response.Write "				return 'Discover';" & Chr(13) & Chr(10)	
	Response.Write "			else" & Chr(13) & Chr(10)	
	Response.Write "				return 'invalid';" & Chr(13) & Chr(10)	
	Response.Write "			break;" & Chr(13) & Chr(10)	
	Response.Write "		case 3:" & Chr(13) & Chr(10)	
	Response.Write "			return 'JCB';" & Chr(13) & Chr(10)	
	Response.Write "			break;" & Chr(13) & Chr(10)
	Response.Write "		default:" & Chr(13) & Chr(10)
	Response.Write "			return 'invalid';" & Chr(13) & Chr(10)
	Response.Write "			break;" & Chr(13) & Chr(10)
	Response.Write "	}" & Chr(13) & Chr(10)
	Response.Write "}" & Chr(13) & Chr(10)
End Function

Function Calendar (FName,Val)
  Response.Write "<input type=text name="&FName&" size=10 class=font9 value="&Val&">" & Chr(13) & Chr(10)
  Response.Write "<a href onclick=cal19.select(document.form."&FName&",'"&FName&"','yyyy/MM/dd');>" & Chr(13) & Chr(10)
  Response.Write "<img border=0 src=../images/date.gif width=16 height=14></a>" & Chr(13) & Chr(10)
End Function

'20131112 Add by GoodTV Tanya:增加onblur的Calendar同步資料
Function Calendar2 (FName,Val,OnblurFName)
  Response.Write "<input type=text name="&FName&" size=10 class=font9 value="&Val&" onblur='chg_Date(this.value)'>" & Chr(13) & Chr(10)
  Response.Write "<a href onclick=cal19.select(document.form."&FName&",'"&FName&"','yyyy/MM/dd');>" & Chr(13) & Chr(10)
  Response.Write "<img border=0 src=../images/date.gif width=16 height=14></a>" & Chr(13) & Chr(10)
  
  Response.Write "<script language=""JavaScript""><!--" & Chr(13) & Chr(10)
  Response.Write "  function chg_Date(DonateFromDate){" & Chr(13) & Chr(10)
  Response.Write "    document.form."&OnblurFName&".value = DonateFromDate;" & Chr(13) & Chr(10) 
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "--></script>" & Chr(13) & Chr(10)
End Function

'20131113 Add by GoodTV Tanya:增加onblur的Calendar驗證日期區間是否重覆
Function Calendar3 (FName,Val,OnblurFName,L_Donate_FromDate,L_Donate_ToDate)
  Response.Write "<input type=text name="&FName&" size=10 class=font9 onblur='chg_Date2()' value="&Val&">" & Chr(13) & Chr(10)
  Response.Write "<a href onclick=cal19.select(document.form."&FName&",'"&FName&"','yyyy/MM/dd');>" & Chr(13) & Chr(10)
  Response.Write "<img border=0 src=../images/date.gif width=16 height=14></a>" & Chr(13) & Chr(10)
  
  Response.Write "<script language=""JavaScript""><!--" & Chr(13) & Chr(10)
  Response.Write "  function chg_Date2(){" & Chr(13) & Chr(10)  
  Response.Write "    var Donate_FromDate = Date.parse(document.form."&OnblurFName&".value);" & Chr(13) & Chr(10)   
  Response.Write "    var L_Donate_FromDate = Date.parse('"&L_Donate_FromDate&"');" & Chr(13) & Chr(10)  
  Response.Write "    var L_DonateToDate = Date.parse('"&L_Donate_ToDate&"');" & Chr(13) & Chr(10) 
  Response.Write "    if (L_Donate_FromDate<=Donate_FromDate&&L_DonateToDate>=Donate_FromDate)" & Chr(13) & Chr(10)
  Response.Write "    	alert('授權期間重覆!');" & Chr(13) & Chr(10)   
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "--></script>" & Chr(13) & Chr(10)
End Function

Function SubmitJ (SubmitType)
  if SubmitType="delete" then Response.Write "if(confirm('您是否確定要刪除？')){" & Chr(13) & Chr(10)
  if SubmitType="update" then Response.Write "if(confirm('您是否確定要修改？')){" & Chr(13) & Chr(10)
  if SubmitType="export" then Response.Write "if(confirm('您是否確定要將查詢結果匯出？')){" & Chr(13) & Chr(10)
  Response.Write "document.form.action.value='"&SubmitType&"';" & Chr(13) & Chr(10)
  Response.Write "document.form.submit();" & Chr(13) & Chr(10)
  if SubmitType="delete" or SubmitType="update" or SubmitType="export" then Response.Write "}" & Chr(13) & Chr(10)
End Function

Function CheckRedirectJ (URLName)
  Response.Write "window.location.href='" & URLName & "';"
End Function

Function Sel_CodeType(SQL,File_TypeName,File_TypeData,File_SubName)
  Row1 = 0
  Sel_TypeCode="<option value=''> </option>"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  While Not RS1.EOF
    If Cstr(RS1(File_TypeName))=Cstr(File_TypeData) Then
      Sel_TypeCode = Sel_TypeCode & "<option value='" & RS1(File_TypeName) & "' selected >" & RS1(File_TypeName) & "</option>"
    Else   
      Sel_TypeCode = Sel_TypeCode & "<option value='" & RS1(File_TypeName) & "'>" & RS1(File_TypeName) & "</option>"
    End If
    If Row1 = 0 Then
      Sub_Code = Sub_Code & "'" & RS1(File_TypeName) & "'" & ","
    Else 
      Sub_Code = Sub_Code & "," & "'" & RS1(File_TypeName) & "'" & ","
    End If    
    Row2=1
    Sub_Code = Sub_Code & "["
    SQL2="Select CodeDesc From CaseCode Where CodeName='"&RS1(File_TypeName)&"' And Menu_Url like 'content.asp%' Order By Seq"
    Set RS2 = Server.CreateObject("ADODB.RecordSet")
    RS2.Open SQL2,Conn,1,1
    Do While Not RS2.EOF    
      If Row2 = 1 Then
        Sub_Code = Sub_Code & "'" & RS2("CodeDesc") & "'"
      Else 
        Sub_Code = Sub_Code & "," & "'" & RS2("CodeDesc") & "'" 
      End If
      Row2=Row2+1      
      RS2.MoveNext
    Loop
    RS2.Close
    Set RS2 = Nothing
    Sub_Code=Sub_Code&"]"    
    Row1=Row1+1
    RS1.MoveNext
  Wend
  RS1.Close
  Set RS1=Nothing
  Response.Write "<Select name='"&File_TypeName&"' Size='1' style='font-size: 9pt; font-family: 新細明體' onchange='Chg_Type(this.value)'>"&Sel_TypeCode&"</Select>"  
  Response.Write "<script language=""JavaScript""><!--" & Chr(13) & Chr(10)
  Response.Write "  function Chg_Type(TypeCode){" & Chr(13) & Chr(10)
  Response.Write "    ClearOption(document.form."&File_SubName&");" & Chr(13) & Chr(10)
  Response.Write "    document.form."&File_SubName&".options[0] = new Option('','');" & Chr(13) & Chr(10)
  Response.Write "    var ArySubCode = new Array("&Sub_Code&")" & Chr(13) & Chr(10)
  Response.Write "    for(var i=0;i<=ArySubCode.length-1;i=i+2){" & Chr(13) & Chr(10)
  Response.Write "      if(ArySubCode[i]==TypeCode){" & Chr(13) & Chr(10)
  Response.Write "        var theSub = ArySubCode[i+1];" & Chr(13) & Chr(10)
  Response.Write "        var k=1;" & Chr(13) & Chr(10)
  Response.Write "        for(var j=0;j<=theSub.length-1;j++){" & Chr(13) & Chr(10)
  Response.Write "          document.form."&File_SubName&".options[k] = new Option(theSub[j],theSub[j]);" & Chr(13) & Chr(10)
  Response.Write "          k++;" & Chr(13) & Chr(10)
  Response.Write "        }" & Chr(13) & Chr(10)
  Response.Write "      }" & Chr(13) & Chr(10)
  Response.Write "    }" & Chr(13) & Chr(10)
  Response.Write "    Chg_SubType();" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10) 
  Response.Write "  function ClearOption(SelObj){" & Chr(13) & Chr(10)
  Response.Write "    var i;" & Chr(13) & Chr(10)
  Response.Write "    for(i=SelObj.length-1;i>=0;i--){" & Chr(13) & Chr(10)
  Response.Write "      SelObj.options[i] = null;" & Chr(13) & Chr(10)
  Response.Write "    }" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "--></script>" & Chr(13) & Chr(10)
End Function

Function Sel_SubType(File_TypeDate,File_SubName,File_SubDate)
  Sel_SubCode="<option value=''> </option>"
  If File_TypeDate<>"" Then
    SQL="Select CodeDesc From CaseCode Where CodeName='"&File_TypeDate&"' Order By Seq"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL,Conn,1,1
    Do While Not RS1.EOF
      If Cstr(RS1("CodeDesc"))=Cstr(File_SubDate) Then
        Sel_SubCode = Sel_SubCode & "<option value='" & RS1("CodeDesc") & "' selected >" & RS1("CodeDesc") & "</option>"
      Else
        Sel_SubCode = Sel_SubCode & "<option value='" & RS1("CodeDesc") & "'>" & RS1("CodeDesc") & "</option>"
      End If
      RS1.MoveNext
    Loop
    RS1.Close
    Set RS1 = Nothing       
  End If
  Response.Write "<Select name='"&File_SubName&"' Size='1' style='font-size: 9pt; font-family: 新細明體' onchange='Chg_SubType()'>"&Sel_SubCode&"</Select>"  
End Function

Function OptionList (SQL,FName,Listfield,BoundColumn,menusize)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  Response.Write "<SELECT Name='" & FName & "' size='" & menusize & "' style='font-size: 9pt; font-family: 新細明體'>"
  'If RS1.EOF Then
  '  Response.Write "<OPTION>" & " " & "</OPTION>"
  'End If
  'If BoundColumn="" or IsNull(BoundColumn) Then
  '  Response.Write "<OPTION selected value=''>" & " " & "</OPTION>"
  'End If    
  Response.Write "<OPTION>" & " " & "</OPTION>"
  While Not RS1.EOF
    If BoundColumn<>"" Then
      If Cstr(RS1(FName))=Cstr(BoundColumn) Then
        strselected = "selected"
      Else
        strselected = ""
      End If
    Else
      strselected = ""
    End If
    Response.Write "<OPTION " & strselected & " value='" & RS1(FName) & "'>" & RS1(Listfield) & "</OPTION>"
    RS1.MoveNext
  Wend
  Response.Write "</SELECT>"
  RS1.Close
  Set RS1=Nothing
End Function

Function OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,BoundNull)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  Response.Write "<SELECT Name='" & FName & "' size='" & menusize & "' style='font-size: 9pt; font-family: 新細明體'>"
  If RS1.EOF Then
    Response.Write "<OPTION>" & " " & "</OPTION>"
  Else
    If BoundColumn="" Or IsNull(BoundColumn) Or (BoundNull="Y" And RS1.Recordcount>1) Then
      Response.Write "<OPTION selected value=''>" & " " & "</OPTION>"
    End If     
  End If   
  While Not RS1.EOF
    If RS1(FName)=BoundColumn Or Cint(RS1.Recordcount)=1 Then
      strselected = "selected"
    Else
      strselected = ""
    End If
    Response.Write "<OPTION " & strselected & " value='" & RS1(FName) & "'>" & RS1(Listfield) & "</OPTION>"
    RS1.MoveNext
  Wend
  Response.Write "</SELECT>"
  RS1.Close
  Set RS1=Nothing
End Function

Function RadioBoxList (SQL,FName,Listfield,BoundColumn,NoChecked)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  Dim i
  I = 1
  While Not RS1.EOF
    If BoundColumn="" Then
      If I = NoChecked Then
        StrChecked="checked"
      Else
        StrChecked=""   
      End If  
    Else
      If RS1(FName)=BoundColumn Then
        StrChecked="checked"
      Else
        StrChecked=""   
      End If
    End If   	
    Response.Write "<Input Type='radio' " & StrChecked & " Name='" & FName &"' ID='" & FName & I & "' value='" & RS1(FName) & "'>" & RS1(Listfield)
    I = I + 1
    RS1.MoveNext
  WEnd
  RS1.Close
  Set RS1=Nothing
End Function

Function CheckBoxList (SQL,FName,Listfield,BoundColumn)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  Dim i
  I = 1
  While Not RS1.EOF
    If Instr(BoundColumn,RS1(FName)) > 0 Then
      StrChecked="checked"
    Else
      StrChecked=""
    End If   
    Response.Write "<Input Type='checkbox' " & StrChecked & " Name='" & FName &"' ID='" & FName & I & "' value='" & RS1(FName) & "'>" & RS1(Listfield) & "&nbsp;"
    I = I + 1
    RS1.MoveNext
  WEnd
  RS1.Close
  Set RS1=Nothing
End Function

Sub SaveButton()
  Response.Write "<button type='button' id='save' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Save_OnClick()'> <img src='../images/save.gIf' width='19' height='20' align='absmiddle'> 存檔</button>&nbsp;"
  Response.Write "<button type='button' id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Cancel_OnClick()'> <img src='../images/icon6.gIf' width='20' height='15' align='absmiddle'> 離開</button>&nbsp;"
End Sub

Sub EditButton()
  Response.Write "<button type='button' id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Update_OnClick()'> <img src='../images/update.gIf' width='20' height='20' align='absmiddle'> 修改</button>&nbsp;"
  Response.Write "<button type='button' id='delete' style='position:relative;left:0;width:75;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Delete_OnClick()'> <img src='../images/delete.gIf' width='20' height='20' align='absmiddle'> 刪除</button>&nbsp;"
  Response.Write "<button type='button' id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Cancel_OnClick()'> <img src='../images/icon6.gIf' width='20' height='15' align='absmiddle'> 離開</button>&nbsp;"
End Sub

Function GridListHit (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,Hit_Nemu)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  If Not RS1.EOF Then 
    FieldsCount = RS1.Fields.Count-1
    totRec=RS1.Recordcount
    If totRec>0 Then 
      RS1.PageSize=PageSize
      If nowPage="" or nowPage=0 Then 
        nowPage=1
      ElseIf cint(nowPage) > RS1.PageCount Then 
        nowPage=RS1.PageCount 
      End If
      session("nowPage")=nowPage
      RS1.AbsolutePage=nowPage
      totPage=RS1.PageCount
      SQL=server.URLEncode(SQL)
    End If
    Response.Write "<Form action='' id=form1 name=form1>" 
    Response.Write "<table border=0 cellspacing='0' cellpadding='1' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='30%'></td>"
    Response.Write "<td width='55%'><span style='font-size: 9pt; font-family: 新細明體'>第" & nowPage & "/" & totPage & "頁&nbsp;&nbsp;</span>"
    Response.Write "<span style='font-size: 9pt; font-family: 新細明體'>共" & FormatNumber(totRec,0) & "筆&nbsp;&nbsp;</span>"
    If cint(nowPage) <>1 Then             
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage-1)&"&SQL="&SQL&"'>上一頁</a>" 
    End If
    If cint(nowPage)<>RS1.PageCount and cint(nowPage)<RS1.PageCount Then 
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage+1)&"&SQL="&SQL&"'>下一頁</a>" 
    End If
    Response.Write " |&nbsp;<span style='font-size: 9pt; font-family: 新細明體'> 跳至第<select name=GoPage size='1' style='font-size: 9pt; font-family: 新細明體' onchange='GoPage_OnChange(this.value)'>"
    For iPage=1 to totPage
      If iPage=cint(nowPage) Then
        strSelected = "selected"
      Else
	      strSelected = "" 
      End If
      Response.Write "<option value='"&iPage&"'" & strSelected & ">" & iPage & "</option>"          
    Next
    Response.Write "</select>頁</span></td>" 
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    End If
    Response.Write "</tr></table>"
    Dim I
    Dim J
    Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    For J = 1 To FieldsCount
      Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J).Name & "</span></font></td>"
    Next
    If Hit_Nemu<>"" Then 
      Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>累計點閱次數</span></font></td>"
      If Hit_Nemu="activity" Then Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>報名明細</span></font></td>"
    End If
    Response.Write "</tr>"
    I = 1
    While Not RS1.EOF And I <= RS1.PageSize
      If Hlink<>"" Then
        Response.Write "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>"
        showhand="style='cursor:hand'"
      End If
      Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
      For J = 1 To FieldsCount
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(J))&"</span></td>"
      Next
      If Hit_Nemu<>"" Then 
        Total=0
        SQL2="Select Total=Isnull(Sum(Hit_Count),0) From HIT Where Hit_Nemu='"&Hit_Nemu&"' And Hit_Object_ID='"&RS1(0)&"'"
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,1
        If Not RS2.EOF Then Total=RS2("Total")
        RS2.Close
        Set RS2=Nothing 
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(Total,0)&"</span></td>"
        If Hit_Nemu="activity" Then 
          Total=0
          SQL2="Select Total=Count(*) From ACTIVITY_SIGNUP Where Activity_Id='"&RS1(0)&"'"
          Set RS2 = Server.CreateObject("ADODB.RecordSet")
          RS2.Open SQL2,Conn,1,1
          If Not RS2.EOF Then Total=RS2("Total")
          RS2.Close
          Set RS2=Nothing 
          Response.Write "</a><td><span style='font-size: 9pt; font-family: 新細明體'><a href='activity_signup.asp?activity_id="&RS1(0)&"' target='main'>&nbsp;報名明細(&nbsp;"&FormatNumber(Total,0)&"&nbsp;)</a></span></td>"       
        End If
      End If
      I = I + 1
      RS1.MoveNext
      Response.Write "</tr>"
      If Hit_Nemu<>"activity" Then Response.Write "</a>"
    Wend
    Response.Write "</table>"
    Response.Write "</form>"
  Else
    Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 沒有符合條件的資料 **</td>"
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    End If  
    Response.Write "</table>"
  End If 
  RS1.Close
  Set RS1=Nothing
End Function

Function GridListHit_S (AddLink)
  Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
  Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 請輸入查詢條件 **</td>"
  If AddLink <> "" Then
    Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
  End If  
  Response.Write "</table>"
End Function

Function GridList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  If Not RS1.EOF Then 
    FieldsCount = RS1.Fields.Count-1
    totRec=RS1.Recordcount
    If totRec>0 Then 
      RS1.PageSize=PageSize
      If nowPage="" or nowPage=0 Then
        nowPage=1
      ElseIf cint(nowPage) > RS1.PageCount Then 
        nowPage=RS1.PageCount 
      End If
      session("nowPage")=nowPage
      RS1.AbsolutePage=nowPage
      totPage=RS1.PageCount
      SQL=server.URLEncode(SQL)
    End If
    Response.Write "<Form action='' id=form1 name=form1>" 
    Response.Write "<table border=0 cellspacing='0' cellpadding='1' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='30%'></td>"
    Response.Write "<td width='55%'><span style='font-size: 9pt; font-family: 新細明體'>第" & nowPage & "/" & totPage & "頁&nbsp;&nbsp;</span>"
    Response.Write "<span style='font-size: 9pt; font-family: 新細明體'>共" & FormatNumber(totRec,0) & "筆&nbsp;&nbsp;</span>"
    If cint(nowPage) <>1 Then
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage-1)&"&SQL="&SQL&"'>上一頁</a>" 
    End If
    If cint(nowPage)<>RS1.PageCount and cint(nowPage)<RS1.PageCount Then 
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage+1)&"&SQL="&SQL&"'>下一頁</a>" 
    End If
    Response.Write " |&nbsp;<span style='font-size: 9pt; font-family: 新細明體'> 跳至第<select name=GoPage size='1' style='font-size: 9pt; font-family: 新細明體' onchange='GoPage_OnChange(this.value)'>"
    For iPage=1 to totPage
      If iPage=cint(nowPage) Then
        strSelected = "selected"
      Else
	      strSelected = ""
      End If
      Response.Write "<option value='"&iPage&"'" & strSelected & ">" & iPage & "</option>"
    Next
    Response.Write "</select>頁</span></td>" 
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    End If   
    Response.Write "</tr></table>"
    Dim I
    Dim J
    Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    For J = 1 To FieldsCount
      Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J).Name & "</span></font></td>"
    Next
    Response.Write "</tr>"
    I = 1
    While Not RS1.EOF And I <= RS1.PageSize
      If Hlink<>"" Then
        Response.Write "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>"
        showhand="style='cursor:hand'"
      End If
      Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
      For J = 1 To FieldsCount
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(J))&"</span></td>"
      Next
      I = I + 1
      RS1.MoveNext
      Response.Write "</tr>"
      Response.Write "</a>"
    Wend
    Response.Write "</table>"
    Response.Write "</form>"
  Else
    Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 沒有符合條件的資料 **</td>"
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    End If
    Response.Write "</table>"
  End If
  RS1.Close
  Set RS1=Nothing
End Function

Function GridList_S (AddLink)
  Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
  Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 請輸入查詢條件 **</td>"
  If AddLink <> "" Then
    Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
  End If  
  Response.Write "</table>"
End Function

Function GridList_Q (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  If Not RS1.EOF Then 
    FieldsCount = RS1.Fields.Count-1
    totRec=RS1.Recordcount
    If totRec>0 Then 
      RS1.PageSize=PageSize
      If nowPage="" or nowPage=0 Then
        nowPage=1
      ElseIf cint(nowPage) > RS1.PageCount Then 
        nowPage=RS1.PageCount 
      End If
      session("nowPage")=nowPage
      RS1.AbsolutePage=nowPage
      totPage=RS1.PageCount
      SQL=server.URLEncode(SQL)
    End If
    Response.Write "<Form action='' id=form1 name=form1>" 
    Response.Write "<table border=0 cellspacing='0' cellpadding='1' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='30%'></td>"
    Response.Write "<td width='55%'><span style='font-size: 9pt; font-family: 新細明體'>第" & nowPage & "/" & totPage & "頁&nbsp;&nbsp;</span>"
    Response.Write "<span style='font-size: 9pt; font-family: 新細明體'>共" & FormatNumber(totRec,0) & "筆&nbsp;&nbsp;</span>"
    If cint(nowPage) <>1 Then
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage-1)&"&SQL="&SQL&"'>上一頁</a>" 
    End If
    If cint(nowPage)<>RS1.PageCount and cint(nowPage)<RS1.PageCount Then 
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage+1)&"&SQL="&SQL&"'>下一頁</a>" 
    End If
    Response.Write " |&nbsp;<span style='font-size: 9pt; font-family: 新細明體'> 跳至第<select name=GoPage size='1' style='font-size: 9pt; font-family: 新細明體' onchange='GoPage_OnChange(this.value)'>"
    For iPage=1 to totPage
      If iPage=cint(nowPage) Then
        strSelected = "selected"
      Else
	      strSelected = ""
      End If
      Response.Write "<option value='"&iPage&"'" & strSelected & ">" & iPage & "</option>"
    Next
    Response.Write "</select>頁</span></td>" 
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>回查詢</a></span></td>"
    End If   
    Response.Write "</tr></table>"
    Dim I
    Dim J
    Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    For J = 1 To FieldsCount
      Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J).Name & "</span></font></td>"
    Next
    Response.Write "</tr>"
    I = 1
    While Not RS1.EOF And I <= RS1.PageSize
      If Hlink<>"" Then
        Response.Write "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>"
        showhand="style='cursor:hand'"
      End If
      Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
      For J = 1 To FieldsCount
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
      Next
      I = I + 1
      RS1.MoveNext
      Response.Write "</tr>"
      Response.Write "</a>"
    Wend
    Response.Write "</table>"
    Response.Write "</form>"
  Else
    Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 沒有符合條件的資料 **</td>"
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>回查詢</a></span></td>"
    End If
    Response.Write "</table>"
  End If
  RS1.Close
  Set RS1=Nothing
End Function

Function ReportList (SQL,ReportName)
  Response.Write "<center><span style='font-size: 12pt; font-family: 標楷體'>"&ReportName&"</span></center>"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id=grid border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='98%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
	  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr>"
	  For I = 1 To FieldsCount
      Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(I))&"</span></td>"
	  Next
	  Response.Flush
    Response.Clear
    RS1.MoveNext
	  Response.Write "</tr>"
  Wend
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function 

Function SubjectList (SQL,HLink,LinkParam,LinkTarget,LinkType)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id=grid border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
    Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i).Name & "</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
    If Hlink<>"" Then
      If LinkType="window" Then
        Response.Write "<a href onclick=""window.open('" & Hlink & RS1(LinkParam)&"','','scrollbars=no,top=100,left=120,width=470,height=320')"">"
      Else
        Response.Write "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>"
      End If   
      showhand="style='cursor:hand'"
    End If
    Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
    For I = 1 To FieldsCount
      Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i) & "</span></td>"
    Next    
    RS1.MoveNext
    Response.Write "</tr>"
    Response.Write "</a>"
  Wend
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function

Function ImageList (SQL,HLink,LinkParam,LinkTarget)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id=grid border=0 cellspacing='0' cellpadding='3' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
    Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i).Name & "</span></font></td>"
  Next
  Response.Write "<td bgcolor='#FFE1AF'><span style='font-size: 9pt; font-family: 新細明體'></span></td>"
  Response.Write "</tr>"
  While Not RS1.EOF
    Response.Write "<tr>"
    For I = 1 To FieldsCount
      Response.Write "<td bgcolor='#FFFFFF' align='center'><a onClick=""window.open('image_left_show.asp?imgfile="&RS1(i)&"','','scrollbars=yes,resizable=yes,top=20,left=40,width=280,height=180')""><img src='../upload/"&RS1(i)&"' border=0 width='90' height='67'></a></td>"
    Next
    Response.Write "<td valign='bottom'><a href='JavaScript:if(confirm(""是否確定要刪除圖片 ?"")){window.location.href=""" & HLink & RS1(LinkParam) & """;}' target='" & LinkTarget &"'><img src='../images/x5.gif' border=0 width='16' height='14' alt='刪除'></a></td>"
    RS1.MoveNext
    Response.Write "</tr>"
  Wend
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function

function DundasUpLoad(FilePath,UploadName,UploadSize,objUpload,MaxFileSize)
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
      SourcePath = objUploadedFile.Originalpath
      SourceName = objUpload.GetFileName(objUploadedFile.Originalpath)
      UploadName = objUpload.GetFileName(objUploadedFile.path)
      UploadSize = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
      ExtName = objUpload.GetFileExt(objUploadedFile.Originalpath)
    End If
  Next
End function

function DundasUpLoad5(FilePath,UploadName1,UploadSize1,ExtName1,UploadName2,UploadSize2,ExtName2,UploadName3,UploadSize3,ExtName3,UploadName4,UploadSize4,ExtName4,UploadName5,UploadSize5,ExtName5,objUpload,MaxFileSize)
  set objUpload = Server.CreateObject("Dundas.Upload")
  on error resume next
  objUpload.MaxFileCount = 5
  objUpload.MaxFileSize = MaxFileSize*1048576
  objUpload.MaxUploadSize = 5*MaxFileSize*1048576
  objUpload.UseUniqueNames = True
  objUpload.UseVirtualDir = True
  objUpload.Save FilePath
  
  I=0
  For Each objUploadedFile in objUpload.Files
    I = I + 1
    If objUploadedFile.Size>0 Then
      If I=1 Then
        SourcePath1 = objUploadedFile.Originalpath
        SourceName1 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName1 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize1 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName1 = objUpload.GetFileExt(objUploadedFile.Originalpath)
      ElseIf I=2 Then
        SourcePath2 = objUploadedFile.Originalpath
        SourceName2 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName2 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize2 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName2 = objUpload.GetFileExt(objUploadedFile.Originalpath)      
      ElseIf I=3 Then
        SourcePath3 = objUploadedFile.Originalpath
        SourceName3 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName3 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize3 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName3 = objUpload.GetFileExt(objUploadedFile.Originalpath)      
      ElseIf I=4 Then
        SourcePath4 = objUploadedFile.Originalpath
        SourceName4 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName4 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize4 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName4 = objUpload.GetFileExt(objUploadedFile.Originalpath)      
      ElseIf I=5 Then
        SourcePath5 = objUploadedFile.Originalpath
        SourceName5 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName5 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize5 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName5 = objUpload.GetFileExt(objUploadedFile.Originalpath)      
      End If
    End If
  Next
End function

function DundasUpLoad10(FilePath,UploadName1,UploadSize1,ExtName1,UploadName2,UploadSize2,ExtName2,UploadName3,UploadSize3,ExtName3,UploadName4,UploadSize4,ExtName4,UploadName5,UploadSize5,ExtName5,UploadName6,UploadSize6,ExtName6,UploadName7,UploadSize7,ExtName7,UploadName8,UploadSize8,ExtName8,UploadName9,UploadSize9,ExtName9,UploadName10,UploadSize10,ExtName10,objUpload,MaxFileSize)
  set objUpload = Server.CreateObject("Dundas.Upload")
  on error resume next
  objUpload.MaxFileCount = 10
  objUpload.MaxFileSize = MaxFileSize*1048576
  objUpload.MaxUploadSize = 10*MaxFileSize*1048576
  objUpload.UseUniqueNames = True
  objUpload.UseVirtualDir = True
  objUpload.Save FilePath
  
  I=0
  For Each objUploadedFile in objUpload.Files
    I = I + 1
    If objUploadedFile.Size>0 Then
      If I=1 Then
        SourcePath1 = objUploadedFile.Originalpath
        SourceName1 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName1 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize1 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName1 = objUpload.GetFileExt(objUploadedFile.Originalpath)
      ElseIf I=2 Then
        SourcePath2 = objUploadedFile.Originalpath
        SourceName2 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName2 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize2 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName2 = objUpload.GetFileExt(objUploadedFile.Originalpath)
      ElseIf I=3 Then
        SourcePath3 = objUploadedFile.Originalpath
        SourceName3 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName3 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize3 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName3 = objUpload.GetFileExt(objUploadedFile.Originalpath)
      ElseIf I=4 Then
        SourcePath4 = objUploadedFile.Originalpath
        SourceName4 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName4 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize4 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName4 = objUpload.GetFileExt(objUploadedFile.Originalpath)
      ElseIf I=5 Then
        SourcePath5 = objUploadedFile.Originalpath
        SourceName5 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName5 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize5 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName5 = objUpload.GetFileExt(objUploadedFile.Originalpath)
      ElseIf I=6 Then
        SourcePath6 = objUploadedFile.Originalpath
        SourceName6 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName6 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize6 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName6 = objUpload.GetFileExt(objUploadedFile.Originalpath)
      ElseIf I=7 Then
        SourcePath7 = objUploadedFile.Originalpath
        SourceName7 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName7 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize7 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName7 = objUpload.GetFileExt(objUploadedFile.Originalpath)
      ElseIf I=8 Then
        SourcePath8 = objUploadedFile.Originalpath
        SourceName8 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName8 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize8 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName8 = objUpload.GetFileExt(objUploadedFile.Originalpath)
      ElseIf I=9 Then
        SourcePath9 = objUploadedFile.Originalpath
        SourceName9 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName9 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize9 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName9 = objUpload.GetFileExt(objUploadedFile.Originalpath)
      ElseIf I=10 Then
        SourcePath10 = objUploadedFile.Originalpath
        SourceName10 = objUpload.GetFileName(objUploadedFile.Originalpath)
        UploadName10 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize10 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
        ExtName10 = objUpload.GetFileExt(objUploadedFile.Originalpath)
      End If
    End If
  Next
End function

Function ShowFlv(FileName,FileType,xWidth,xHeight)
  If xHeight<>"" Then xHeight="height="&xHeight
  If Filetype="flash" Then
    Response.write "<object classid='clsid:D27CDB6E-AE6D-11CF-96B8-444553540000' id='obj1' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0' border='0' width='"&xWidth&"'>"
    Response.write "<param name='movie' value='"&FileName&"'>"
    Response.write "<param name='quality' value='High'>"
    Response.write "<embed src='"&FileName&"' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' name='obj1' width='"&xWidth&"' quality='High'></object>"
  Elseif FileType="wmv" Then
    Response.write "<EMBED src='"&FileName&"' autostart=true controls=console width='"&xWidth&"' "&xHeight&" type=video/x-ms-wmv></EMBED>" 
  Elseif FileType="flv" Then
    Response.write "<embed src='../include/vcastr.swf?vcastr_file="&FileName&"' allowfullscreen='true' showMovieInfo=0 pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' wmode='transparent' quality='high' width='"&xWidth&"' "&xHeight&"></embed>" 
  End If
End Function

Function ShowMedia(FileName,FileType,xWidth,xHeight)
  If xHeight<>"" Then xHeight="height="&xHeight
  If Filetype="flash" Then
    Response.write "<object classid='clsid:D27CDB6E-AE6D-11CF-96B8-444553540000' id='obj1' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0' border='0' width='"&xWidth&"'>"
    Response.write "<param name='movie' value='"&FileName&"'>"
    Response.write "<param name='quality' value='High'>"
    Response.write "<embed src='"&FileName&"' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' name='obj1' width='"&xWidth&"' quality='High'></object>"
  Elseif FileType="wmv" Then
    Response.write "<EMBED src='"&FileName&"' autostart=true controls=console width='"&xWidth&"' "&xHeight&" type=video/x-ms-wmv></EMBED>" 
  Elseif FileType="flv" Then
    Response.write "<embed src='../include/vcastr.swf?vcastr_file="&FileName&"' allowfullscreen='true' showMovieInfo=0 pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' wmode='transparent' quality='high' width='"&xWidth&"' "&xHeight&"></embed>" 
  End If
End Function

Function CodeCity (form,ZipCode,ZipValue,CityCode,CityValue,AreaCode,AreaValue,Address,AddressValue,AddressSize,JS)
  '縣市選單
  RS1Row = 0
  SelCity="<option value=''>縣&nbsp;&nbsp;&nbsp;&nbsp;市</option>"
  SQL1 = "Select mCode,mValue From CODECITY Where codeMetaID = 'Addr0' Order By Seq,mSortValue"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  Do While Not RS1.EOF
    If RS1("mCode")=CityValue Then
      SelCity = SelCity & "<option value='" & RS1("mCode") & "' Selected>" & RS1("mValue") & "</option>"
    Else
      SelCity = SelCity & "<option value='" & RS1("mCode") & "'>" & RS1("mValue") & "</option>"
    End If
    
    If RS1Row = 0 Then
      Aera_Code = Aera_Code & "'" & RS1("mCode") & "'" & ","
    Else 
      Aera_Code = Aera_Code & "," & "'" & RS1("mCode") & "'" & ","
    End If
    RS2Row=1
    Aera_Code = Aera_Code & "["
    SQL2 = "Select mCode,mValue From CODECITY Where codeMetaID = 'Addr0R"&RS1("mCode")&"' Order By mSortValue"
    Set RS2 = Server.CreateObject("ADODB.RecordSet")
    RS2.Open SQL2,Conn,1,1
    Do While Not RS2.EOF
      If RS2Row = 1 Then
        Aera_Code = Aera_Code & "'" & RS2("mCode") & "'" & "," & "'" & RS2("mValue") & "'" 
      Else 
        Aera_Code = Aera_Code & "," & "'" & RS2("mCode") & "'" & "," & "'" & RS2("mValue") & "'" 
      End If
      RS2Row=RS2Row+1
      RS2.MoveNext
    Loop
    RS2.Close
    Set RS2 = Nothing
    Aera_Code=Aera_Code&"]"
    RS1Row=RS1Row+1
    RS1.MoveNext
  Loop
  RS1.Close
  Set RS1 = Nothing  
  '鄉鎮市區選單
  SelArea="<option value=''>鄉鎮市區</option>"
  If CityValue<>"" And AreaValue<>"" Then
    SQL1 = "Select mCode,mValue From CodeCity Where codeMetaID = 'Addr0R"&CityValue&"' Order By mSortValue"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    Do While Not RS1.EOF
      If RS1("mCode")=AreaValue Then
        SelArea = SelArea & "<option value='" & RS1("mCode") & "' Selected>" & RS1("mValue") & "</option>"
      Else
        SelArea = SelArea & "<option value='" & RS1("mCode") & "'>" & RS1("mValue") & "</option>"
      End If  
      RS1.MoveNext
    Loop 
    RS1.Close
    Set RS1 = Nothing
  End If
      
  Response.Write "<input type='text' name='"&ZipCode&"' size='3' class='font9' readonly maxlength='3' value='"&ZipValue&"'>&nbsp;"
  Response.Write "<Select name='"&CityCode&"' Size='1' style='font-size: 9pt; font-family: 新細明體' onchange='JavaScript:ChgCity(this.value,document."&form&"."&AreaCode&",document."&form&"."&ZipCode&");'>"&SelCity&"</Select>&nbsp;"
  Response.Write "<Select name='"&AreaCode&"' Size='1' style='font-size: 9pt; font-family: 新細明體' onchange='JavaScript:ChgArea(this.value,document."&form&"."&ZipCode&");'>"&SelArea&"</Select>&nbsp;"
  Response.Write "<input type='text' class='font9' name='"&Address&"' size='"&AddressSize&"' maxlength='80' onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}' value='"&AddressValue&"'>" & Chr(13) & Chr(10) 
  If JS="Y" Then
    '縣市/鄉鎮市區選單連動
    Response.Write "<script language=""JavaScript""><!--" & Chr(13) & Chr(10)
    Response.Write "  function ChgCity(CityCode,AreaObj,ZipObj){" & Chr(13) & Chr(10)
    Response.Write "    ZipObj.value='';" & Chr(13) & Chr(10)
    Response.Write "    ClearOption(AreaObj);" & Chr(13) & Chr(10)
    Response.Write "    AreaObj.options[0] = new Option('鄉鎮市區','');" & Chr(13) & Chr(10)
    Response.Write "    var AryAreaCode = new Array("&Aera_Code&")" & Chr(13) & Chr(10)
    Response.Write "    for(var i=0;i<=AryAreaCode.length-1;i=i+2){" & Chr(13) & Chr(10)
    Response.Write "      if(AryAreaCode[i]==CityCode){" & Chr(13) & Chr(10)
    Response.Write "        var theArea = AryAreaCode[i+1];" & Chr(13) & Chr(10)
    Response.Write "        var k=1;" & Chr(13) & Chr(10)
    Response.Write "        for(var j=0;j<=theArea.length-1;j=j+2){" & Chr(13) & Chr(10)
    Response.Write "          AreaObj.options[k] = new Option(theArea[j+1],theArea[j]);" & Chr(13) & Chr(10)
    Response.Write "          k++;" & Chr(13) & Chr(10)
    Response.Write "        }" & Chr(13) & Chr(10)
    Response.Write "      }" & Chr(13) & Chr(10)
    Response.Write "    }" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10) 
    Response.Write "  function ClearOption(SelObj){" & Chr(13) & Chr(10)
    Response.Write "    var i;" & Chr(13) & Chr(10)
    Response.Write "    for(i=SelObj.length-1;i>=0;i--){" & Chr(13) & Chr(10)
    Response.Write "      SelObj.options[i] = null;" & Chr(13) & Chr(10)
    Response.Write "    }" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "  function ChgArea(AreaCode,ZipObj){" & Chr(13) & Chr(10)
    Response.Write "    ZipObj.value=AreaCode.substring(0,3);" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "  function ChgAddress(AddressObj){" & Chr(13) & Chr(10)
    Response.Write "    if(AddressObj=='請輸入路段號'){" & Chr(13) & Chr(10)
    Response.Write "      AddressObj='';" & Chr(13) & Chr(10)
    Response.Write "    }" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "--></script>" & Chr(13) & Chr(10)
  End If
End Function

Function CodeCity2 (form,ZipCode,ZipValue,CityCode,CityValue,AreaCode,AreaValue,JS)
  '縣市選單
  RS1Row = 0
  SelCity="<option value=''>縣&nbsp;&nbsp;&nbsp;&nbsp;市</option>"
  SQL1 = "Select mCode,mValue,telCode From CODECITY Where codeMetaID = 'Addr0' Order By Seq,mSortValue"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  Do While Not RS1.EOF
    If RS1("mCode")=CityValue Then
      SelCity = SelCity & "<option value='" & RS1("mCode") & "' Selected>" & RS1("mValue") & "</option>"
	  SelTelCode = SelTelCode & "<option value='" & RS1("mCode") & "' Selected>" & RS1("telCode") & "</option>"
    Else
      SelCity = SelCity & "<option value='" & RS1("mCode") & "'>" & RS1("mValue") & "</option>"
	  SelTelCode = SelTelCode & "<option value='" & RS1("mCode") & "'>" & RS1("telCode") & "</option>"
    End If
    
    If RS1Row = 0 Then
      Aera_Code = Aera_Code & "'" & RS1("mCode") & "'" & ","
    Else 
      Aera_Code = Aera_Code & "," & "'" & RS1("mCode") & "'" & ","
    End If
    RS2Row=1
    Aera_Code = Aera_Code & "["
    SQL2 = "Select mCode,mValue From CODECITY Where codeMetaID = 'Addr0R"&RS1("mCode")&"' Order By mSortValue"
    Set RS2 = Server.CreateObject("ADODB.RecordSet")
    RS2.Open SQL2,Conn,1,1
    Do While Not RS2.EOF
      If RS2Row = 1 Then
        Aera_Code = Aera_Code & "'" & RS2("mCode") & "'" & "," & "'" & RS2("mValue") & "'" 
      Else 
        Aera_Code = Aera_Code & "," & "'" & RS2("mCode") & "'" & "," & "'" & RS2("mValue") & "'" 
      End If
      RS2Row=RS2Row+1
      RS2.MoveNext
    Loop
    RS2.Close
    Set RS2 = Nothing
    Aera_Code=Aera_Code&"]"
    RS1Row=RS1Row+1
    RS1.MoveNext
  Loop
  RS1.Close
  Set RS1 = Nothing  
  '鄉鎮市區選單
  SelArea="<option value=''>鄉鎮市區</option>"
  If CityValue<>"" And AreaValue<>"" Then
    SQL1 = "Select mCode,mValue From CodeCity Where codeMetaID = 'Addr0R"&CityValue&"' Order By mSortValue"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    Do While Not RS1.EOF
      If RS1("mCode")=AreaValue Then
        SelArea = SelArea & "<option value='" & RS1("mCode") & "' Selected>" & RS1("mValue") & "</option>"
      Else
        SelArea = SelArea & "<option value='" & RS1("mCode") & "'>" & RS1("mValue") & "</option>"
      End If  
      RS1.MoveNext
    Loop 
    RS1.Close
    Set RS1 = Nothing
  End If
      
  Response.Write "<input type='text' name='"&ZipCode&"' size='3' class='font9' maxlength='5' value='"&ZipValue&"' onKeypress='if(event.keyCode==13){JavaScript:ChgZip(this.value,document."&form&"."&CityCode&",document."&form&"."&AreaCode&");}'>&nbsp;"
  Response.Write "<Select name='"&CityCode&"' Size='1' style='font-size: 9pt; font-family: 新細明體' onchange='JavaScript:ChgCity(this.value,document."&form&"."&AreaCode&",document."&form&"."&ZipCode&");'>"&SelCity&"</Select>&nbsp;"
  Response.Write "<Select name='"&AreaCode&"' Size='1' style='font-size: 9pt; font-family: 新細明體' onchange='JavaScript:ChgArea(this.value,document."&form&"."&ZipCode&");'>"&SelArea&"</Select>&nbsp;"
  'Response.Write "<input type='text' class='font9' name='"&Address&"' size='"&AddressSize&"' maxlength='80' value='"&AddressValue&"'>" & Chr(13) & Chr(10) 
  '2013/12/19 Modify by GoodTV 詩儀:增加區碼連動
  If JS="Y" Then
	TelCode="TelCode"
  Else
	TelCode="Invoice_TelCode"
  End If
  Response.Write "<Select name='"&TelCode&"' Size='1' style='font-size: 1pt; font-family: 新細明體' >"&SelTelCode&"</Select>&nbsp;"
  Response.Write "<script language=""JavaScript""><!--" & Chr(13) & Chr(10)
  Response.Write ""&form&"."&TelCode&".style.visibility=""hidden"";" & Chr(13) & Chr(10) '隱藏TelCode區碼下拉式選單
  Response.Write "--></script>" & Chr(13) & Chr(10)
  
  If JS="Y" Then
    '縣市/鄉鎮市區/區碼選單連動
    Response.Write "<script language=""JavaScript""><!--" & Chr(13) & Chr(10)
    Response.Write "  function ChgZip(ZipCode,CityObj,AreaObj){" & Chr(13) & Chr(10)
    Response.Write "    if(ZipCode.length>=3){" & Chr(13) & Chr(10)
    Response.Write "      var AryAreaCode = new Array("&Aera_Code&")" & Chr(13) & Chr(10)
    Response.Write "      for(var i=1;i<=AryAreaCode.length-1;i=i+2){" & Chr(13) & Chr(10)
    Response.Write "        var theArea=AryAreaCode[i];" & Chr(13) & Chr(10)
    Response.Write "        for(var j=0;j<=theArea.length-1;j=j+2){" & Chr(13) & Chr(10)
    Response.Write "          if(theArea[j]==ZipCode.substring(0,3)){" & Chr(13) & Chr(10)
    Response.Write "            CityObj.value=AryAreaCode[i-1];" & Chr(13) & Chr(10)
	'------------------------------------------增加區碼連動開始-------------------------------------
	Response.Write "            document."&form&"."&TelCode&".value = CityObj.value;" & Chr(13) & Chr(10)
	Response.Write "            var objSelect = document."&form&"."&TelCode&";" & Chr(13) & Chr(10)
	Response.Write "            var objText =  objSelect.options[objSelect.selectedIndex].text;" & Chr(13) & Chr(10)
	Response.Write "            document."&form&".Tel_Office_Loc.value = objText;" & Chr(13) & Chr(10)
	Response.Write "            document."&form&".Fax_Loc.value = objText;" & Chr(13) & Chr(10)
	'------------------------------------------增加區碼連動結束-------------------------------------
    Response.Write "            ClearOption(AreaObj);" & Chr(13) & Chr(10)
    Response.Write "            AreaObj.options[0] = new Option('鄉鎮市區','');" & Chr(13) & Chr(10)
    Response.Write "            var l=1;" & Chr(13) & Chr(10)
    Response.Write "            for(var k=0;k<=theArea.length-1;k=k+2){" & Chr(13) & Chr(10)
    Response.Write "              AreaObj.options[l] = new Option(theArea[k+1],theArea[k]);" & Chr(13) & Chr(10)
    Response.Write "              l++;" & Chr(13) & Chr(10)
    Response.Write "            }" & Chr(13) & Chr(10)
    Response.Write "            AreaObj.value=ZipCode.substring(0,3);" & Chr(13) & Chr(10)
    Response.Write "            j=theArea.length;" & Chr(13) & Chr(10)
    Response.Write "            i=AryAreaCode.length;" & Chr(13) & Chr(10)
    Response.Write "          }" & Chr(13) & Chr(10)
    Response.Write "        }" & Chr(13) & Chr(10)
    Response.Write "      }" & Chr(13) & Chr(10)
    Response.Write "    }else{" & Chr(13) & Chr(10)
    Response.Write "      CityObj.value='';" & Chr(13) & Chr(10)
    Response.Write "      ClearOption(AreaObj);" & Chr(13) & Chr(10)
    Response.Write "      AreaObj.options[0] = new Option('鄉鎮市區','');" & Chr(13) & Chr(10)
    Response.Write "    }" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "  function ChgCity(CityCode,AreaObj,ZipObj){" & Chr(13) & Chr(10)
    Response.Write "    ZipObj.value='';" & Chr(13) & Chr(10)
    Response.Write "    ClearOption(AreaObj);" & Chr(13) & Chr(10)
    Response.Write "    AreaObj.options[0] = new Option('鄉鎮市區','');" & Chr(13) & Chr(10)
    Response.Write "    var AryAreaCode = new Array("&Aera_Code&")" & Chr(13) & Chr(10)
    Response.Write "    for(var i=0;i<=AryAreaCode.length-1;i=i+2){" & Chr(13) & Chr(10)
    Response.Write "      if(AryAreaCode[i]==CityCode){" & Chr(13) & Chr(10)
    Response.Write "        var theArea = AryAreaCode[i+1];" & Chr(13) & Chr(10)
    Response.Write "        var k=1;" & Chr(13) & Chr(10)
    Response.Write "        for(var j=0;j<=theArea.length-1;j=j+2){" & Chr(13) & Chr(10)
    Response.Write "          AreaObj.options[k] = new Option(theArea[j+1],theArea[j]);" & Chr(13) & Chr(10)
    Response.Write "          k++;" & Chr(13) & Chr(10)
    Response.Write "        }" & Chr(13) & Chr(10)
    Response.Write "        i=AryAreaCode.length" & Chr(13) & Chr(10)
	'------------------------------------------增加區碼連動開始-------------------------------------
	Response.Write "        document."&form&"."&TelCode&".value = document."&form&"."&CityCode&".value;" & Chr(13) & Chr(10)
	Response.Write "        var objSelect = document."&form&"."&TelCode&";" & Chr(13) & Chr(10)
	Response.Write "        var objText =  objSelect.options[objSelect.selectedIndex].text;" & Chr(13) & Chr(10)
	Response.Write "        document."&form&".Tel_Office_Loc.value = objText;" & Chr(13) & Chr(10)
	Response.Write "        document."&form&".Fax_Loc.value = objText;" & Chr(13) & Chr(10)
	'------------------------------------------增加區碼連動結束-------------------------------------
    Response.Write "      }" & Chr(13) & Chr(10)
    Response.Write "    }" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10) 
    Response.Write "  function ClearOption(SelObj){" & Chr(13) & Chr(10)
    Response.Write "    var i;" & Chr(13) & Chr(10)
    Response.Write "    for(i=SelObj.length-1;i>=0;i--){" & Chr(13) & Chr(10)
    Response.Write "      SelObj.options[i] = null;" & Chr(13) & Chr(10)
    Response.Write "    }" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "  function ChgArea(AreaCode,ZipObj){" & Chr(13) & Chr(10)
    Response.Write "    ZipObj.value=AreaCode.substring(0,3);" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "  function ChgAddress(AddressObj){" & Chr(13) & Chr(10)
    Response.Write "    if(AddressObj=='請輸入路段號'){" & Chr(13) & Chr(10)
    Response.Write "      AddressObj='';" & Chr(13) & Chr(10)
    Response.Write "    }" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "--></script>" & Chr(13) & Chr(10)
  End If
End Function

Function CodeArea (form,CityCode,CityValue,AreaCode,AreaValue,JS)
  '縣市選單
  RS1Row = 0
  SelCity="<option value=''>縣&nbsp;&nbsp;&nbsp;&nbsp;市</option>"
  SQL1 = "Select mCode,mValue From CODECITY Where codeMetaID = 'Addr0' Order By Seq,mSortValue"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  Do While Not RS1.EOF
    If RS1("mCode")=CityValue Then
      SelCity = SelCity & "<option value='" & RS1("mCode") & "' Selected>" & RS1("mValue") & "</option>"
    Else
      SelCity = SelCity & "<option value='" & RS1("mCode") & "'>" & RS1("mValue") & "</option>"
    End If
    
    If RS1Row = 0 Then
      Aera_Code = Aera_Code & "'" & RS1("mCode") & "'" & ","
    Else 
      Aera_Code = Aera_Code & "," & "'" & RS1("mCode") & "'" & ","
    End If
    RS2Row=1
    Aera_Code = Aera_Code & "["
    SQL2 = "Select mCode,mValue From CODECITY Where codeMetaID = 'Addr0R"&RS1("mCode")&"' Order By mSortValue"
    Set RS2 = Server.CreateObject("ADODB.RecordSet")
    RS2.Open SQL2,Conn,1,1
    Do While Not RS2.EOF
      If RS2Row = 1 Then
        Aera_Code = Aera_Code & "'" & RS2("mCode") & "'" & "," & "'" & RS2("mValue") & "'" 
      Else 
        Aera_Code = Aera_Code & "," & "'" & RS2("mCode") & "'" & "," & "'" & RS2("mValue") & "'" 
      End If
      RS2Row=RS2Row+1
      RS2.MoveNext
    Loop
    RS2.Close
    Set RS2 = Nothing
    Aera_Code=Aera_Code&"]"
    RS1Row=RS1Row+1
    RS1.MoveNext
  Loop
  RS1.Close
  Set RS1 = Nothing
  '鄉鎮市區選單
  SelArea="<option value=''>鄉鎮市區</option>"
  If CityValue<>"" And AreaValue<>"" Then
    SQL1 = "Select mCode,mValue From CodeCity Where codeMetaID = 'Addr0R"&CityValue&"' Order By mSortValue"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    Do While Not RS1.EOF
      If RS1("mCode")=AreaValue Then
        SelArea = SelArea & "<option value='" & RS1("mCode") & "' Selected>" & RS1("mValue") & "</option>"
      Else
        SelArea = SelArea & "<option value='" & RS1("mCode") & "'>" & RS1("mValue") & "</option>"
      End If  
      RS1.MoveNext
    Loop 
    RS1.Close
    Set RS1 = Nothing
  End If

  Response.Write "<Select name='"&CityCode&"' Size='1' style='font-size: 9pt; font-family: 新細明體' onchange='JavaScript:ChgCity(this.value,document."&form&"."&AreaCode&");'>"&SelCity&"</Select>&nbsp;"
  Response.Write "<Select name='"&AreaCode&"' Size='1' style='font-size: 9pt; font-family: 新細明體'>"&SelArea&"</Select>&nbsp;"
  If JS="Y" Then
    '縣市/鄉鎮市區選單連動
    Response.Write "<script language=""JavaScript""><!--" & Chr(13) & Chr(10)
    Response.Write "  function ChgCity(CityCode,AreaObj){" & Chr(13) & Chr(10)
    Response.Write "    ClearOption(AreaObj);" & Chr(13) & Chr(10)
    Response.Write "    AreaObj.options[0] = new Option('鄉鎮市區','');" & Chr(13) & Chr(10)
    Response.Write "    var AryAreaCode = new Array("&Aera_Code&")" & Chr(13) & Chr(10)
    Response.Write "    for(var i=0;i<=AryAreaCode.length-1;i=i+2){" & Chr(13) & Chr(10)
    Response.Write "      if(AryAreaCode[i]==CityCode){" & Chr(13) & Chr(10)
    Response.Write "        var theArea = AryAreaCode[i+1];" & Chr(13) & Chr(10)
    Response.Write "        var k=1;" & Chr(13) & Chr(10)
    Response.Write "        for(var j=0;j<=theArea.length-1;j=j+2){" & Chr(13) & Chr(10)
    Response.Write "          AreaObj.options[k] = new Option(theArea[j+1],theArea[j]);" & Chr(13) & Chr(10)
    Response.Write "          k++;" & Chr(13) & Chr(10)
    Response.Write "        }" & Chr(13) & Chr(10)
    Response.Write "      }" & Chr(13) & Chr(10)
    Response.Write "    }" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10) 
    Response.Write "  function ClearOption(SelObj){" & Chr(13) & Chr(10)
    Response.Write "    var i;" & Chr(13) & Chr(10)
    Response.Write "    for(i=SelObj.length-1;i>=0;i--){" & Chr(13) & Chr(10)
    Response.Write "      SelObj.options[i] = null;" & Chr(13) & Chr(10)
    Response.Write "    }" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "--></script>" & Chr(13) & Chr(10)
  End If
End Function

Function Web_Style(Style_Type,Style_Name)
  Web_Style=""
  Style_SQL="Select Style_Name,Style_Value From WEB_STYLE Where Style_Type='"&Style_Type&"' "
  If Style_Name<>"" Then Style_SQL=Style_SQL&"And Style_Name='"&Style_Name&"' "
  Style_SQL=Style_SQL&"Order By Style_Seq"  
  Set Style_RS = Server.CreateObject("ADODB.RecordSet")
  Style_RS.Open Style_SQL,Conn,1,1
  If Not Style_RS.EOF Then
    While Not Style_RS.EOF
      If Web_Style="" Then
        If Style_Name="" Then
          Web_Style=Style_RS("Style_Name")&","&Style_RS("Style_Value")
        Else
          Web_Style=Style_RS("Style_Value")
        End If
      Else
        If Style_Name="" Then
          Web_Style=Web_Style&","&Style_RS("Style_Name")&","&Style_RS("Style_Value")
        Else
          Web_Style=Web_Style&","&Style_RS("Style_Value")
        End If
      End If
      Style_RS.MoveNext
		Wend
  End If
  Style_RS.Close
  Set Style_RS = Nothing
End Function

Function Get_Invoice_No(Dept_Id,Donate_Date,Act_Id)
  Get_Invoice_No=""
  SQL_Dept="Select * From Dept Where Dept_Id='"&Dept_Id&"'"
  Call QuerySQL(SQL_Dept,RS_Dept)
  If Not RS_Dept.EOF Then
    Invoice_Pre=RS_Dept("Invoice_Pre")
    Invoice_Rule_Len=RS_Dept("Invoice_Rule_Len")
  End If
  RS_Dept.Close
  Set RS_Dept=Nothing
  
  Donate_Year=Year(Donate_Date)
  Donate_Month=Month(Donate_Date)
  If Len(Donate_Month)=1 Then Donate_Month="0"&Donate_Month
  Donate_Day=Day(Donate_Date)
  If Len(Donate_Day)=1 Then Donate_Day="0"&Donate_Day
  InvoiceNo=Donate_Year&Donate_Month&Donate_Day
  
  SQL_InvoiceNo="Select Top 1 Invoice_No,Invoice_Right=Right(Invoice_No,"&Invoice_Rule_Len&") From DONATE Where Invoice_No<>'' And Invoice_Pre='"&Invoice_Pre&"' And Issue_Type_Keep<>'M' And (Issue_Type='' Or Issue_Type Is null Or Issue_Type='D') And Dept_Id='"&Dept_Id&"' And Year(Donate_Date)="&Donate_Year&" And Month(Donate_Date)="&Donate_Month&" And Day(Donate_Date)="&Donate_Day&" Order By Invoice_Right Desc"
  Call QuerySQL(SQL_InvoiceNo,RS_InvoiceNo)
  If RS_InvoiceNo.EOF Then
    Get_Invoice_No=Invoice_Pre&"/"&InvoiceNo&"0001"
  Else
    Invoice_Right=Cstr(Cdbl(RS_InvoiceNo("Invoice_Right"))+1)
    Invoice_Right=Left("0000",4-Len(Invoice_Right))&Invoice_Right
    Get_Invoice_No=Invoice_Pre&"/"&InvoiceNo&Invoice_Right
  End If
  RS_InvoiceNo.Close
  Set RS_InvoiceNo=Nothing
End Function

Function Get_Invoice_NoY(Invoice_Pre_Type,Dept_Id,Donate_Date,Invoice_Rule_Type2,DonateDesc,Invoice_Rule_Type,Invoice_Rule_YMD,Invoice_Rule_Len,Invoice_Rule_Pub)
  InvoiceNo=""
  InvoiceNo2=""
  For IRL=1 To Invoice_Rule_Len
    InvoiceNo2=InvoiceNo2&"0"
  Next  
  If Invoice_Rule_Type="R" Or Invoice_Rule_Type="A" Then
    Donate_Year=Cstr(Year(Donate_Date))
    Donate_Month=Cstr(Month(Donate_Date))
    If Len(Donate_Month)=1 Then Donate_Month="0"&Donate_Month
    Donate_Day=Cstr(Day(Donate_Date))
    If Len(Donate_Day)=1 Then Donate_Day="0"&Donate_Day
    If Invoice_Rule_Type="R" Then
      Donate_Date_RA=(Donate_Year-1911)
    Else
      Donate_Date_RA=Donate_Year
    End If
    SQL_InvoiceNo="Select Top 1 Invoice_No From DONATE Where Invoice_No<>'' And Issue_Type_Keep<>'M' And (Issue_Type='' Or Issue_Type Is null Or Issue_Type='D') "
    If Invoice_Rule_Pub="N" Then SQL_InvoiceNo=SQL_InvoiceNo&"And Dept_Id='"&Dept_Id&"' "
    If Cstr(Invoice_Pre_Type)="1" Then
      If Invoice_Rule_Type2<>"P" Then SQL_InvoiceNo=SQL_InvoiceNo&"And (Donate_Payment<>'"&DonateDesc&"' Or Donate_Payment Is Null) "
    Else
      If Invoice_Rule_Type2<>"P" Then SQL_InvoiceNo=SQL_InvoiceNo&"And Donate_Payment='"&DonateDesc&"' "
    End If
    If Invoice_Rule_YMD="Y" Then 
      SQL_InvoiceNo=SQL_InvoiceNo&"And Invoice_No Like '"&Donate_Date_RA&"%' "
      If Invoice_Rule_Type="R" Then
        InvoiceNo=Right("000"&Cstr(Donate_Date_RA),3)
      Else
        InvoiceNo=Cstr(Donate_Date_RA)
      End If
    ElseIf Invoice_Rule_YMD="M" Then
      SQL_InvoiceNo=SQL_InvoiceNo&"And Invoice_No Like '"&Donate_Date_RA&Donate_Month&"%' "
      If Invoice_Rule_Type="R" Then
        InvoiceNo=Right("000"&Cstr(Donate_Date_RA),3)&Right("00"&Cstr(Donate_Month),2)
      Else
        InvoiceNo=Cstr(Donate_Date_RA)&Right("00"&Cstr(Donate_Month),2)
      End If
    Else
      SQL_InvoiceNo=SQL_InvoiceNo&"And Invoice_No Like '"&Donate_Date_RA&Donate_Month&Donate_Day&"%' "
      If Invoice_Rule_Type="R" Then
        InvoiceNo=Right("000"&Cstr(Donate_Date_RA),3)&Right("00"&Cstr(Donate_Month),2)&Right("00"&Cstr(Donate_Day),2)
      Else
        InvoiceNo=Cstr(Donate_Date_RA)&Right("00"&Cstr(Donate_Month),2)&Right("00"&Cstr(Donate_Day),2)
      End If
    End If
    SQL_InvoiceNo=SQL_InvoiceNo&"Order By Right(Invoice_No,"&Invoice_Rule_Len&") Desc"
    Call QuerySQL(SQL_InvoiceNo,RS_InvoiceNo)
    If RS_InvoiceNo.EOF Then
      Get_Invoice_NoY=InvoiceNo&Right(InvoiceNo2&CStr(CDbl(Right(InvoiceNo2,Invoice_Rule_Len))+1),Invoice_Rule_Len)
    Else
      Get_Invoice_NoY=InvoiceNo&Right(InvoiceNo2&CStr(CDbl(Right(RS_InvoiceNo("Invoice_No"),Invoice_Rule_Len))+1),Invoice_Rule_Len)
    End If
    RS_InvoiceNo.Close
    Set RS_InvoiceNo=Nothing
  ElseIf Invoice_Rule_Type="S" Then
    SQL_InvoiceNo="Select Top 1 Invoice_No From DONATE Where Invoice_No<>'' And Issue_Type_Keep<>'M' And (Issue_Type='' Or Issue_Type Is null Or Issue_Type='D') And Convert(numeric,Right(Invoice_No,"&Invoice_Rule_Len&"))>0 "
    If Invoice_Rule_Pub="N" Then SQL_InvoiceNo=SQL_InvoiceNo&"And Dept_Id='"&Dept_Id&"' "
    If Cstr(Invoice_Pre_Type)="1" Then
      If Invoice_Rule_Type2<>"P" Then SQL_InvoiceNo=SQL_InvoiceNo&"And (Donate_Payment<>'"&DonateDesc&"' Or Donate_Payment Is Null) "
    Else
      If Invoice_Rule_Type2<>"P" Then SQL_InvoiceNo=SQL_InvoiceNo&"And Donate_Payment='"&DonateDesc&"' "
    End If
    SQL_InvoiceNo=SQL_InvoiceNo&"Order By Convert(numeric,Right(Invoice_No,"&Invoice_Rule_Len&")) Desc"
    Call QuerySQL(SQL_InvoiceNo,RS_InvoiceNo)
    If RS_InvoiceNo.EOF Then
      Get_Invoice_NoY=Right(InvoiceNo2&CStr(CDbl(Right(InvoiceNo2,Invoice_Rule_Len))+1),Invoice_Rule_Len)
    Else
      Get_Invoice_NoY=Right(InvoiceNo2&CStr(CDbl(Right(RS_InvoiceNo("Invoice_No"),Invoice_Rule_Len))+1),Invoice_Rule_Len)
    End If
    RS_InvoiceNo.Close
    Set RS_InvoiceNo=Nothing
  ElseIf Invoice_Rule_Type="O" Then
    '收據編號客制化
      
  End If
End Function

Function Get_Invoice_No2(Invoice_Kind,Dept_Id,Donate_Date,InvoiceType,Act_Id)
  Get_Invoice_No2=""
  
  InvoiceTypeY="年度匯整"
  SQL_InvoiceY="Select InvoiceTypeY=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Like '%年%' Order By Seq"
  Call QuerySQL(SQL_InvoiceY,RS_InvoiceY)
  If Not RS_InvoiceY.EOF Then InvoiceTypeY=RS_InvoiceY("InvoiceTypeY")
  RS_InvoiceY.Close
  Set RS_InvoiceY=Nothing
  
  '判斷有無設定取號規則
  Check_Invoice=False

  '勸募活動取號規則
  If Check_Invoice=False And Act_Id<>"" Then
    SQL_Act="Select * From ACT Where Act_Id='"&Act_Id&"' And Act_Flg='Y' And Act_Flg2='Y'"
    Call QuerySQL(SQL_Act,RS_Act)
    If Not RS_Act.EOF Then
      If RS_Act("Act_Invoice")="N" And (Cstr(InvoiceType)=Cstr(InvoiceTypeY)) Then Exit Function
      If Invoice_Kind="1" Then
        Invoice_Pre=RS_Act("Act_Pre")
        Invoice_Rule_Type=RS_Act("Act_Rule_Type")
        Invoice_Rule_YMD=RS_Act("Act_Rule_YMD")
        Invoice_Rule_Len=RS_Act("Act_Rule_Len")
        Invoice_Rule_Pub=RS_Act("Act_Rule_Pub")
        Invoice_Rule_Type2=RS_Act("Act_Rule_Type2")
        Check_Invoice=True
      ElseIf Invoice_Kind="2" Then
        If RS_Act("Act_Rule_Type2")="P" Then
          Invoice_Pre=RS_Act("Act_Pre")
          Invoice_Rule_Type=RS_Act("Act_Rule_Type")
          Invoice_Rule_YMD=RS_Act("Act_Rule_YMD")
          Invoice_Rule_Len=RS_Act("Act_Rule_Len")
          Invoice_Rule_Pub=RS_Act("Act_Rule_Pub")
          Invoice_Rule_Type2=RS_Act("Act_Rule_Type2")
        Else
          Invoice_Pre=RS_Act("Act_Pre2")
          Invoice_Rule_Type=RS_Act("Act_Rule_Type2")
          Invoice_Rule_YMD=RS_Act("Act_Rule_YMD2")
          Invoice_Rule_Len=RS_Act("Act_Rule_Len2")
          Invoice_Rule_Pub=RS_Act("Act_Rule_Pub2")
          Invoice_Rule_Type2=RS_Act("Act_Rule_Type2")
        End If
        Check_Invoice=True
      End If
    End If
    RS_Act.Close
    Set RS_Act=Nothing
  End If
    
  If Check_Invoice=False Then
    If Invoice_Kind="1" Then
      '捐款取號規則
      SQL_Dept="Select * From Dept Where Dept_Id='"&Dept_Id&"'"
      Call QuerySQL(SQL_Dept,RS_Dept)
      If Not RS_Dept.EOF Then
        If RS_Dept("Donate_Invoice")="N" And (Cstr(InvoiceType)=Cstr(InvoiceTypeY)) Then Exit Function
        Invoice_Pre=RS_Dept("Invoice_Pre")
        Invoice_Rule_Type=RS_Dept("Invoice_Rule_Type")
        Invoice_Rule_YMD=RS_Dept("Invoice_Rule_YMD")
        Invoice_Rule_Len=RS_Dept("Invoice_Rule_Len")
        Invoice_Rule_Pub=RS_Dept("Invoice_Rule_Pub")
        Invoice_Rule_Type2=RS_Dept("Invoice_Rule_Type2")
        Check_Invoice=True
      End If
      RS_Dept.Close
      Set RS_Dept=Nothing
    ElseIf Invoice_Kind="2" Then
      '物品捐贈取號規則
      SQL_Dept="Select * From Dept Where Dept_Id='"&Dept_Id&"'"
      Call QuerySQL(SQL_Dept,RS_Dept)
      If Not RS_Dept.EOF Then
        If RS_Dept("Donate_Invoice")="N" And (Cstr(InvoiceType)=Cstr(InvoiceTypeY)) Then Exit Function
        If RS_Dept("Invoice_Rule_Type2")="P" Then
          Invoice_Pre=RS_Dept("Invoice_Pre")
          Invoice_Rule_Type=RS_Dept("Invoice_Rule_Type")
          Invoice_Rule_YMD=RS_Dept("Invoice_Rule_YMD")
          Invoice_Rule_Len=RS_Dept("Invoice_Rule_Len")
          Invoice_Rule_Pub=RS_Dept("Invoice_Rule_Pub")
          Invoice_Rule_Type2=RS_Dept("Invoice_Rule_Type2")
        Else
          Invoice_Pre=RS_Dept("Invoice_Pre2")
          Invoice_Rule_Type=RS_Dept("Invoice_Rule_Type2")
          Invoice_Rule_YMD=RS_Dept("Invoice_Rule_YMD2")
          Invoice_Rule_Len=RS_Dept("Invoice_Rule_Len2")
          Invoice_Rule_Pub=RS_Dept("Invoice_Rule_Pub2")
          Invoice_Rule_Type2=RS_Dept("Invoice_Rule_Type2")
        End If
        Check_Invoice=True
      End If
      RS_Dept.Close
      Set RS_Dept=Nothing     
    ElseIf Invoice_Kind="3" Then
      '物品領用取號規則
      SQL_Dept="Select * From Dept Where Dept_Id='"&Dept_Id&"'"
      Call QuerySQL(SQL_Dept,RS_Dept)
      If Not RS_Dept.EOF Then
        Invoice_Pre=RS_Dept("Invoice_Pre3")
        Invoice_Rule_Type=RS_Dept("Invoice_Rule_Type3")
        Invoice_Rule_YMD=RS_Dept("Invoice_Rule_YMD3")
        Invoice_Rule_Len=RS_Dept("Invoice_Rule_Len3")
        Invoice_Rule_Pub=RS_Dept("Invoice_Rule_Pub3")
        Check_Invoice=True
      End If
      RS_Dept.Close
      Set RS_Dept=Nothing
    End If  
  End If

  '找到取號規則開始取號
  If Check_Invoice Then
    If Invoice_Kind="1" Then
      TableName="DONATE"
    ElseIf Invoice_Kind="2" Then
      TableName="CONTRIBUTE"
    ElseIf Invoice_Kind="3" Then
      TableName="CONTRIBUTE_ISSUE"
    End If
    InvoiceNo=""
    InvoiceNo2=""
    For IRL=1 To Invoice_Rule_Len
      InvoiceNo2=InvoiceNo2&"0"
    Next
    If Invoice_Rule_Type="R" Or Invoice_Rule_Type="A" Then
      Donate_Year=Year(Donate_Date)
      Donate_Month=Month(Donate_Date)
      Donate_Day=Day(Donate_Date)
      If Invoice_Rule_Type="R" Then
        Donate_Date_RA=(Donate_Year-1911)
      Else
        Donate_Date_RA=Donate_Year
      End If
      SQL_InvoiceNo="Select Top 1 Invoice_No,Invoice_Right=Right(Invoice_No,"&Invoice_Rule_Len&") From "&TableName&" Where Invoice_No<>'' And Invoice_Pre='"&Invoice_Pre&"' And Issue_Type_Keep<>'M' And (Issue_Type='' Or Issue_Type Is null Or Issue_Type='D') "
      If Invoice_Rule_Pub="N" Then SQL_InvoiceNo=SQL_InvoiceNo&"And Dept_Id='"&Dept_Id&"' "
      If Invoice_Rule_YMD="Y" Then
        SQL_InvoiceNo=SQL_InvoiceNo&"And Year(Donate_Date)="&Donate_Year&" "
        If Invoice_Rule_Type="R" Then
          InvoiceNo=Right("000"&Cstr(Donate_Date_RA),3)
        Else
          InvoiceNo=Cstr(Donate_Date_RA)
        End If
      ElseIf Invoice_Rule_YMD="M" Then
        SQL_InvoiceNo=SQL_InvoiceNo&"And Year(Donate_Date)="&Donate_Year&" And Month(Donate_Date)="&Donate_Month&" "
        If Invoice_Rule_Type="R" Then
          InvoiceNo=Right("000"&Cstr(Donate_Date_RA),3)&Right("00"&Cstr(Donate_Month),2)
        Else
          InvoiceNo=Cstr(Donate_Date_RA)&Right("00"&Cstr(Donate_Month),2)
        End If
      Else
        SQL_InvoiceNo=SQL_InvoiceNo&"And Year(Donate_Date)="&Donate_Year&" And Month(Donate_Date)="&Donate_Month&" And Day(Donate_Date)="&Donate_Day&" "
        If Invoice_Rule_Type="R" Then
          InvoiceNo=Right("000"&Cstr(Donate_Date_RA),3)&Right("00"&Cstr(Donate_Month),2)&Right("00"&Cstr(Donate_Day),2)
        Else
          InvoiceNo=Cstr(Donate_Date_RA)&Right("00"&Cstr(Donate_Month),2)&Right("00"&Cstr(Donate_Day),2)
        End If
      End If
      SQL_InvoiceNo=SQL_InvoiceNo&"Order By Invoice_Right Desc"
      If Invoice_Rule_Type2="P" And (Invoice_Kind="1" Or Invoice_Kind="2") Then
        If Invoice_Kind="1" Then
          SQLInvoiceNo=SQL_InvoiceNo
          Ary_SQL=Split(SQLInvoiceNo,"Order By Invoice_Right Desc")
          SQL_InvoiceNo=replace(Ary_SQL(0)&" Union "&replace(replace(Ary_SQL(0),"DONATE","CONTRIBUTE"),"Donate_Date","Contribute_Date")&" Order By Invoice_Right Desc ","Top 1","")
        ElseIf Invoice_Kind="2" Then
          SQLInvoiceNo=replace(SQL_InvoiceNo,"Donate_Date","Contribute_Date")
          Ary_SQL=Split(SQLInvoiceNo,"Order By Invoice_Right Desc")
          SQL_InvoiceNo=replace(Ary_SQL(0)&" Union "&replace(replace(Ary_SQL(0),"CONTRIBUTE","DONATE"),"Contribute_Date","Donate_Date")&" Order By Invoice_Right Desc ","Top 1","")
        End If
      Else
        If Invoice_Kind="2" Then SQL_InvoiceNo=replace(SQL_InvoiceNo,"Donate_Date","Contribute_Date")
        If Invoice_Kind="3" Then SQL_InvoiceNo=replace(replace(SQL_InvoiceNo,"Invoice","Issue"),"Donate_Date","Issue_Date")
      End If

      Call QuerySQL(SQL_InvoiceNo,RS_InvoiceNo)
      If RS_InvoiceNo.EOF Then
        Get_Invoice_No2=Invoice_Pre&"/"&InvoiceNo&Right(InvoiceNo2&CStr(CDbl(Right(InvoiceNo2,Invoice_Rule_Len))+1),Invoice_Rule_Len)
      Else
        If Invoice_Kind="3" Then
          Get_Invoice_No2=Invoice_Pre&"/"&InvoiceNo&Right(InvoiceNo2&CStr(CDbl(Right(RS_InvoiceNo("Issue_No"),Invoice_Rule_Len))+1),Invoice_Rule_Len)
        Else
          Get_Invoice_No2=Invoice_Pre&"/"&InvoiceNo&Right(InvoiceNo2&CStr(CDbl(Right(RS_InvoiceNo("Invoice_No"),Invoice_Rule_Len))+1),Invoice_Rule_Len)
        End If
      End If
      RS_InvoiceNo.Close
      Set RS_InvoiceNo=Nothing
    ElseIf Invoice_Rule_Type="S" Then
      SQL_InvoiceNo="Select Top 1 Invoice_No,Invoice_Right=Convert(numeric,Right(Invoice_No,"&Invoice_Rule_Len&")) From "&TableName&" Where Invoice_No<>'' And Invoice_Pre='"&Invoice_Pre&"' And Issue_Type_Keep<>'M' And (Issue_Type='' Or Issue_Type Is null Or Issue_Type='D') And Convert(numeric,Right(Invoice_No,"&Invoice_Rule_Len&"))>0 "
      If Invoice_Rule_Pub="N" Then SQL_InvoiceNo=SQL_InvoiceNo&"And Dept_Id='"&Dept_Id&"' "
      SQL_InvoiceNo=SQL_InvoiceNo&"Order By Convert(numeric,Right(Invoice_No,"&Invoice_Rule_Len&")) Desc"
      If Invoice_Rule_Type2="P" Then
        If Invoice_Kind="1" Then
          SQLInvoiceNo=SQL_InvoiceNo
          Ary_SQL=Split(SQLInvoiceNo,"Order By Invoice_Right Desc")
          SQL_InvoiceNo=Ary_SQL(0)&" Union "&replace(replace(Ary_SQL(0),"DONATE","CONTRIBUTE"),"Donate_Date","Contribute_Date")&" Order By Invoice_Right Desc "
        ElseIf Invoice_Kind="2" Then
          SQLInvoiceNo=SQL_InvoiceNo
          Ary_SQL=Split(SQLInvoiceNo,"Order By Invoice_Right Desc")
          SQL_InvoiceNo=Ary_SQL(0)&" Union "&replace(replace(Ary_SQL(0),"CONTRIBUTE","DONATE"),"Contribute_Date","Donate_Date")&" Order By Invoice_Right Desc "
        End If
      Else
        SQL_InvoiceNo=SQL_InvoiceNo&"Order By Convert(numeric,Right(Invoice_No,"&Invoice_Rule_Len&")) Desc"  
      End If
      If Invoice_Kind="3" Then SQL_InvoiceNo=replace(replace(SQL_InvoiceNo,"Invoice","Issue"),"Donate_Date","Issue_Date")
      Call QuerySQL(SQL_InvoiceNo,RS_InvoiceNo)
      If RS_InvoiceNo.EOF Then
        Get_Invoice_No2=Invoice_Pre&"/"&Right(InvoiceNo2&CStr(CDbl(Right(InvoiceNo2,Invoice_Rule_Len))+1),Invoice_Rule_Len)
      Else
        If Invoice_Kind="3" Then
          Get_Invoice_No2=Invoice_Pre&"/"&Right(InvoiceNo2&CStr(CDbl(Right(RS_InvoiceNo("Issue_No"),Invoice_Rule_Len))+1),Invoice_Rule_Len)
        Else
          Get_Invoice_No2=Invoice_Pre&"/"&Right(InvoiceNo2&CStr(CDbl(Right(RS_InvoiceNo("Invoice_No"),Invoice_Rule_Len))+1),Invoice_Rule_Len)
        End If
      End If
      RS_InvoiceNo.Close
      Set RS_InvoiceNo=Nothing
    ElseIf Invoice_Rule_Type="O" Then
      '收據編號客制化
      
    End If
  End If
End Function

Function Get_Close(Close_Kind,Dept_Id,Donate_Date,UserID)
  Get_Close=False
  SQL_Close="Select * From DEPT Where Dept_Id='"&Dept_Id&"'"
  Call QuerySQL(SQL_Close,RS_Close)
  If Not RS_Close.EOF Then
    If Close_Kind="1" Then
      Donate_Close=RS_Close("Donate_Close")
      Donate_Open_User=RS_Close("Donate_Open_User")
      Donate_Open_LastDate=RS_Close("Donate_Open_LastDate")
    ElseIf Close_Kind="2" Then
      Donate_Close=RS_Close("Contribute_Close")
      Donate_Open_User=RS_Close("Contribute_Open_User")
      Donate_Open_LastDate=RS_Close("Contribute_Open_LastDate")
    End If
    If Donate_Close<>"" Then
      If CDate(Donate_Date)<=CDate(Donate_Close) Then
      	If Donate_Open_LastDate<>"" And Donate_Open_User<>"" Then
      	  If CDate(Donate_Open_LastDate)=CDate(Date()) And Cstr(Donate_Open_User)=Cstr(UserID) Then Get_Close=True
        End If
      Else
        Get_Close=True
      End If
    Else
      Get_Close=True
    End If
  End If
  RS_Close.Close
  Set RS_Close=Nothing
End Function

Function syslog (log_type,log_desc)
  sys_date=date()
  client_IP=Request.ServerVariables("REMOTE_ADDR")
  sql="insert into sys_log values (getdate(),'"&sys_date&"','"&session("user_id")&"','"&session("user_name")&"','"&session("dept_id")&"','"&log_type&"','"&log_desc&"','"&client_IP&"')"
  On error resume next
  Set RS=Conn.Execute(SQL)
end Function

Function ChineseMoney( lmoney )
sDallor = cstr( lMoney)
chDallor = array("", "拾", "佰", "仟", "萬","拾","佰", "仟","億","拾","佰", "仟","兆")
chB = array("", "萬","億","兆")
for i= 1 to len(sDallor) step 1
        if mid(sDallor, i, 1)<>"0" then
                xx = xx & mid(sDallor, i, 1) & chDallor(len(sDallor)-i) '&"萬"
        else
                if right(xx, 1)<>"零" then
                        If ((len(sDallor)-i) mod 4) = 0 then
                        xx = xx & chB( ( len(sDallor) - i) \ 4 )
                        else 
                        xx = xx & "零"
                        End If
                else
                        If ((len(sDallor)-i) mod 4) = 0 then
                        'msgbox "i = " & i & " xx=" & xx
                        xx= mid(xx, 1, len(xx)-1) & chB( ( len(sDallor) - i) \ 4 )
                        End If
                end if
        end if
next

        if right(xx, 1)="零" then xx = mid(xx,1,len(xx)-1)
        xx = replace( xx, "1", "壹")
        xx = replace( xx, "2", "貳")
        xx = replace( xx, "3", "參")
        xx = replace( xx, "4", "肆")
        xx = replace( xx, "5", "伍")
        xx = replace( xx, "6", "陸")
        xx = replace( xx, "7", "柒")
        xx = replace( xx, "8", "捌")
        xx = replace( xx, "9", "玖")

        ChineseMoney= xx & "元整"
End Function

Function LinkGrid (SQL,HLink,LinkParam,LinkTarget,LinkType)
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL,Conn,1,1
    FieldsCount = RS1.Fields.Count-1
    Dim i
    Response.Write "<table id=grid border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    For i = 1 To FieldsCount
	  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i).Name & "</span></font></td>"
    Next
    Response.Write "</tr>"
    While Not RS1.EOF
      if Hlink<>"" then
         if LinkType="window" then
            Response.Write "<a href onclick=""window.open('" & Hlink & RS1(LinkParam)&"','','scrollbars=no,top=100,left=120,width=470,height=400')"">"
         else
            Response.Write "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>"
         end if   
         showhand="style='cursor:hand'"
      end if
	  Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
	  For i = 1 To FieldsCount
           Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i) & "</span></td>"
      Next    
      RS1.MoveNext
      Response.Write "</tr>"
      Response.Write "</a>"
    Wend
    Response.Write "</table>"
    RS1.Close
End Function

Function Get_Shopping_OrderNo(Od_Sob_Type)
  Set od_sobRec = Server.CreateObject("ADODB.RecordSet")
  YYYY = Cstr(Year(Now()))
  MM = Cstr(Month(Now()))
  DD = Cstr(Day(Now()))
  If Len(MM) = 1 Then MM = "0"&MM
  If Len(DD) = 1 Then DD = "0"&DD
  Uniqueod_sob = "False"
  While Uniqueod_sob = "False"
    Get_Shopping_OrderNo = YYYY & MM & DD
    For I = 1 to 4
      Randomize
      Get_Shopping_OrderNo = Get_Shopping_OrderNo & Chr(int(26*Rnd)+65)
    Next
    SQL_Sob="Select * From DONATE_OD_SOB Where od_sob='"&Get_Shopping_OrderNo&"'"
    Set RS_Sob = Server.CreateObject("ADODB.RecordSet")
    RS_Sob.Open SQL_Sob,Conn,1,1
    If RS_Sob.EOF Then 
      Uniqueod_sob = "True"
      SQL1_Sob="DONATE_OD_SOB"
      Set RS1_Sob = Server.CreateObject("ADODB.RecordSet")
      RS1_Sob.Open SQL1_Sob,Conn,1,3
      RS1_Sob.Addnew
      RS1_Sob("Od_Sob")=Get_Shopping_OrderNo
      RS1_Sob("Od_Sob_Type")=Od_Sob_Type
      RS1_Sob("Create_Date")=Date()
      RS1_Sob("Create_DateTime")=Now()
      RS1_Sob("Create_User")=session("user_name")
      RS1_Sob("Create_IP")=Request.ServerVariables("REMOTE_HOST")
      RS1_Sob.Update
      RS1_Sob.Close
      Set RS1_Sob=Nothing
    End If  
    RS_Sob.Close
    Set RS_Sob = Nothing
  Wend
End Function

Function Declare_DonorId (Donor_Id)
  SQL="DECLARE @Begin_DonateDate datetime " & _
      "DECLARE @Last_DonateDate datetime " & _
      "DECLARE @Donate_No numeric " & _
      "DECLARE @Donate_Total numeric " & _
      "Select Top 1 @Begin_DonateDate=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_Amt>0 Order By Donate_Date " & _
      "Select Top 1 @Last_DonateDate=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_Amt>0 Order By Donate_Date Desc " & _    
      "Select @Donate_No=Count(*) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_Amt>0 " & _
      "Select @Donate_Total=IsNull(Sum(Donate_Amt),0) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_Amt>0 " & _
      "Update DONOR Set Begin_DonateDate=@Begin_DonateDate,Last_DonateDate=@Last_DonateDate,Donate_No=@Donate_No,Donate_Total=@Donate_Total Where Donor_Id='"&Donor_Id&"'"
  Set RS=Conn.Execute(SQL)

  SQL="DECLARE @Begin_DonateDateD datetime " & _
      "DECLARE @Last_DonateDateD datetime " & _
      "DECLARE @Donate_NoD numeric " & _
      "DECLARE @Donate_TotalD numeric " & _
      "Select Top 1 @Begin_DonateDateD=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtD>0 Order By Donate_Date " & _
      "Select Top 1 @Last_DonateDateD=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtD>0 Order By Donate_Date Desc " & _    
      "Select @Donate_NoD=Count(*) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtD>0 " & _
      "Select @Donate_TotalD=IsNull(Sum(Donate_AmtD),0) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtD>0 " & _
      "Update DONOR Set Begin_DonateDateD=@Begin_DonateDateD,Last_DonateDateD=@Last_DonateDateD,Donate_NoD=@Donate_NoD,Donate_TotalD=@Donate_TotalD Where Donor_Id='"&Donor_Id&"'"
  Set RS=Conn.Execute(SQL)

  SQL="DECLARE @Last_PledgeDate datetime " & _
      "Select Top 1 @Last_PledgeDate=Donate_ToDate From PLEDGE Where Donor_Id='"&Donor_Id&"' Order By Donate_ToDate Desc " & _
      "Update DONOR Set Last_PledgeDate=@Last_PledgeDate Where Donor_Id='"&Donor_Id&"'"
  Set RS=Conn.Execute(SQL)
  
  SQL="DECLARE @Begin_DonateDateC datetime " & _
      "DECLARE @Last_DonateDateC datetime " & _
      "DECLARE @Donate_NoC numeric " & _
      "DECLARE @Donate_TotalC numeric " & _
      "Select Top 1 @Begin_DonateDateC=Contribute_Date From CONTRIBUTE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' Order By Contribute_Date " & _
      "Select Top 1 @Last_DonateDateC=Contribute_Date From CONTRIBUTE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' Order By Contribute_Date Desc " & _
      "Select @Donate_NoC=Count(*) From CONTRIBUTE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' " & _
      "Select @Donate_TotalC=Isnull(Sum(Contribute_Amt),0) From CONTRIBUTE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' " & _
      "Update DONOR Set Begin_DonateDateC=@Begin_DonateDateC,Last_DonateDateC=@Last_DonateDateC,Donate_NoC=@Donate_NoC,Donate_TotalC=@Donate_TotalC Where Donor_Id='"&Donor_Id&"'"
  Set RS=Conn.Execute(SQL)

  SQL="DECLARE @Begin_DonateDateM datetime " & _
      "DECLARE @Last_DonateDateM datetime " & _
      "DECLARE @Donate_NoM numeric " & _
      "DECLARE @Donate_TotalM numeric " & _
      "Select Top 1 @Begin_DonateDateM=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtM>0 Order By Donate_Date " & _
      "Select Top 1 @Last_DonateDateM=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtM>0 Order By Donate_Date Desc " & _    
      "Select @Donate_NoM=Count(*) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtM>0 " & _
      "Select @Donate_TotalM=IsNull(Sum(Donate_AmtM),0) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtM>0 " & _
      "Update DONOR Set Begin_DonateDateM=@Begin_DonateDateM,Last_DonateDateM=@Last_DonateDateM,Donate_NoM=@Donate_NoM,Donate_TotalM=@Donate_TotalM Where Donor_Id='"&Donor_Id&"'"
  Set RS=Conn.Execute(SQL)

  SQL="DECLARE @Begin_DonateDateA datetime " & _
      "DECLARE @Last_DonateDateA datetime " & _
      "DECLARE @Donate_NoA numeric " & _
      "DECLARE @Donate_TotalA numeric " & _
      "Select Top 1 @Begin_DonateDateA=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtA>0 Order By Donate_Date " & _
      "Select Top 1 @Last_DonateDateA=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtA>0 Order By Donate_Date Desc " & _    
      "Select @Donate_NoA=Count(*) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtA>0 " & _
      "Select @Donate_TotalA=IsNull(Sum(Donate_AmtA),0) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtA>0 " & _
      "Update DONOR Set Begin_DonateDateA=@Begin_DonateDateA,Last_DonateDateA=@Last_DonateDateA,Donate_NoA=@Donate_NoA,Donate_TotalA=@Donate_TotalA Where Donor_Id='"&Donor_Id&"'"
  Set RS=Conn.Execute(SQL)

  SQL="DECLARE @Begin_DonateDateS datetime " & _
      "DECLARE @Last_DonateDateS datetime " & _
      "DECLARE @Donate_NoS numeric " & _
      "DECLARE @Donate_TotalS numeric " & _
      "Select Top 1 @Begin_DonateDateS=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtS>0 Order By Donate_Date " & _
      "Select Top 1 @Last_DonateDateS=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtS>0 Order By Donate_Date Desc " & _    
      "Select @Donate_NoS=Count(*) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtS>0 " & _
      "Select @Donate_TotalS=IsNull(Sum(Donate_AmtS),0) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And Donate_AmtS>0 " & _
      "Update DONOR Set Begin_DonateDateS=@Begin_DonateDateS,Last_DonateDateS=@Last_DonateDateS,Donate_NoS=@Donate_NoS,Donate_TotalS=@Donate_TotalS Where Donor_Id='"&Donor_Id&"'"
  Set RS=Conn.Execute(SQL)

  SQL="DECLARE @Begin_DonateDateND datetime " & _
      "DECLARE @Last_DonateDateND datetime " & _
      "DECLARE @Donate_NoND numeric " & _
      "DECLARE @Donate_TotalND numeric " & _
      "Select Top 1 @Begin_DonateDateND=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And (Donate_AmtM>0 Or Donate_AmtA>0 Or Donate_AmtS>0) Order By Donate_Date " & _
      "Select Top 1 @Last_DonateDateND=Donate_Date From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And (Donate_AmtM>0 Or Donate_AmtA>0 Or Donate_AmtS>0) Order By Donate_Date Desc " & _    
      "Select @Donate_NoND=Count(*) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And (Donate_AmtM>0 Or Donate_AmtA>0 Or Donate_AmtS>0) " & _
      "Select @Donate_TotalND=IsNull(Sum(Donate_AmtM)+Sum(Donate_AmtA)+Sum(Donate_AmtS),0) From DONATE Where Donor_Id='"&Donor_Id&"' And Issue_Type<>'D' And (Donate_AmtM>0 Or Donate_AmtA>0 Or Donate_AmtS>0) " & _
      "Update DONOR Set Begin_DonateDateND=@Begin_DonateDateND,Last_DonateDateND=@Last_DonateDateND,Donate_NoND=@Donate_NoND,Donate_TotalND=@Donate_TotalND Where Donor_Id='"&Donor_Id&"'"
  Set RS=Conn.Execute(SQL)
End Function

Function ProgDesc(Prog_Id)
  ProgDesc=""
  SQL_Prog="Select Prog_Desc From PROG Where Prog_Id='"&Prog_Id&"'"
  Set RS_Prog = Server.CreateObject("ADODB.RecordSet")
  RS_Prog.Open SQL_Prog,Conn,1,1
  If Not RS_Prog.EOF Then ProgDesc=RS_Prog("Prog_Desc")
  RS_Prog.Close
  Set RS_Prog=Nothing
End Function

Function Data_Plus(Transfer_Data)
  Data_Plus=""
  If Transfer_Data<>"" Then Data_Plus=replace(Transfer_Data,"'","''")
End Function

Function Data_Minus(Transfer_Data)
  Data_Minus=""
  If Transfer_Data<>"" Then Data_Minus=replace(Transfer_Data,"''","'")
End Function

Function Pledge_Transfer(Pledge_Id,Donate_FromDate,Donate_ToDate,Donate_Period)
  SQL_Transfer="Delete From PLEDGE_TRANSFER Where Pledge_Id='"&Pledge_Id&"' And Transfer_Status='待授權'"
  Set RS_Transfer=Conn.Execute(SQL_Transfer)
  Transfer_Date=Donate_FromDate
  While CDate(Transfer_Date)<CDate(Donate_ToDate)    
    SQL_Transfer="Select * From PLEDGE_TRANSFER Where Pledge_Id='"&Pledge_Id&"' And Transfer_Status='待授權' And Pledge_Transfer_Year='"&Year(Transfer_Date)&"' And Pledge_Transfer_Month='"&Month(Transfer_Date)&"'"
    response.write SQL_Transfer&"<br >"
    Set RS_Transfer = Server.CreateObject("ADODB.RecordSet")
    RS_Transfer.Open SQL_Transfer,Conn,1,3
    If RS_Transfer.EOF Then RS_Transfer.Addnew
    RS_Transfer("Pledge_Id")=Pledge_Id
    RS_Transfer("Pledge_Transfer_Year")=Year(Transfer_Date)
    RS_Transfer("Pledge_Transfer_Month")=Month(Transfer_Date)
    RS_Transfer("Transfer_Status")="待授權"
    RS_Transfer("od_sob")=""
    RS_Transfer.Update
    RS_Transfer.Close
    Set RS_Transfer=Nothing
    If Donate_Period="0" Then
      Transfer_Date=Donate_ToDate
    Else
      Transfer_Date=DateAdd("M",Cdbl(Donate_Period),Transfer_Date)
    End If
  Wend
End Function

'20130725 Add by GoodTV Tanya:增加性別與稱謂連動
Function CodeSex (SQL,FName,Listfield,BoundColumn,menusize)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  Response.Write "<SELECT Name='" & FName & "' size='" & menusize & "' style='font-size: 9pt; font-family: 新細明體' onchange='JavaScript:ChgSex(this.value);'>"  
  'If RS1.EOF Then
  '  Response.Write "<OPTION>" & " " & "</OPTION>"
  'End If
  'If BoundColumn="" or IsNull(BoundColumn) Then
  '  Response.Write "<OPTION selected value=''>" & " " & "</OPTION>"
  'End If    
  Response.Write "<OPTION>" & " " & "</OPTION>"
  While Not RS1.EOF
    If BoundColumn<>"" Then
      If Cstr(RS1(FName))=Cstr(BoundColumn) Then
        strselected = "selected"
      Else
        strselected = ""
      End If
    Else
      strselected = ""
    End If
    Response.Write "<OPTION " & strselected & " value='" & RS1(FName) & "'>" & RS1(Listfield) & "</OPTION>"
    RS1.MoveNext
  Wend
  Response.Write "</SELECT>"
  RS1.Close
  Set RS1=Nothing
  '性別/稱謂選單連動
  Response.Write "<script language=""JavaScript""><!--" & Chr(13) & Chr(10)
  Response.Write "  function ChgSex(SexValue){" & Chr(13) & Chr(10)
  'Response.Write "    alert(SexValue);" & Chr(13) & Chr(10)
  Response.Write "    if (SexValue=='男'){" & Chr(13) & Chr(10)
  Response.Write "    		document.form.Title.value='先生'" & Chr(13) & Chr(10)
  Response.Write "    }else if (SexValue=='女'){" & Chr(13) & Chr(10)
  Response.Write "    		document.form.Title.value='小姐'" & Chr(13) & Chr(10)
  Response.Write "    }else if (SexValue=='歿'){" & Chr(13) & Chr(10)
  Response.Write "    		document.form.Title.value=''" & Chr(13) & Chr(10)
  Response.Write "    }else{" & Chr(13) & Chr(10)
  Response.Write "    		document.form.Title.value='先生/小姐/寶號'" & Chr(13) & Chr(10)
  Response.Write "    }" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "--></script>" & Chr(13) & Chr(10)
  
End Function
%>