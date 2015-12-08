<%
'頁尾機構連絡資訊
SQL1="Select *,Dept_Address=Zip_Code+A.mValue+B.mValue+Address,Dept_Address2=Zip_Code2+C.mValue+D.mValue+Address2 From Dept " & _
     "Left Join CodeCity As A On Dept.City_Code=A.mCode " & _
     "Left Join CodeCity As B On Dept.Area_Code=B.mCode " & _
     "Left Join CodeCity As C On Dept.City_Code2=C.mCode " & _
     "Left Join CodeCity As D On Dept.Area_Code2=D.mCode "
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then
  Comp_Name=RS1("Comp_Name")
  Comp_Slogan=RS1("Comp_Slogan")
	Web_Name="http://"&Request.ServerVariables("SERVER_NAME")
	Dept_Address=RS1("Dept_Address")
	Dept_Address2=RS1("Dept_Address2")
	Dept_Tel=RS1("TEL")
	Dept_Fax=RS1("Fax")
	Account=RS1("Account")
	Dept_Licence=RS1("Licence")
	Dept_Uniform_No=RS1("Uniform_No")
	Dept_EMail=RS1("EMail")
	Meta_Foot=RS1("Meta_Foot")
End If
RS1.Close
Set RS1=Nothing

'捐款參數
SQL1="Select * From DONATE_SYSTEM Order By System_Key"
Call QuerySQL(SQL1,RS1)
While Not RS1.EOF
  If RS1("System_Key")="Card_Code"          Then Card_Code=RS1("System_Value")
  If RS1("System_Key")="Card_PayName"       Then Card_PayName=RS1("System_Value")
  If RS1("System_Key")="Card_Verify"        Then Card_Verify=RS1("System_Value")
  If RS1("System_Key")="Check_ID"           Then Check_ID=RS1("System_Value")
  If RS1("System_Key")="ECBank_Barcode"     Then ECBank_Barcode=RS1("System_Value")
  If RS1("System_Key")="ECBank_BarcodeDate" Then ECBank_BarcodeDate=RS1("System_Value")
  If RS1("System_Key")="ECBank_Code"        Then ECBank_Code=RS1("System_Value")
  If RS1("System_Key")="ECBank_FamiPort"    Then ECBank_FamiPort=RS1("System_Value")
  If RS1("System_Key")="ECBank_Ibon"        Then ECBank_Ibon=RS1("System_Value")
  If RS1("System_Key")="ECBank_Key"         Then ECBank_Key=RS1("System_Value")
  If RS1("System_Key")="ECBank_LiftET"      Then ECBank_LiftET=RS1("System_Value")
  If RS1("System_Key")="ECBank_OKGo"        Then ECBank_OKGo=RS1("System_Value")
  If RS1("System_Key")="ECBank_OKGo"        Then ECBank_OKGo=RS1("System_Value")
  If RS1("System_Key")="ECBank_Vacc"        Then ECBank_Vacc=RS1("System_Value")
  If RS1("System_Key")="ECBank_WebAtm"      Then ECBank_WebAtm=RS1("System_Value")
  If RS1("System_Key")="SendMailDefault"    Then SendMailDefault=RS1("System_Value")
  If RS1("System_Key")="SendMailType"       Then SendMailType=RS1("System_Value")
  RS1.MoveNext
Wend
RS1.Close
Set RS1=Nothing

'轉帳授權書
Download_FileURL="/upload/定額捐款授權4合一表.doc"
SQL1="Select TOP 1 * From UPLOAD Where Object_ID='0' And Ap_Name='pledge' And Attach_Type='doc' Order By Ser_No Desc"
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
If Not RS1.EOF Then Download_FileURL="/upload/"&RS1("Upload_FileURL")
RS1.Close
Set RS1=Nothing
    
'本頁文件名
UrlName=Mid(Request.ServerVariables("URL"),InstrRev(Request.ServerVariables("URL"),"/")+1)
    
If Session("Hit_Index")="" Then
  SQL1="Select * From HIT Where Hit_Nemu='Index' And Hit_Year='"&Year(Date())&"' And Hit_Month='"&Month(Date())&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  If RS1.EOF Then
    SQL2="HIT"
    Set RS2 = Server.CreateObject("ADODB.RecordSet")
    RS2.Open SQL2,Conn,1,3
    RS2.Addnew
    RS2("Hit_Nemu")="index"
    RS2("Hit_Object_ID")="0"
    RS2("Hit_Page")="0"
    RS2("Hit_Year")=Year(Date())
    RS2("Hit_Month")=Month(Date())
    RS2("Hit_Count")="1"
    RS2("Last_Date")=date()
    RS2("Last_DateTime")=now()
    RS2("Last_AddIp")=Request.ServerVariables("REMOTE_HOST")
    RS2.Update
    RS2.Close
    Set RS2=Nothing
  Else
    Hit_Count=Cdbl(RS1("Hit_Count"))+1
    RS1("Hit_Count")=Hit_Count
    RS1("Last_Date")=date()
    RS1("Last_DateTime")=now()
    RS1("Last_AddIp")=Request.ServerVariables("REMOTE_HOST")
    RS1.Update
  End If
  RS1.Close
  Set RS1=Nothing
  Session("Hit_Index")="Welome"
  
  SQL1="Update DEPT_HIT Set HitNo=HitNo+1"
  Set RS1=Conn.Execute(SQL1)  
End If

'訪客人數
'If UrlName="index.asp" Then
  SQL1="Select * From DEPT_HIT"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then Dept_Hit=RS1("HitNo")
  RS1.Close
  Set RS1=Nothing
'End If

Sub Message()
  If session("errnumber")=1 Then
    Response.Write "<script language='javascript'>alert('"&session("msg")&"')</script>"
  End If
  session("msg")=""
  session("errnumber")=0
End Sub

Function QuerySQL (SQL,RS)
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,conn,1,1
End Function

Function InsertSQL (SQL,RS)
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,conn,1,3
  RS.Addnew
End Function

Function UpdateSQL (SQL,RS)
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,conn,1,3
End Function

Function HIT(Nemu,Object_ID,Page)
  SQL_HIT="Select * From HIT Where Hit_Nemu='"&Nemu&"' And Hit_Object_ID='"&Object_ID&"' And Hit_Page='"&Page&"' And Hit_Year='"&Year(Date())&"' And Hit_Month='"&Month(Date())&"'"
  Set RS_HIT = Server.CreateObject("ADODB.RecordSet")
  RS_HIT.Open SQL_HIT,Conn,1,3
  If RS_HIT.EOF Then
    SQL2_HIT="HIT"
    Set RS2_HIT = Server.CreateObject("ADODB.RecordSet")
    RS2_HIT.Open SQL2_HIT,Conn,1,3
    RS2_HIT.Addnew
    RS2_HIT("Hit_Nemu")=Nemu
    RS2_HIT("Hit_Object_ID")=Object_ID
    RS2_HIT("Hit_Page")=Page
    RS2_HIT("Hit_Year")=Year(Date())
    RS2_HIT("Hit_Month")=Month(Date())
    RS2_HIT("Hit_Count")="1"
    RS2_HIT("Last_Date")=date()
    RS2_HIT("Last_DateTime")=now()
    RS2_HIT("Last_AddIp")=Request.ServerVariables("REMOTE_HOST")
    RS2_HIT.Update
    RS2_HIT.Close
    Set RS2_HIT=Nothing
  Else
    'Last_Date=RS_HIT("Last_Date")
    'Last_AddIp=RS_HIT("Last_AddIp")
    'If Cstr(Last_Date)<>Cstr(date()) Or Cstr(Last_AddIp)<>Cstr(Request.ServerVariables("REMOTE_HOST")) Then
      Hit_Count=Cdbl(RS_HIT("Hit_Count"))+1
      RS_HIT("Hit_Count")=Hit_Count
      RS_HIT("Last_Date")=date()
      RS_HIT("Last_DateTime")=now()
      RS_HIT("Last_AddIp")=Request.ServerVariables("REMOTE_HOST")
      RS_HIT.Update
    'End If
  End If
  RS_HIT.Close
  Set RS_HIT=Nothing
End Function

Function CheckKeyWordJ(Form,FName,FSpace,Ftype,FLen,ListField)
  If FSpace="Y" Then
    Response.Write "if(document." & Form & "." & FName & ".value==''){" & Chr(13) & Chr(10)
    Response.Write "  alert('"& ListField & "  欄位不可為空白！');" & Chr(13) & Chr(10)
    Response.Write "  document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
    Response.Write "  return;" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)  
  End If
  If Ftype="email" Then
    Response.Write "if(document." & Form & "." & FName & ".value!=''){" & Chr(13) & Chr(10)
    Response.Write "  if(document." & Form & "." & FName & ".value.indexOf(""@"")==-1||document." & Form & "." & FName & ".value.indexOf(""."")==-1||document." & Form & "." & FName & ".value.indexOf(""@"")==1||document." & Form & "." & FName & ".value.indexOf(""."")==1||document." & Form & "." & FName & ".value.indexOf(""@"")==document." & Form & "." & FName & ".value.length||document." & Form & "." & FName & ".value.indexOf(""."")==document." & Form & "." & FName & ".value.length){" & Chr(13) & Chr(10)
    Response.Write "    alert('"& ListField & "  欄位格式錯誤！');" & Chr(13) & Chr(10)
    Response.Write "    document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
    Response.Write "    return;" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "  if(!(/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(document." & Form & "." & FName & ".value))){" & Chr(13) & Chr(10)
    Response.Write "    alert('"& ListField & "  欄位格式錯誤！');" & Chr(13) & Chr(10)
    Response.Write "    document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
    Response.Write "    return;" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)
  ElseIf Ftype="number" Then
    Response.Write "if(document." & Form & "." & FName & ".value!=''){" & Chr(13) & Chr(10)
    Response.Write "  if(isNaN(Number(document." & Form & "." & FName & ".value))==true){" & Chr(13) & Chr(10)
    Response.Write "    alert('"& ListField & "  欄位必須為數字！');" & Chr(13) & Chr(10)
    Response.Write "    document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
    Response.Write "    return;" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)
  ElseIf Ftype="url" Then
    Response.Write "if(document." & Form & "." & FName & ".value!=''){" & Chr(13) & Chr(10)
    Response.Write "  if(document." & Form & "." & FName & ".value.indexOf(""."")==-1||document." & Form & "." & FName & ".value.indexOf(""."")==1||document." & Form & "." & FName & ".value.indexOf(""."")==document." & Form & "." & FName & ".value.length){" & Chr(13) & Chr(10)
    Response.Write "    alert('"& ListField & "  欄位格式錯誤！');" & Chr(13) & Chr(10)
    Response.Write "    document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
    Response.Write "    return;" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)  
  End If
  If Cint(FLen)>0 Then
    Response.Write "if(document." & Form & "." & FName & ".value!=''){" & Chr(13) & Chr(10)
    Response.Write "var cnt=0;" & Chr(13) & Chr(10)
    Response.Write "var sName=document." & Form & "." & FName & ".value;" & Chr(13) & Chr(10)
    Response.Write "for(var i=0;i<sName.length;i++ ){" & Chr(13) & Chr(10)
    Response.Write "  if(escape(sName.charAt(i)).length >= 4) cnt+=2;" & Chr(13) & Chr(10)
    Response.Write "  else cnt++;" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)
    Response.Write "if(cnt>"&FLen&"){" & Chr(13) & Chr(10)
    Response.Write "  alert('"& ListField & "  欄位長度超過限制！');" & Chr(13) & Chr(10)
    Response.Write "  document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
    Response.Write "  return;" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)
  End If
  Response.Write "if(document." & Form & "." & FName & ".value!=''){" & Chr(13) & Chr(10)
  Response.Write "  var keyword=document." & Form & "." & FName & ".value.toLowerCase();" & Chr(13) & Chr(10)
  If qs_type="content" Then
    Response.Write "  var AryKey = new Array('1=1','a=a','--','./');" & Chr(13) & Chr(10)
  ElseIf Ftype="tel" Then 
    Response.Write "  var AryKey = new Array('`','~','!','$','%','^','&','*','+','=','[',']','{','}','\\','|',';',':','\'','""','<','>','//','1=1','a=a','--',';','./');" & Chr(13) & Chr(10)
  ElseIf qs_type="url" Then
    Response.Write "  var AryKey = new Array('`','~','!','#','$','%','^','*','(',')','+','[',']','{','}','\\','|',';',':','\'','""','<','>','//','1=1','a=a','--',';','./');" & Chr(13) & Chr(10)
  ElseIf qs_type="money" Then
    Response.Write "  var AryKey = new Array('`','~','!','#','%','^','&','*','(',')','+','=','[',']','{','}','\\','|',';',':','\'','""','<','>','//','1=1','a=a','--',';','./');" & Chr(13) & Chr(10)
  Else  
    Response.Write "  var AryKey = new Array('`','~','!','#','$','%','^','&','*','(',')','+','=','[',']','{','}','\\','|',';',':','\'','""','<','>','//','1=1','a=a','--',';','./');" & Chr(13) & Chr(10)
  End If
  Response.Write "  for(var i=0;i<=AryKey.length-1;i++){" & Chr(13) & Chr(10)
  Response.Write "    if(keyword.indexOf(AryKey[i])!=-1){" & Chr(13) & Chr(10)
  Response.Write "      alert('"& ListField & "  欄位請勿使用特殊字元！');" & Chr(13) & Chr(10)
  Response.Write "      document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "      return;" & Chr(13) & Chr(10)
  Response.Write "    }" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "  var AryKey = new Array('.asp','.bak','.cfm','.css','.dos','.htm','.inc','.ini','.js','.php','.txt','/*','*/','.@','@.','@@','@C','@T','@P','a href','admin','acunetix','alert','and','application','begin','cast','cache','config','char','chr','cookie','count','create','css','cursor','deallocate','declare','delete','dir','drop','echo','end','eval','exec','exists','execute','fetch','from','hidden','iframe','insert','into','is_','join','js','kill','left','manage','master','mid','nchar','ntext','nvarchar','open','password','right','script','set','select','session','src','sys','sysobjects','syscolumns','table','text','truncate','url','user','update','varchar','where','while','xtype','%2527','address.tst');" & Chr(13) & Chr(10)
  Response.Write "  for(var i=0;i<=AryKey.length-1;i++){" & Chr(13) & Chr(10)
  Response.Write "    if(keyword.indexOf(AryKey[i])!=-1){" & Chr(13) & Chr(10)
  Response.Write "      alert('"& ListField & "  欄位請勿使用保留字元！');" & Chr(13) & Chr(10)
  Response.Write "      document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "      return;" & Chr(13) & Chr(10)
  Response.Write "    }" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function CheckExtJ (Form,FName,FExt,ListField)
  Response.Write "if(document." & Form & "." & FName & ".value!=''){" & Chr(13) & Chr(10)
  Response.Write "  if(document." & Form & "." & FName & ".value.length>4){" & Chr(13) & Chr(10)
  Response.Write "    if(document." & Form & "." & FName & ".value.indexOf('.')>-1){" & Chr(13) & Chr(10)
  Response.Write "      var ext_check=false;" & Chr(13) & Chr(10) 
  Response.Write "      var filename=document." & Form & "." & FName & ".value.toLowerCase();" & Chr(13) & Chr(10)
  Response.Write "      var ext=filename.substr(filename.lastIndexOf('.')+1,filename.length);" & Chr(13) & Chr(10)
  If FExt="img" Then
    Response.Write "      var StrExt = 'jpg,gif,bmp';" & Chr(13) & Chr(10)
  ElseIf FExt="flash" Then
    Response.Write "      var StrExt = 'swf';" & Chr(13) & Chr(10)  
  ElseIf FExt="doc" Then
    Response.Write "      var StrExt = 'doc,pdf,xls';" & Chr(13) & Chr(10)
  ElseIf FExt="wmv" Then
    Response.Write "      var StrExt = 'wmv,asf,avi,mpg';" & Chr(13) & Chr(10)
  ElseIf FExt="flv" Then
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

Function OptionList (SQL,FName,Listfield,BoundColumn,menusize)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  Response.Write "<SELECT Name='" & FName & "' size='" & menusize & "' style='font-size: 9pt; font-family: 新細明體'>"
  If RS1.EOF Then
    Response.Write "<OPTION>" & " " & "</OPTION>"
  End If
  If BoundColumn="" or IsNull(BoundColumn) Then
    Response.Write "<OPTION selected value=''>" & " " & "</OPTION>"
  End If    
  While Not RS1.EOF
    If RS1(FName)=BoundColumn Then
      stRS1elected = "selected"
    Else
      stRS1elected = ""
    End If      
    Response.Write "<OPTION " & stRS1elected & " value='" & RS1(FName) & "'>" & RS1(Listfield) & "</OPTION>"
    RS1.MoveNext
  WEnd
  Response.Write "</SELECT>"
  RS1.Close
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

Function OptionList_FromClass (SQL,FName,Listfield,BoundColumn,menusize,FromClass)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  Response.Write "<SELECT Name='" & FName & "' size='" & menusize & "' class='"&FromClass&"'>"
  If RS1.EOF Then
    Response.Write "<OPTION>" & " " & "</OPTION>"
  End If
  If BoundColumn="" or IsNull(BoundColumn) Then
    Response.Write "<OPTION selected value=''>" & " " & "</OPTION>"
  End If    
  While Not RS1.EOF
    If RS1(FName)=BoundColumn Then
      stRS1elected = "selected"
    Else
      stRS1elected = ""
    End If      
    Response.Write "<OPTION " & stRS1elected & " value='" & RS1(FName) & "'>" & RS1(Listfield) & "</OPTION>"
    RS1.MoveNext
  WEnd
  Response.Write "</SELECT>"
  RS1.Close
End Function

Function RadioBoxList (SQL,FName,Listfield,BoundColumn,NoChecked,Other)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  Dim i
  I = 1
  While Not RS1.EOF
    If BoundColumn<>"" Then
      If RS1(FName)=BoundColumn Then
        StrChecked="checked"
      Else
        StrChecked=""
      End If
    Else
      If Cstr(I) = Cstr(NoChecked) Then
        StrChecked="checked"
      Else
        StrChecked=""
      End If  
    End If   	
    If I>1 Then Response.Write "&nbsp;&nbsp;&nbsp;"
    Response.Write "<input type='radio' " & StrChecked & " name='" & FName &"' id='" & FName & I & "' value='" & RS1(FName) & "'>&nbsp;"&RS1(Listfield)
    If RS1(FName)="其他" Then Response.Write "&nbsp;<input name='" & FName &"_Other' type='text' size='10' maxlength='20' value='" & Other & "' />" 	
    I = I + 1
    RS1.MoveNext
  WEnd
  RS1.Close
End Function

Function RadioBoxList2 (SQL,FName,Listfield,BoundColumn,NoChecked,Br,Other)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  Dim i
  I = 1
  While Not RS1.EOF
    If BoundColumn="" Then
      If Cstr(I) = Cstr(NoChecked) Then
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
    If I>1 Then Response.Write "&nbsp;"
    If Cstr(Br)<>"" Then
      If Instr(Cstr(Br),",")>0 Then
        Ary_Br = Split(Cstr(Br),",")
        For J = 0 To UBound(Ary_Br)
          If Cstr(I)=Cstr(Ary_Br(J)) Then Response.Write "<br>"
        Next
      Else
         If Cstr(I)=Cstr(Br) Then Response.Write "<br>"
      End If  
    End If
    Response.Write "<input type='radio' " & StrChecked & " name='" & FName &"' id='" & FName & I & "' value='" & RS1(FName) & "'>"&RS1(Listfield)
    If RS1(FName)="其他" Or RS1(FName)="其它" Then Response.Write "&nbsp;<input name='" & FName &"_Other' type='text' size='10' maxlength='20' value='" & Other & "' />" 	
    I = I + 1
    RS1.MoveNext
  WEnd
  RS1.Close
End Function

Function CheckBoxList (SQL,FName,Listfield,BoundColumn)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  session("FieldsetCount")=RS1.RecordCount
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
End Function

Function CheckBoxList2 (SQL,FName,Listfield,BoundColumn,Br,Other)
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
    If Cstr(Br)<>"" Then
      If Instr(Cstr(Br),",")>0 Then
        Ary_Br = Split(Cstr(Br),",")
        For J = 0 To UBound(Ary_Br)
          If Cstr(I)=Cstr(Ary_Br(J)) Then Response.Write "<br>"
        Next
      Else
         If Cstr(I)=Cstr(Br) Then Response.Write "<br>"
      End If
    End If
    If RS1(FName)="其他" Or RS1(FName)="其它" Then 
      Response.Write "<input type='checkbox' " & StrChecked & " name='" & FName &"' id='" & FName & I & "' value='" & RS1(FName) & "'>&nbsp;" & RS1(Listfield) & "&nbsp;"
      Response.Write "&nbsp;<input name='" & FName &"_Other' type='text' size='10' maxlength='20' value='" & Other & "' />" 
    Else
      Response.Write "<input type='checkbox' " & StrChecked & " name='" & FName &"' id='" & FName & I & "' value='" & RS1(FName) & "'>&nbsp;" & RS1(Listfield) & "&nbsp;"
    End if
    I = I + 1
    RS1.MoveNext
  WEnd
  RS1.Close
End Function

Function CheckMinNumberJ (Form,FName,ListField,MinNum)  
  Response.Write "if(Number(document." & Form & "." & FName & ".value)<" & MinNum & "){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  必須大於" & MinNum & "元！');" & Chr(13) & Chr(10)
  Response.Write "  document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function CheckMaxNumberJ (Form,FName,ListField,MaxNum)  
  Response.Write "if(Number(document." & Form & "." & FName & ".value)>" & MaxNum & "){" & Chr(13) & Chr(10)
  Response.Write "  alert('"& ListField & "  必須小於" & MaxNum & "元！');" & Chr(13) & Chr(10)
  Response.Write "  document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function DiffDateJ (Form,FName1,FName2)
  Response.Write "var begindate=new Date(document." & Form & "." & FName1 & ".value)" & Chr(13) & Chr(10)
  Response.Write "var enddate=new Date(document." & Form & "." & FName2 & ".value)" & Chr(13) & Chr(10)
  Response.Write "var diffdate=(Date.parse(begindate.toString())-Date.parse(enddate.toString()))/(1000*60*60*24)" & Chr(13) & Chr(10)
  Response.Write "if(parseInt(diffdate)>0){" & Chr(13) & Chr(10)
  Response.Write "  alert('刊登起日不可大於迄日！');" & Chr(13) & Chr(10)
  Response.Write "  document." & Form & "." & FName1 & ".focus();" & Chr(13) & Chr(10)
  Response.Write "  return;" & Chr(13) & Chr(10)  
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function CheckImage_Width_HeightJ (Form,FName,image_width,image_height)
  Response.Write "if(document." & Form & "." & FName & ".value!=''){" & Chr(13) & Chr(10)
  Response.Write "  var image = new Image();" & Chr(13) & Chr(10)
  Response.Write "  image.src = document." & Form & "."&FName&".value;" & Chr(13) & Chr(10)
  Response.Write "  var iwidth=image.width;" & Chr(13) & Chr(10)
  Response.Write "  var iheight=image.height;" & Chr(13) & Chr(10)
  Response.Write "  document." & Form & "." & image_width & ".value=iwidth;" & Chr(13) & Chr(10)
  Response.Write "  document." & Form & "." & image_height & ".value=iheight; " & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10) 
End Function

Function Calendar (Form,FName,Val)
  Response.Write "<input type=text name="&FName&" size=10 class=font9 value="&Val&">" & Chr(13) & Chr(10)
  Response.Write "<a href onclick=cal19.select(document."&Form&"."&FName&",'"&FName&"','yyyy/MM/dd');>" & Chr(13) & Chr(10)
  Response.Write "<img border=0 src=../images/date.gif width=16 height=14 style=cursor:hand;></a>" & Chr(13) & Chr(10)
End Function

Function SubmitJ (Form,SubmitType)
  if SubmitType="delete" then Response.Write "if(confirm('您是否確定要刪除？')){" & Chr(13) & Chr(10)
  if SubmitType="update" then Response.Write "if(confirm('您是否確定要修改？')){" & Chr(13) & Chr(10)
  if SubmitType="logout" then Response.Write "if(confirm('您是否確定要登出？')){" & Chr(13) & Chr(10)
  Response.Write "document." & Form & ".action.value='"&SubmitType&"';" & Chr(13) & Chr(10)
  Response.Write "document." & Form & ".submit();" & Chr(13) & Chr(10)
  if SubmitType="delete" or SubmitType="update"  or SubmitType="logout" then Response.Write "}" & Chr(13) & Chr(10)
End Function

Function MaxFileUpLoad(FilePath,UploadName1,UploadSize1,objUpload,MaxFileSize)
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

Function MaxFileUpLoads(FilePath,UploadName1,UploadSize1,UploadName2,UploadSize2,UploadName3,UploadSize3,UploadName4,UploadSize4,UploadName5,UploadSize5,UploadName6,UploadSize6,UploadName7,UploadSize7,UploadName8,UploadSize8,UploadName9,UploadSize9,UploadName10,UploadSize10,objUpload,MaxFileSize,MaxUploadSize)
  set objUpload = Server.CreateObject("Dundas.Upload")
  on error resume next
  objUpload.MaxFileCount = 10
  objUpload.MaxFileSize = MaxFileSize*1048576
  objUpload.MaxUploadSize = MaxUploadSize*1048576
  objUpload.UseUniqueNames = True
  objUpload.UseVirtualDir = True
  objUpload.Save FilePath  
  
  I=0
  For Each objUploadedFile in objUpload.Files
    I = I + 1
    If objUploadedFile.Size>0 Then
      If I=1 Then
        UploadName1 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize1 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
      ElseIf I=2 Then
        UploadName2 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize2 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
      ElseIf I=3 Then
        UploadName3 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize3 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
      ElseIf I=4 Then
        UploadName4 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize4 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
      ElseIf I=5 Then
        UploadName5 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize5 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
      ElseIf I=6 Then
        UploadName6 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize6 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
      ElseIf I=7 Then
        UploadName7 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize7 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","") 
      ElseIf I=8 Then
        UploadName8 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize8 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
      ElseIf I=9 Then
        UploadName9 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize9 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")
      ElseIf I=10 Then
        UploadName10 = objUpload.GetFileName(objUploadedFile.path)
        UploadSize10 = replace(FormatNumber((objUploadedFile.Size/1024),2),",","")                                                       
      End If
    End If
  Next
End Function

Function MaxFileUpLoadv(FilePath,UploadName1,UploadSize1,objUpload,MaxUploadSize)
  set objUpload = Server.CreateObject("Dundas.Upload")
  on error resume next
  objUpload.MaxFileCount = 1
  objUpload.MaxFileSize = MaxUploadSize*1048576
  objUpload.MaxUploadSize = MaxUploadSize*1048576
  objUpload.UseUniqueNames = false
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

Function ONumber()
  YYYY = Cstr(Cint(Year(Now()))-2000)
  MM = Cstr(Month(Now()))
  DD = Cstr(Day(Now()))
  If Len(MM) = 1 Then MM = "0"&MM
  If Len(DD) = 1 Then DD = "0"&DD
  SQL1="Select OrderNumber=Isnull(Max(MPP_OrderNumber),000) From DONATE_WEB Where MPP_OrderNumber Like '"&YYYY&MM&DD&"%'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  If Not RS1.EOF Then 
    ONumber=Cstr(Cint(Right(RS1("OrderNumber"),3))+1)
    ONumber=YYYY&MM&DD&Left("000",3-Len(Cstr(ONumber)))&Cstr(ONumber)
  Else
    ONumber=YYYY&MM&DD&"001"
  End If
  RS1.Close
  Set RS1 = Nothing
End Function

Function OrgNumber()
  YYYY = Cstr(Year(Now()))
  MM = Cstr(Month(Now()))
  DD = Cstr(Day(Now()))
  If Len(MM) = 1 Then MM = "0"&MM
  If Len(DD) = 1 Then DD = "0"&DD
  Uniqueod_OrgNumber = "False"
  While Uniqueod_OrgNumber = "False"
    OrgNumber = YYYY & MM & DD
    For I = 1 to 4
      Randomize
      OrgNumber = OrgNumber & Chr(int(26*Rnd)+65)
    Next
    SQL1="Select * From DONATE_WEB Where MPP_OrgOrderNumber='"&OrgNumber&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    If RS1.EOF Then 
      Uniqueod_OrgNumber = "True"
    End If  
    RS1.Close
    Set RS1 = Nothing
  Wend
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

Function ShowMedia(FileType,FileUrl,xWidth,xHeight)
  If FileType="YouTube" Or FileType="無名小站" Then 
    For I=1 To Len(FileUrl)-6
      If Mid(FileUrl,I,5)="width" Then Album_Width=Mid(FileUrl,I,11)
      If Mid(FileUrl,I,6)="height" Then Album_Height=Mid(FileUrl,I,12)
      If Album_Width<>"" And Album_Height<>"" Then Exit For
    Next
    FileUrl=Replace(FileUrl,Album_Width,"width="""&xWidth&"""")
    FileUrl=Replace(FileUrl,Album_Height,"height="""&xHeight&"""")
    FileUrl=Replace(FileUrl,"allowfullscreen=""true""","allowfullscreen=""true"" wmode=""Opaque""")
    Response.Write ""&FileUrl&""& vbCRLF
  ElseIf FileType="flash" Or FileType="WMV" Or FileType="flv" Then
    If xHeight<>"" Then xHeight="height='"&xHeight&"'"
    If Filetype="flash" Then
      Response.write "<object classid='clsid:D27CDB6E-AE6D-11CF-96B8-444553540000' id='obj1' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0' border='0' width='"&xWidth&"'>"
      Response.write "<param name='movie' value='"&FileUrl&"'>"
      Response.write "<param name='quality' value='High'>"
      Response.write "<embed src='"&FileUrl&"' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' name='obj1' width='"&xWidth&"' quality='High'></object>"
    Elseif FileType="wmv" Then
      Response.write "<EMBED src='"&FileUrl&"' autostart=true controls=console width='"&xWidth&"' "&xHeight&" type=video/x-ms-wmv></EMBED>" 
    Elseif FileType="flv" Then
      Response.write "<embed src='include/vcastr.swf?vcastr_file=../"&FileUrl&"' allowfullscreen='true' showMovieInfo=0 pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' wmode='transparent' quality='high' width='"&xWidth&"' "&xHeight&"></embed>" 
    End If
  ElseIf FileType="av" Then
    Response.write "<object id='MediaPlayer' classid='CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6' VIEWASTEXT>"
    Response.write "  <param name='AllowChangeDisplaySize' value='1'>"
    Response.write "  <param name='AutoStart' value='false'>"
    Response.write "  <param name='AutoSize' value='0'>"
    Response.write "  <param name='AnimationAtStart' value='1'>"
    Response.write "  <param name='ClickToPlay' value='1'>"
    Response.write "  <param name='EnableContextMenu' value='0'>"
    Response.write "  <param name='EnablePositionControls' value='1'>"
    Response.write "  <param name='EnableFullScreenControls' value='1'>"
    Response.write "  <param name='URL' value='"&FileUrl&"'>"
    Response.write "  <param name='ShowControls' value='1'>"
    Response.write "  <param name='ShowAudioControls' value='1'>"
    Response.write "  <param name='ShowDisplay' value='0'>"
    Response.write "  <param name='ShowGotoBar' value='0'>"
    Response.write "  <param name='ShowPositionControls' value='1'>"
    Response.write "  <param name='ShowStatusBar' value='1'>"
    Response.write "  <param name='ShowTracker' value='1'>"
    Response.write "  <embed src='"&FileUrl&"'" 
    Response.write "    type='video/x-ms-wmv'" 
    Response.write "    width='"&xWidth&"' height='"&xHeight&"'"
    Response.write "    AutoStart='0'"
    Response.write "    ShowControls='0'"
    Response.write "    AutoSize='1'"
    Response.write "    AnimationAtStart='1'"
    Response.write "    ClickToPlay='1'"
    Response.write "    EnableContextMenu='0'"
    Response.write "    EnablePositionControls='1'"
    Response.write "    EnableFullScreenControls='1'"
    Response.write "    ShowControls='1'"
    Response.write "    ShowAudioControls='1'"
    Response.write "    ShowDisplay='0'"
    Response.write "    ShowGotoBar='0'"
    Response.write "    ShowPositionControls='1'"
    Response.write "    ShowStatusBar='1'"
    Response.write "    ShowTracker='1'>"
    Response.write "  </embed>"
    Response.write " </object>"
  End If
End Function

Function SignUp_Msg(Activity_Id,Activity_Type,Activity_Page)
  SignUp_Msg=""
  SQL_Act="Select * From ACTIVITY Where Ser_No='"&Activity_Id&"'"
  Call QuerySQL(SQL_Act,RS_Act)
  If Not RS_Act.EOF Then
    If RS_Act("Activity_Signup")="Y" And RS_Act("Activity_Signup_BeginDate")<>"" And RS_Act("Activity_Signup_BeginTime")<>"" And RS_Act("Activity_Signup_EndDate")<>"" And RS_Act("Activity_Signup_EndTime")<>"" And RS_Act("Activity_SignUp_Limit")<>"" Then
      HH=Cstr(Hour(Now()))
      If Len(HH)=1 Then HH="0"&HH
      MM=Cstr(Minute(Now()))
      If Len(MM)=1 Then MM="0"&MM
      If DateDiff("D",Date(),RS_Act("Activity_Signup_BeginDate"))>0 Or (DateDiff("D",Date(),RS_Act("Activity_Signup_BeginDate"))=0 And Cint(HH&MM)<Cint(RS_Act("Activity_Signup_BeginTime"))) Then 
        SignUp_Msg="開放報名時間<br>"&Left("00",2-Len(Month(RS_Act("Activity_Signup_BeginDate"))))&Month(RS_Act("Activity_Signup_BeginDate"))&"/"&Left("00",2-Len(Day(RS_Act("Activity_Signup_BeginDate"))))&Day(RS_Act("Activity_Signup_BeginDate"))&"&nbsp;"&Left(RS_Act("Activity_Signup_BeginTime"),2)&":"&Right(RS_Act("Activity_Signup_BeginTime"),2)
        Exit Function
      End If
      If DateDiff("D",RS_Act("Activity_Signup_EndDate"),Date())>0 Or (DateDiff("D",RS_Act("Activity_Signup_EndDate"),Date())=0 And Cint(HH&MM)>Cint(RS_Act("Activity_Signup_EndTime"))) Then
		    SignUp_Msg="<img src=""public/image/signup.gif"" width=""8"" height=""8"" alt=""報名截止"">&nbsp;<font color=""#FF0000"">截止</font>"
        If RS_Act("Activity_Member")="1" Then
          SignUp_Msg=SignUp_Msg&"&nbsp;<img src=""public/image/signup.gif"" width=""8"" height=""8"" alt=""取消報名"">&nbsp;<a href=""activity_cancel.asp?"&activity_url&""" class=""signup"">取消</a>"
        ElseIf RS_Act("Activity_Member")="2" Then
          If Session("Member_Id")<>"" Or Session("Member_No")<>"" Then
            SQL_Cancel="Select * From ACTIVITY_SIGNUP Where Activity_Id='"&Activity_Id&"' And (Member_Id='"&Session("Member_Id")&"' Or Member_Id='"&Session("Member_No")&"')"
            Call QuerySQL(SQL_Cancel,RS_Cancel)
            If Not RS_Cancel.EOF Then SignUp_Msg="&nbsp;<img src=""public/image/signup.gif"" width=""8"" height=""8"" alt=""取消報名"">&nbsp;<a href=""activity_cancel.asp?"&activity_url&""" class=""signup"">取消</a>"
            RS_Cancel.Close
            Set RS_Cancel=Nothing  
          End If
        ElseIf RS_Act("Activity_Member")="3" Then
        
        End If
        Exit Function
      End If
      activity_url="id="&Activity_Id
      If Activity_Type<>"" Then activity_url=activity_url&"&type="&Activity_Type
      If Activity_Page<>"" Then activity_url=activity_url&"&activity_page="&Activity_Page 
      SQL_SignUp="Select Total=Count(*) From ACTIVITY_SignUp Where Activity_Id='"&RS_Act("Ser_No")&"'"
      Call QuerySQL(SQL_SignUp,RS_SignUp)
      If Cint(RS_SignUp("Total"))>=Cint(RS_Act("Activity_SignUp_Limit")) Then
        SignUp_Msg="<img src=""public/image/signup.gif"" width=""8"" height=""8"" alt=""報名額滿"">&nbsp;<font color=""#FF0000"">額滿</font>"	
        If RS_Act("Activity_Member")="1" Then
          SignUp_Msg=SignUp_Msg&"&nbsp;<img src=""public/image/signup.gif"" width=""8"" height=""8"" alt=""取消報名"">&nbsp;<a href=""activity_cancel.asp?"&activity_url&""" class=""signup"">取消</a>"
        ElseIf RS_Act("Activity_Member")="2" Then
          If Session("Member_Id")<>"" Or Session("Member_No")<>"" Then
            SQL_Cancel="Select * From ACTIVITY_SIGNUP Where Activity_Id='"&Activity_Id&"' And (Member_Id='"&Session("Member_Id")&"' Or Member_Id='"&Session("Member_No")&"')"
            Call QuerySQL(SQL_Cancel,RS_Cancel)
            If Not RS_Cancel.EOF Then SignUp_Msg="&nbsp;<img src=""public/image/signup.gif"" width=""8"" height=""8"" alt=""取消報名"">&nbsp;<a href=""activity_cancel.asp?"&activity_url&""" class=""signup"">取消</a>"
            RS_Cancel.Close
            Set RS_Cancel=Nothing  
          End If
        ElseIf RS_Act("Activity_Member")="3" Then
        
        End If
        Exit Function
      Else
        If RS_Act("Activity_Member")="1" Then
          SignUp_Msg="<img src=""public/image/signup.gif"" width=""8"" height=""8"" alt=""我要報名"">&nbsp;<a href=""activity_signup.asp?"&activity_url&""" class=""signup"">報名</a>&nbsp;<img src=""public/image/signup.gif"" width=""8"" height=""8"" alt=""取消報名"">&nbsp;<a href=""activity_cancel.asp?"&activity_url&""" class=""signup"">取消</a>"
        ElseIf RS_Act("Activity_Member")="2" Then
          If (Session("Member_Id")<>"" Or Session("Member_No")<>"") Then
            SQL_Cancel="Select * From ACTIVITY_SIGNUP Where Activity_Id='"&Activity_Id&"' And (Member_Id='"&Session("Member_Id")&"' Or Member_Id='"&Session("Member_No")&"')"
            Call QuerySQL(SQL_Cancel,RS_Cancel)
            If RS_Cancel.EOF Then 
              SignUp_Msg="<img src=""public/image/signup.gif"" width=""8"" height=""8"" alt=""我要報名"">&nbsp;<a href=""activity_signup.asp?"&activity_url&""" class=""signup"">報名</a>"
            Else
              SignUp_Msg="<img src=""public/image/signup.gif"" width=""8"" height=""8"" alt=""取消報名"">&nbsp;<a href=""activity_cancel.asp?"&activity_url&""" class=""signup"">取消</a>"
            End If
            RS_Cancel.Close
            Set RS_Cancel=Nothing
          Else
            SignUp_Msg="<img src=""public/image/signup.gif"" width=""8"" height=""8"" alt=""報名請登入"">&nbsp;報名請登入"
          End If
        ElseIf RS_Act("Activity_Member")="3" Then
          
        End If
      End If  
    End If 
  End If 
  RS_Act.Close
  Set RS_Act=Nothing
End Function

Function Media_URL(Album_URL,Default_Width,Default_Height)
  album_width=""
  album_height=""
  For I=1 To Len(Album_URL)-6
    If Mid(Album_URL,I,5)="width" Then album_width=Mid(Album_URL,I,11)
    If Mid(Album_URL,I,6)="height" Then album_height=Mid(Album_URL,I,12)
    If album_width<>"" And album_height<>"" Then Exit For
  Next
  Media_URL=Album_URL
  Media_URL=Replace(Media_URL,album_width,"width="""&Default_Width&"""")
  Media_URL=Replace(Media_URL,album_height,"height="""&Default_Height&"""")
  Media_URL=Replace(Media_URL,"allowfullscreen=""true""","allowfullscreen=""true"" wmode=""Opaque""")
End Function

Function CodeCity (form,ZipCode,ZipValue,CityCode,CityValue,AreaCode,AreaValue,Address,AddressValue,AddressSize,JS)
  '縣市選單
  RS1_City_Row = 0
  SelCity="<option value=''>縣&nbsp;&nbsp;市</option>"
  SQL_City = "Select mCode,mValue From CODECITY Where codeMetaID = 'Addr0' Order By mSortValue"
  Set RS1_City = Server.CreateObject("ADODB.RecordSet")
  RS1_City.Open SQL_City,Conn,1,1
  Do While Not RS1_City.EOF
    If RS1_City("mCode")=CityValue Then
      SelCity = SelCity & "<option value='" & RS1_City("mCode") & "' Selected>" & RS1_City("mValue") & "</option>"
    Else
      SelCity = SelCity & "<option value='" & RS1_City("mCode") & "'>" & RS1_City("mValue") & "</option>"
    End If
    
    If RS1_City_Row = 0 Then
      Aera_Code = Aera_Code & "'" & RS1_City("mCode") & "'" & ","
    Else 
      Aera_Code = Aera_Code & "," & "'" & RS1_City("mCode") & "'" & ","
    End If
    RS_AreaRow=1
    Aera_Code = Aera_Code & "["
    SQL_Area = "Select mCode,mValue From CODECITY Where codeMetaID = 'Addr0R"&RS1_City("mCode")&"' Order By mSortValue"
    Set RS_Area = Server.CreateObject("ADODB.RecordSet")
    RS_Area.Open SQL_Area,Conn,1,1
    Do While Not RS_Area.EOF
      If RS_AreaRow = 1 Then
        Aera_Code = Aera_Code & "'" & RS_Area("mCode") & "'" & "," & "'" & RS_Area("mValue") & "'" 
      Else 
        Aera_Code = Aera_Code & "," & "'" & RS_Area("mCode") & "'" & "," & "'" & RS_Area("mValue") & "'" 
      End If
      RS_AreaRow=RS_AreaRow+1
      RS_Area.MoveNext
    Loop
    RS_Area.Close
    Set RS_Area = Nothing
    Aera_Code=Aera_Code&"]"
    RS1_City_Row=RS1_City_Row+1
    RS1_City.MoveNext
  Loop
  RS1_City.Close
  Set RS1_City = Nothing
  '鄉鎮市區選單
  SelArea="<option value=''>鄉鎮市區</option>"
  If CityValue<>"" And AreaValue<>"" Then
    SQL_City = "Select mCode,mValue From CodeCity Where codeMetaID = 'Addr0R"&CityValue&"' Order By mSortValue"
    Set RS1_City = Server.CreateObject("ADODB.RecordSet")
    RS1_City.Open SQL_City,Conn,1,1
    Do While Not RS1_City.EOF
      If RS1_City("mCode")=AreaValue Then
        SelArea = SelArea & "<option value='" & RS1_City("mCode") & "' Selected>" & RS1_City("mValue") & "</option>"
      Else
        SelArea = SelArea & "<option value='" & RS1_City("mCode") & "'>" & RS1_City("mValue") & "</option>"
      End If  
      RS1_City.MoveNext
    Loop 
    RS1_City.Close
    Set RS1_City = Nothing
  End If
      
  Response.Write "<input type='text' name='"&ZipCode&"' size='3' readonly maxlength='3' value='"&ZipValue&"'>"
  Response.Write "<Select name='"&CityCode&"' Size='1' class='from_input' onchange='JavaScript:ChgCity(this.value,document."&form&"."&AreaCode&",document."&form&"."&ZipCode&");'>"&SelCity&"</Select>"
  Response.Write "<Select name='"&AreaCode&"' Size='1' class='from_input' onchange='JavaScript:ChgArea(this.value,document."&form&"."&ZipCode&");'>"&SelArea&"</Select>"
  Response.Write "<input type='text' name='"&Address&"' size='"&AddressSize&"' maxlength='100' value='"&AddressValue&"'>" & Chr(13) & Chr(10) 
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

Function CodeCity2 (form,ZipCode,ZipValue,CityCode,CityValue,AreaCode,AreaValue,Address,AddressValue,AddressSize,JS)
  '縣市選單
  RS1_City_Row = 0
  SelCity="<option value=''>縣&nbsp;&nbsp;市</option>"
  SQL_City = "Select mCode,mValue From CODECITY Where codeMetaID = 'Addr0' Order By mSortValue"
  Set RS1_City = Server.CreateObject("ADODB.RecordSet")
  RS1_City.Open SQL_City,Conn,1,1
  Do While Not RS1_City.EOF
    If RS1_City("mCode")=CityValue Then
      SelCity = SelCity & "<option value='" & RS1_City("mCode") & "' Selected>" & RS1_City("mValue") & "</option>"
    Else
      SelCity = SelCity & "<option value='" & RS1_City("mCode") & "'>" & RS1_City("mValue") & "</option>"
    End If
    
    If RS1_City_Row = 0 Then
      Aera_Code = Aera_Code & "'" & RS1_City("mCode") & "'" & ","
    Else 
      Aera_Code = Aera_Code & "," & "'" & RS1_City("mCode") & "'" & ","
    End If
    RS_AreaRow=1
    Aera_Code = Aera_Code & "["
    SQL_Area = "Select mCode,mValue From CODECITY Where codeMetaID = 'Addr0R"&RS1_City("mCode")&"' Order By mSortValue"
    Set RS_Area = Server.CreateObject("ADODB.RecordSet")
    RS_Area.Open SQL_Area,Conn,1,1
    Do While Not RS_Area.EOF
      If RS_AreaRow = 1 Then
        Aera_Code = Aera_Code & "'" & RS_Area("mCode") & "'" & "," & "'" & RS_Area("mValue") & "'" 
      Else 
        Aera_Code = Aera_Code & "," & "'" & RS_Area("mCode") & "'" & "," & "'" & RS_Area("mValue") & "'" 
      End If
      RS_AreaRow=RS_AreaRow+1
      RS_Area.MoveNext
    Loop
    RS_Area.Close
    Set RS_Area = Nothing
    Aera_Code=Aera_Code&"]"
    RS1_City_Row=RS1_City_Row+1
    RS1_City.MoveNext
  Loop
  RS1_City.Close
  Set RS1_City = Nothing
  '鄉鎮市區選單
  SelArea="<option value=''>鄉鎮市區</option>"
  If CityValue<>"" And AreaValue<>"" Then
    SQL_City = "Select mCode,mValue From CodeCity Where codeMetaID = 'Addr0R"&CityValue&"' Order By mSortValue"
    Set RS1_City = Server.CreateObject("ADODB.RecordSet")
    RS1_City.Open SQL_City,Conn,1,1
    Do While Not RS1_City.EOF
      If RS1_City("mCode")=AreaValue Then
        SelArea = SelArea & "<option value='" & RS1_City("mCode") & "' Selected>" & RS1_City("mValue") & "</option>"
      Else
        SelArea = SelArea & "<option value='" & RS1_City("mCode") & "'>" & RS1_City("mValue") & "</option>"
      End If  
      RS1_City.MoveNext
    Loop 
    RS1_City.Close
    Set RS1_City = Nothing
  End If

  Response.Write "<Select name='"&CityCode&"' Size='1' class='from_input' onchange='JavaScript:ChgCity(this.value,document."&form&"."&AreaCode&",document."&form&"."&ZipCode&");'>"&SelCity&"</Select>"
  Response.Write "<Select name='"&AreaCode&"' Size='1' class='from_input' onchange='JavaScript:ChgArea(this.value,document."&form&"."&ZipCode&");'>"&SelArea&"</Select>"
  Response.Write "<input type='text' name='"&Address&"' size='"&AddressSize&"' maxlength='100' value='"&AddressValue&"'>" & Chr(13) & Chr(10) 
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

Function Ads_Logo(Ads_Type,Ads_Width,Ads_Height)
  SQL_Ads="Select TOP 1 * From ADS Where Ads_Type='"&Ads_Type&"' And ('"&Date()&"' Between Ads_BeginDate And Ads_EndDate) And (Ads_TitleImg<>'' Or Convert(varchar(1000), Ads_Source)<>'') Order By Ads_Type,Ads_Top,Ads_Seq,Ads_BeginDate Desc,Ads_EndDate Desc,Ads_RegDate Desc,Ser_No Desc"
	Call QuerySQL(SQL_Ads,RS_Ads)
  If Not RS_Ads.EOF Then
	  If RS_Ads("Ads_Item")="image" Then
		  If RS_Ads("Ads_Redirect_Url")<>"" Then
		    Response.Write "<a href=""ads_url.asp?id="&RS_Ads("Ser_No")&""" target="""&RS_Ads("Ads_Open")&"""><img border=""0"" src=""upload/"&RS_Ads("Ads_TitleImg")&""" width="""&Ads_Width&""" alt="""&RS_Ads("Ads_Subject")&"""></a>"			            
	    Else
				Response.Write "<img border=""0"" src=""upload/"&RS_Ads("Ads_TitleImg")&""" width="""&Ads_Width&""" alt="""&RS_Ads("Ads_Subject")&""">"
			End If
	  ElseIf RS_Ads("Ads_Item")="flash" Or RS_Ads("Ads_Item")="flv" Then
      Call ShowMedia(RS_Ads("Ads_Item"),"upload/"&RS_Ads("Ads_TitleImg"),""&Ads_Width&"",""&Ads_Height&"")	  
		ElseIf RS_Ads("Ads_Item")="source" Then
			Response.Write RS_Ads("Ads_Source")
		Else
			Call ShowMedia(RS_Ads("Ads_Item"),RS_Ads("Ads_TitleImg"),""&Ads_Width&"",""&Ads_Height&"")	
		End If
  Else
    Response.Write "<table width="""&Ads_Width&""" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" bgcolor=""#666666"">"& vbCRLF
    Response.Write "  <tr>"& vbCRLF
    Response.Write "    <td height="""&Ads_Height&"""> </td>"& vbCRLF
    Response.Write "  </tr>"& vbCRLF
    Response.Write "</table>"& vbCRLF
  End If
  RS_Ads.Close
  Set RS_Ads=Nothing
End Function

Function Check_Vote (Vote_Id)
  Check_Vote=True
  SQL_Vote="Select * From VOTE Where Ser_No='"&Vote_Id&"'"
  Call QuerySQL(SQL_Vote,RS_Vote)
  If Not RS_Vote.EOF Then
    Vote_Max=1
    If RS_Vote("Vote_SurveyType")="3" And RS_Vote("Vote_SurveyMax")<>"" Then Vote_Max=Cint(RS_Vote("Vote_SurveyMax"))
    If DateDiff("D",RS_Vote("Vote_BeginDate"),Date())<0 Or DateDiff("D",Date(),RS_Vote("Vote_EndDate"))<0 Then
      Check_Vote="投票時間已結束"
      Exit Function
    End If
    If RS_Vote("Vote_Type")="webuser" Then
      If RS_Vote("Vote_SurveyType")="1" Or RS_Vote("Vote_SurveyType")="3" Then
        SQL_Survey="Select Total=Count(*) From VOTE_SURVEY Where Survey_IP='"&Request.ServerVariables("REMOTE_HOST")&"'"
      ElseIf RS_Vote("Vote_SurveyType")="2" Then
        SQL_Survey="Select Total=Count(*) From VOTE_SURVEY Where Survey_IP='"&Request.ServerVariables("REMOTE_HOST")&"' And Survey_Date='"&Date()&"'"
      End If
      Call QuerySQL(SQL_Survey,RS_Survey)
      Total=Cint(RS_Survey("Total"))
      RS_Survey.Close
      Set RS_Survey=Nothing
      If Total>=Vote_Max Then
        Check_Vote="很抱歉系統已經有您投票紀錄，請勿重覆投票"
        Exit Function
      End If  
    End If
    If RS_Vote("Vote_Type")="member" Then
      If session("Member_No")="" Then
        Check_Vote="投票請先登入"
        Exit Function
      End If
      If RS_Vote("Vote_SurveyType")="1" Or RS_Vote("Vote_SurveyType")="3" Then
        SQL_Survey="Select Total=Count(*) From VOTE_SURVEY Where Survey_MemberId='"&session("Member_Id")&"'"
      ElseIf RS_Vote("Vote_SurveyType")="2" Then
        SQL_Survey="Select Total=Count(*) From VOTE_SURVEY Where Survey_MemberId='"&session("Member_Id")&"' And Survey_Date='"&Date()&"'"
      End If
      Call QuerySQL(SQL_Survey,RS_Survey)
      Total=RS_Survey("Total")
      RS_Survey.Close
      Set RS_Survey=Nothing
      If Total>=Vote_Max Then
        Check_Vote="很抱歉系統已經有您投票紀錄，請勿重覆投票"
        Exit Function
      End If 
    End If    
  Else
    Check_Vote=False
  End If
  RS_Vote.Close
  Set RS_Vote=Nothing
End Function

Function Check_Url()
  Check_Url = True
  URL1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
  URL2 = Cstr(Request.ServerVariables("SERVER_NAME"))
  If Instr(Cstr(Request.ServerVariables("HTTP_REFERER")),"https")>0 Then
    URL3=Mid(URL1,9,Len(URL2))
  Else
    URL3=Mid(URL1,8,Len(URL2))
  End If
  If URL3<>URL2 Then Check_Url=False
End Function

Function Check_Request_Form(Page_Name)
  Form_Check=True
  If Check_Url()=False Or Request.ServerVariables("query_string")<>"" Then Form_Check=False
  If Donate_Type<>"creditcard" And Donate_Type<>"webatm" And Donate_Type<>"vacc" And Donate_Type<>"barcode" And Donate_Type<>"ibon" And Donate_Type<>"famiport" And Donate_Type<>"lifeet" And Donate_Type<>"okgo" Then Form_Check=False
  'If check_id<>"1" And check_id<>"2" Then Form_Check=False
  If Page_Name="a" Then
    If check_id="1" Then
      If Donate_Name="" Or Donate_IDNO="" Then Form_Check=False
      SQL1="Select TOP 1 * From DONATE_WEB Where Donate_IDNO='"&UCase(Donate_IDNO)&"' Order By Donor_Id Desc,Ser_No Desc"
      Call QuerySQL(SQL1,RS1)
      If Not RS1.EOF Then
        If Cstr(RS1("Donate_DonorName"))<>Cstr(Donate_Name) Then 
          session("errnumber")=1
          session("msg")="您的身分證號 / 統一編號已使用過，請您再次確認，如有疑問請洽本會。"
          Response.Redirect "ecpay.asp"
        End If
      End If
      RS1.Close
      Set RS1=Nothing
    ElseIf check_id="2" Then
      If Donate_Name="" Or Donate_ZipCode="" Or Donate_Address="" Then Form_Check=False
    End If
  ElseIf Page_Name="b" Then
    If Dept_Id="" Or Donate_Type="" Or Donate_Amount="" Or Donate_Name="" Or Donate_CellPhone="" Or Donate_ZipCode="" Or Donate_Address="" Or Donate_Email="" Then Form_Check=False
  End If
  If Form_Check=False Then 
    session("errnumber")=1
    session("msg")="很抱歉您輸入的訊息不足，誠摯地請您再試一次！"
    Response.Redirect "ecpay.asp"
  Else
    session("errnumber")=0
    session("msg")=""  
  End If
End Function

Function get_od_sob(Od_Sob_Type)
  Set od_sobRec = Server.CreateObject("ADODB.RecordSet")
  YYYY = Cstr(Year(Now()))
  MM = Cstr(Month(Now()))
  DD = Cstr(Day(Now()))
  If Len(MM) = 1 Then MM = "0"&MM
  If Len(DD) = 1 Then DD = "0"&DD
  Uniqueod_sob = "False"
  While Uniqueod_sob = "False"
    get_od_sob = YYYY & MM & DD
    For I = 1 to 4
      Randomize
      get_od_sob = get_od_sob & Chr(int(26*Rnd)+65)
    Next
    SQL_Sob="Select * From DONATE_OD_SOB Where od_sob='"&get_od_sob&"'"
    Set RS_Sob = Server.CreateObject("ADODB.RecordSet")
    RS_sob.Open SQL_Sob,Conn,1,1
    If RS_sob.EOF Then 
      Uniqueod_sob = "True"
      SQL1_Sob="DONATE_OD_SOB"
      Set RS1_Sob = Server.CreateObject("ADODB.RecordSet")
      RS1_Sob.Open SQL1_Sob,Conn,1,3
      RS1_Sob.Addnew
      RS1_Sob("Od_Sob")=get_od_sob
      RS1_Sob("Od_Sob_Type")=Od_Sob_Type
      RS1_Sob("Create_Date")=Date()
      RS1_Sob("Create_DateTime")=Now()
      RS1_Sob("Create_User")=session("user_name")
      RS1_Sob("Create_IP")=Request.ServerVariables("REMOTE_HOST")
      RS1_Sob.Update
      RS1_Sob.Close
      Set RS1_Sob=Nothing
    End If  
    RS_sob.Close
    Set RS_sob = Nothing    
    'SQL_sob="Select * From DONATE_WEB Where od_sob='"&get_od_sob&"'"
    'Set RS_sob = Server.CreateObject("ADODB.RecordSet")
    'RS_sob.Open SQL_sob,Conn,1,1
    'If RS_sob.EOF Then 
    '  Uniqueod_sob = "True"
    'End If  
    'RS_sob.Close
    'Set RS_sob = Nothing
  Wend
End Function

Function Get_InvoiceNo(Dept_Id,Donate_Date,Invoice_Pre_Type,Donate_Invoice_Type)
  Get_InvoiceNo=""
  '取號參數
  InvoiceTypeY="年度匯整"
  Set RS_InvoiceNo=Server.CreateObject("ADODB.RecordSet")
  RS_InvoiceNo.Open "Select InvoiceTypeY=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Like '%年%' Order By Seq", Conn, 1, 1
  If Not RS_InvoiceNo.EOF Then InvoiceTypeY=RS_InvoiceNo("InvoiceTypeY")
  RS_InvoiceNo.Close
  Set RS_InvoiceNo=Nothing

  InvoiceTypeN="不寄收據"
  Set RS_InvoiceNo=Server.CreateObject("ADODB.RecordSet")
  RS_InvoiceNo.Open "Select InvoiceTypeY=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Like '%不%' Order By Seq", Conn, 1, 1
  If Not RS_InvoiceNo.EOF Then InvoiceTypeN=RS_InvoiceNo("InvoiceTypeY")
  RS_InvoiceNo.Close
  Set RS_InvoiceNo=Nothing

  DonateDesc="物資捐贈"
  Set RS_InvoiceNo=Server.CreateObject("ADODB.RecordSet")
  RS_InvoiceNo.Open "Select DonateDesc=CodeDesc From CASECODE Where CodeType='Payment' And CodeDesc Like '%物%' Order By Seq", Conn, 1, 1
  If Not RS_InvoiceNo.EOF Then DonateDesc=RS_InvoiceNo("DonateDesc")
  RS_InvoiceNo.Close
  Set RS_InvoiceNo=Nothing
        
  Set RS_InvoiceNo=Server.CreateObject("ADODB.RecordSet")
  RS_InvoiceNo.Open "Select * From DEPT Where Dept_Id='"&Dept_Id&"'", Conn, 1, 1      
  Invoice_Pre=RS_InvoiceNo("Invoice_Pre")
  Invoice_Rule_Type=RS_InvoiceNo("Invoice_Rule_Type")
  Invoice_Rule_YMD=RS_InvoiceNo("Invoice_Rule_YMD")
  Invoice_Rule_Len=RS_InvoiceNo("Invoice_Rule_Len")
  Invoice_Rule_Pub=RS_InvoiceNo("Invoice_Rule_Pub")
  Invoice_Pre2=RS_InvoiceNo("Invoice_Pre2")
  Invoice_Rule_Type2=RS_InvoiceNo("Invoice_Rule_Type2")
  Invoice_Rule_YMD2=RS_InvoiceNo("Invoice_Rule_YMD2")
  Invoice_Rule_Len2=RS_InvoiceNo("Invoice_Rule_Len2")
  Invoice_Rule_Pub2=RS_InvoiceNo("Invoice_Rule_Pub2")
  Donate_Invoice=RS_InvoiceNo("Donate_Invoice")
  RS_InvoiceNo.Close
  Set RS_InvoiceNo=Nothing
  
  '捐款(物)收據取號
  Invoice_No=""
  Check_Invoice=True
  If Cstr(Donate_Invoice_Type)=Cstr(InvoiceTypeY) And Donate_Invoice="N" Then Check_Invoice=False
  If Check_Invoice Then Get_InvoiceNo=Get_Invoice_No(Invoice_Pre_Type,Dept_Id,Donate_Date,Invoice_Rule_Type,DonateDesc,Invoice_Rule_Type,Invoice_Rule_YMD,Invoice_Rule_Len,Invoice_Rule_Pub)   
End Function

Function Get_Invoice_No(Invoice_Pre_Type,Dept_Id,Donate_Date,Invoice_Rule_Type2,DonateDesc,Invoice_Rule_Type,Invoice_Rule_YMD,Invoice_Rule_Len,Invoice_Rule_Pub)
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
    SQL_InvoiceNo="Select Top 1 Invoice_No From DONATE Where Invoice_No<>'' And (Issue_Type='' Or Issue_Type Is null) "
    If Invoice_Rule_Pub="N" Then SQL_InvoiceNo=SQL_InvoiceNo&"And Dept_Id='"&Dept_Id&"' "
    If Cstr(Invoice_Pre_Type)="1" Then
      If Invoice_Rule_Type2<>"P" Then SQL_InvoiceNo=SQL_InvoiceNo&"And (Donate_Payment<>'"&DonateDesc&"' Or Donate_Payment Is Null) "
    Else
      If Invoice_Rule_Type2<>"P" Then SQL_InvoiceNo=SQL_InvoiceNo&"And Donate_Payment='"&DonateDesc&"' "
    End If
    If Invoice_Rule_YMD="Y" Then
      SQL_InvoiceNo=SQL_InvoiceNo&"And Year(Donate_Date)="&Donate_Year&" "
      If Invoice_Rule_Type="R" Then
        InvoiceNo=Right("000"&Cstr(Donate_Date_RA),3)
      Else
        InvoiceNo=Right("0000"&Cstr(Donate_Date_RA),4)
      End If
    ElseIf Invoice_Rule_YMD="M" Then
      SQL_InvoiceNo=SQL_InvoiceNo&"And Year(Donate_Date)="&Donate_Year&" And Month(Donate_Date)="&Donate_Month&" "
      If Invoice_Rule_Type="R" Then
        InvoiceNo=Right("000"&Cstr(Donate_Date_RA),3)&Right("00"&Cstr(Donate_Month),2)
      Else
        InvoiceNo=Right("0000"&Cstr(Donate_Date_RA),4)&Right("00"&Cstr(Donate_Month),2)
      End If
    Else
      SQL_InvoiceNo=SQL_InvoiceNo&"And Year(Donate_Date)="&Donate_Year&" And Month(Donate_Date)="&Donate_Month&" And Day(Donate_Date)="&Donate_Day&" "
      If Invoice_Rule_Type="R" Then
        InvoiceNo=Right("000"&Cstr(Donate_Date_RA),3)&Right("00"&Cstr(Donate_Month),2)&Right("00"&Cstr(Donate_Day),2)
      Else
        InvoiceNo=Right("0000"&Cstr(Donate_Date_RA),4)&Right("00"&Cstr(Donate_Month),2)&Right("00"&Cstr(Donate_Day),2)
      End If
    End If
    SQL_InvoiceNo=SQL_InvoiceNo&"Order By Right(Invoice_No,"&Invoice_Rule_Len&") Desc"
    Call QuerySQL(SQL_InvoiceNo,RS_InvoiceNo)
    If RS_InvoiceNo.EOF Then
      Get_Invoice_No=InvoiceNo&Right(InvoiceNo2&CStr(CDbl(Right(InvoiceNo2,Invoice_Rule_Len))+1),Invoice_Rule_Len)
    Else
        Get_Invoice_No=InvoiceNo&Right(InvoiceNo2&CStr(CDbl(Right(RS_InvoiceNo("Invoice_No"),Invoice_Rule_Len))+1),Invoice_Rule_Len)
    End If
    RS_InvoiceNo.Close
    Set RS_InvoiceNo=Nothing
  ElseIf Invoice_Rule_Type="S" Then
    SQL_InvoiceNo="Select Top 1 Invoice_No From DONATE Where Invoice_No<>'' And (Issue_Type='' Or Issue_Type Is null) And Convert(numeric,Right(Invoice_No,"&Invoice_Rule_Len&"))>0 "
    If Invoice_Rule_Pub="N" Then SQL_InvoiceNo=SQL_InvoiceNo&"And Dept_Id='"&Dept_Id&"' "
    If Cstr(Invoice_Pre_Type)="1" Then
      If Invoice_Rule_Type2<>"P" Then SQL_InvoiceNo=SQL_InvoiceNo&"And (Donate_Payment<>'"&DonateDesc&"' Or Donate_Payment Is Null) "
    Else
      If Invoice_Rule_Type2<>"P" Then SQL_InvoiceNo=SQL_InvoiceNo&"And Donate_Payment='"&DonateDesc&"' "
    End If
    SQL_InvoiceNo=SQL_InvoiceNo&"Order By Convert(numeric,Right(Invoice_No,"&Invoice_Rule_Len&")) Desc"
    Call QuerySQL(SQL_InvoiceNo,RS_InvoiceNo)
    If RS_InvoiceNo.EOF Then
      Get_Invoice_No=Right(InvoiceNo2&CStr(CDbl(Right(InvoiceNo2,Invoice_Rule_Len))+1),Invoice_Rule_Len)
    Else
      Get_Invoice_No=Right(InvoiceNo2&CStr(CDbl(Right(RS_InvoiceNo("Invoice_No"),Invoice_Rule_Len))+1),Invoice_Rule_Len)
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
      SQL_InvoiceNo="Select Top 1 Invoice_No,Invoice_Right=Right(Invoice_No,"&Invoice_Rule_Len&") From "&TableName&" Where Invoice_No<>'' And Invoice_Pre='"&Invoice_Pre&"' And (Issue_Type_Keep='' Or Issue_Type_Keep Is null) And (Issue_Type='' Or Issue_Type Is null Or Issue_Type='D') "
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
      SQL_InvoiceNo="Select Top 1 Invoice_No,Invoice_Right=Convert(numeric,Right(Invoice_No,"&Invoice_Rule_Len&")) From "&TableName&" Where Invoice_No<>'' And Invoice_Pre='"&Invoice_Pre&"' And (Issue_Type_Keep='' Or Issue_Type_Keep Is null) And (Issue_Type='' Or Issue_Type Is null Or Issue_Type='D') And Convert(numeric,Right(Invoice_No,"&Invoice_Rule_Len&"))>0 "
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
      'Else
      '  SQL_InvoiceNo=SQL_InvoiceNo&"Order By Convert(numeric,Right(Invoice_No,"&Invoice_Rule_Len&")) Desc"  
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

Function get_close_url (gwsr,amount)
  get_close_url=""
  PayName=Card_PayName
  Client_Card=Trim(Card_Code)
  If Len(Client_Card)>6 Then Client_Card=Len(Client_Card)
  CkeckD=Card_Verify
  If Csng(Client_Card)<>3 Then
    '商店代碼Card_Code(6碼,靠右,不足左補0)  
    Client_Card=Left("000000",6-Len(Cstr(Client_Card)))&Cstr(Client_Card)
    '授權單號gwsrcode(10碼,靠右,不足左捕0)
    gwsrcode=Left("0000000000",10-Len(Cstr(gwsr))) & Cstr(gwsr)
    '金額money(8碼,靠右,不足左捕0)
    money=Left("00000000",8-Len(Cstr(amount))) & Cstr(amount)      
    '檢核碼CkeckE(9碼,靠右,不足左捕0)
    Ckeck=Client_Card & gwsrcode & money
    CkeckA=""
    CkeckB=""
    For I = 1 To Len(Ckeck)
      If I>=5 Then
        If I Mod 5 = 0 Then
          If CkeckA<>"" Then
            CkeckA=CkeckA & Mid(Ckeck,I,1)
          Else
            CkeckA=Mid(Ckeck,I,1)
          End If
        End If  
      End If
      If I>=3 Then
        If I Mod 3 = 0 Then
          If CkeckB<>"" Then
            CkeckB=CkeckB & Mid(Ckeck,I,1)
          Else
            CkeckB=Mid(Ckeck,I,1)
          End If
        End If
      End If
    Next
    '亂數碼CkeckC
    For I = 1 to 3
      Randomize
      CkeckC = CkeckC & Chr(int(10*Rnd)+48)
    Next
    CkeckE_Mod=(CkeckA*CkeckC+CkeckB) Mod CkeckD
    CkeckE=Left("000000000",9-Len(Cstr(CkeckE_Mod))) & Cstr(CkeckE_Mod)
    '關帳URL
    If PayName="ecpay" Then
      get_close_url="https://ecpay.com.tw/g_get.php?" & "C" & Client_Card & gwsrcode & money & CkeckE & CkeckC
    ElseIf PayName="gwpay" Then
      get_close_url="https://gwpay.com.tw/close_get.php?" & "C" & Client_Card & gwsrcode & money & CkeckE & CkeckC
    End If
  Else
    get_close_url="3"
  End If
End Function
%>