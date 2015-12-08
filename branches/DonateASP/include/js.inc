<%
Function CheckFileJ(Form,FName,FSpace,Ftype,FLen,ListField)
  If FSpace="Y" Then
    If Ftype="checkbox" Then
      Response.Write "if(CheckBox(document." & Form & "." & FName & ")==false){" & Chr(13) & Chr(10)
      Response.Write "  alert('"& ListField & "  欄位不可為空白！');" & Chr(13) & Chr(10)
      Response.Write "  return;" & Chr(13) & Chr(10)
      Response.Write "}" & Chr(13) & Chr(10)
    ElseIf Ftype="radio" Then
      Response.Write "if(Radio(document." & Form & "." & FName & ")==false){" & Chr(13) & Chr(10)
      Response.Write "  alert('"& ListField & "  欄位不可為空白！');" & Chr(13) & Chr(10)
      Response.Write "  return;" & Chr(13) & Chr(10)
      Response.Write "}" & Chr(13) & Chr(10)      
    Else
      Response.Write "if(document." & Form & "." & FName & ".value==''){" & Chr(13) & Chr(10)
      Response.Write "  alert('"& ListField & "  欄位不可為空白！');" & Chr(13) & Chr(10)
      Response.Write "  document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
      Response.Write "  return;" & Chr(13) & Chr(10)
      Response.Write "}" & Chr(13) & Chr(10)
    End If
  End If
  If Cint(FLen)>0 Then
    If Ftype="checkbox" Then
      Response.Write "if(CheckBoxLen(document." & Form & "." & FName & ","&FLen&")==false){" & Chr(13) & Chr(10)
      Response.Write "  alert('"& ListField & "  欄位長度超過限制！');" & Chr(13) & Chr(10)
      Response.Write "  return;" & Chr(13) & Chr(10)
      Response.Write "}" & Chr(13) & Chr(10)
    ElseIf Ftype="radio" Then
      Response.Write "if(RadioLen(document." & Form & "." & FName & ","&FLen&")==false){" & Chr(13) & Chr(10)
      Response.Write "  alert('"& ListField & "  欄位長度超過限制！');" & Chr(13) & Chr(10)
      Response.Write "  return;" & Chr(13) & Chr(10)
      Response.Write "}" & Chr(13) & Chr(10)            
    Else
      Response.Write "if(Len(document." & Form & "." & FName & ".value,"&FLen&")==false){" & Chr(13) & Chr(10)
      Response.Write "  alert('"& ListField & "  欄位長度超過限制！');" & Chr(13) & Chr(10)
      Response.Write "  document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
      Response.Write "  return;" & Chr(13) & Chr(10)
      Response.Write "}" & Chr(13) & Chr(10)
    End If
  End If
  If Ftype<>"checkbox" And Ftype<>"radio" Then
    Response.Write "if(document." & Form & "." & FName & ".value!=''){" & Chr(13) & Chr(10)
    Response.Write "  if(Column(document." & Form & "." & FName & ".value,"""&Ftype&""")==false){" & Chr(13) & Chr(10)
    If Ftype="date" Then
      Response.Write "  alert('"& ListField & "  欄位格式錯誤(西元年/月/日)！');" & Chr(13) & Chr(10)
    Else
      Response.Write "  alert('"& ListField & "  欄位格式錯誤！');" & Chr(13) & Chr(10)
    End If
    Response.Write "    document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
    Response.Write "    return;" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    Response.Write "  if(KeyWord(document." & Form & "." & FName & ".value,"""&Ftype&""")==false){" & Chr(13) & Chr(10)
    Response.Write "    alert('"& ListField & "  欄位請勿使用特殊字元！');" & Chr(13) & Chr(10)
    Response.Write "    document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
    Response.Write "    return;" & Chr(13) & Chr(10)
    Response.Write "  }" & Chr(13) & Chr(10)
    If Ftype<>"email" Then
      Response.Write "  if(KeepWord(document." & Form & "." & FName & ".value)==false){" & Chr(13) & Chr(10)
      Response.Write "    alert('"& ListField & "  欄位請勿使用保留字串！');" & Chr(13) & Chr(10)
      Response.Write "    document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
      Response.Write "    return;" & Chr(13) & Chr(10)
      Response.Write "  }" & Chr(13) & Chr(10)      
    End If
    Response.Write "}" & Chr(13) & Chr(10)
  End If
End Function

Function CheckExtJ (Form,FName,FSpace,Ftype,ListField)
  If FSpace="Y" Then
    Response.Write "if(document." & Form & "." & FName & ".value==''){" & Chr(13) & Chr(10)
    Response.Write "  alert('"& ListField & "  欄位不可為空白！');" & Chr(13) & Chr(10)
    Response.Write "  document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
    Response.Write "  return;" & Chr(13) & Chr(10)
    Response.Write "}" & Chr(13) & Chr(10)
  End If
  Response.Write "if(document." & Form & "." & FName & ".value!=''){" & Chr(13) & Chr(10)
  If Ftype="img" Or Ftype="doc" Or Ftype="wmv" Or Ftype="flash" Or Ftype="flv" Then
    Response.Write "if(Ext(document." & Form & "." & FName & ".value,"""&Ftype&""")==false){" & Chr(13) & Chr(10)
  Else
    Response.Write "if(ExtDIY(document." & Form & "." & FName & ".value,"""&Ftype&""")==false){" & Chr(13) & Chr(10)
  End If
  Response.Write "    alert('"& ListField & "  此檔案系統不支援！');" & Chr(13) & Chr(10)
  Response.Write "    document." & Form & "." & FName & ".focus();" & Chr(13) & Chr(10)
  Response.Write "    return;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "}" & Chr(13) & Chr(10)
End Function

Function RadioOtherJ(Form,FName,FName_Other,FLen,ListField)
  Response.Write "if(RadioOther(document." & Form & "." & FName & ")){" & Chr(13) & Chr(10)
  Response.Write "  if(document." & Form & "." & FName_Other & ".value==''){" & Chr(13) & Chr(10)
  Response.Write "    alert('"& ListField & "  欄位不可為空白！');" & Chr(13) & Chr(10)
  Response.Write "    document." & Form & "." & FName_Other & ".focus();" & Chr(13) & Chr(10)
  Response.Write "    return;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "  if(CheckBoxLen(document." & Form & "." & FName_Other & ","&FLen&")==false){" & Chr(13) & Chr(10)
  Response.Write "    alert('"& ListField & "  欄位長度超過限制！');" & Chr(13) & Chr(10)
  Response.Write "    return;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)  
  Response.Write "}" & Chr(13) & Chr(10) 
End Function

Function CheckboxOtherJ(Form,FName,FName_Other,FLen,ListField)
  Response.Write "if(CheckboxOther(document." & Form & "." & FName & ")){" & Chr(13) & Chr(10)
  Response.Write "  if(document." & Form & "." & FName_Other & ".value==''){" & Chr(13) & Chr(10)
  Response.Write "    alert('"& ListField & "  欄位不可為空白！');" & Chr(13) & Chr(10)
  Response.Write "    document." & Form & "." & FName_Other & ".focus();" & Chr(13) & Chr(10)
  Response.Write "    return;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)
  Response.Write "  if(CheckBoxLen(document." & Form & "." & FName_Other & ","&FLen&")==false){" & Chr(13) & Chr(10)
  Response.Write "    alert('"& ListField & "  欄位長度超過限制！');" & Chr(13) & Chr(10)
  Response.Write "    return;" & Chr(13) & Chr(10)
  Response.Write "  }" & Chr(13) & Chr(10)  
  Response.Write "}" & Chr(13) & Chr(10) 
End Function

Function Check_Post()
  Check_Post = True
  URL1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
  URL2 = Cstr(Request.ServerVariables("SERVER_NAME"))
  If Mid(URL1,8,Len(URL2))<>URL2 Then Check_Post = False
End Function
%>
<Script language="JavaScript"><!--
  function Len(FileNmame,MaxLen){
    var cnt=0;
    for(var i=0;i<FileNmame.length;i++ ){
      if(escape(FileNmame.charAt(i)).length >= 4) cnt+=2;
      else cnt++;
    }
    if(cnt>MaxLen) return false;
    else return true;
  }
  function CLen(FileNmame,MaxLen){
    var cnt=0;
    for(var i=0;i<FileNmame.length;i++ ){
      if(escape(FileNmame.charAt(i)).length >= 4) cnt+=2;
    }
    if(cnt>MaxLen) return true;
    else return false;
  }   
  function Column(FileNmame,ColumnType){
  	if(ColumnType=='number'){
  		if(isNaN(Number(FileNmame))==true) return false;
  	}else if(ColumnType=='verify'){
  		if(isNaN(Number(FileNmame))==true) return false;
  		if(FileNmame.length!=4) return false;
  	}else if(ColumnType=='email'){
  	  if(FileNmame.indexOf("@")==-1||FileNmame.indexOf(".")==-1||FileNmame.indexOf("@")==1||FileNmame.indexOf(".")==1||FileNmame.indexOf("@")==FileNmame.length||FileNmame.indexOf(".")==FileNmame.length) return false;
  	  if(!(/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(FileNmame))) return false;
  	  var iLast=FileNmame.split(".")[FileNmame.split(".").length-1].length;   
      if(iLast!=2&&iLast!=3) return false; 
  	}else if(ColumnType=='url'){
  	  if(FileNmame.indexOf(".")==-1||FileNmame.indexOf(".")==1||FileNmame.indexOf(".")==FileNmame.length) return false;
  	}else if(ColumnType=='date'){
  	  if(FileNmame.indexOf("/")==-1||FileNmame.indexOf("/")==1||FileNmame.indexOf("/")==FileNmame.length) return false;
  	  Ary_Date=FileNmame.split("/");
  	  if(Ary_Date.length!=3) return false;
  	  for(i=0;i<3;i++){
  	  	if(isNaN(Number(Ary_Date[i]))==true) return false;
  	  }
  	  if(Ary_Date[0].length!=4) return false;
  	  if(parseInt(Number(Ary_Date[0]))<1000) return false;
  	  if(parseInt(Number(Ary_Date[1]))<1||parseInt(Number(Ary_Date[1]))>12) return false;
  	  var YYYY=parseInt(Number(Ary_Date[0]))
  	  var MM=parseInt(Number(Ary_Date[1]))
  	  var DD=parseInt(Number(Ary_Date[2]))
  	    	  
  	  if(MM==1||MM==3||MM==5||MM==7||MM==8||MM==10||MM==12){
  	  	if(DD<1||DD>31) return false;
  	  }else if(MM==4||MM==6||MM==9||MM==11){
  	    if(DD<1||DD>30) return false;
  	  }else if(MM==2){
  	    if(LeapYear(YYYY)){
  	    	if(DD<1||DD>29) return false;
  	    }else{
  	      if(DD<1||DD>28) return false;
  	    }
  	  }else return false;
  	}else if(ColumnType=='idno'){
      if(FileNmame.length==8){
      	if(isNaN(Number(FileNmame))==true) return false;
      }else if(FileNmame.length==10){	
        if(isNaN(Number(FileNmame.substring(0,1)))==false) return false;
        if(isNaN(Number(FileNmame.substring(1,10)))==true) return false;
        var alpha = FileNmame.substring(0,1).toUpperCase();
        var d1 = FileNmame.substring(1,2);
        var d2 = FileNmame.substring(2,3);
        var d3 = FileNmame.substring(3,4);
        var d4 = FileNmame.substring(4,5);
        var d5 = FileNmame.substring(5,6);
        var d6 = FileNmame.substring(6,7);
        var d7 = FileNmame.substring(7,8);
        var d8 = FileNmame.substring(8,9);
        var d9 = FileNmame.substring(9,10);
        var acc=0;
        switch(alpha){
          case "A":
            acc=1;
            break;
          case "B":
            acc=10;
            break;
          case "C":
            acc=19;
            break;
          case "D":
            acc=28;
            break;
          case "E":
            acc=37;
            break;
          case "F":
            acc=46;
            break;
          case "G":
            acc=55;
            break;
          case "H":
            acc=64;
            break;
          case "I":
            acc=39;
            break;
          case "J":
            acc=73;
            break;
          case "K":
            acc=82;
            break;
          case "L":
            acc=2;
            break;
          case "M":
            acc=11;
            break;
          case "N":
            acc=20;
            break;
          case "O":
            acc=48;
            break;
          case "P":
            acc=29;
            break;
          case "Q":
            acc=38;
            break;
          case "R":
            acc=47;
            break;
          case "S":
            acc=56;
            break;
          case "T":
            acc=65;
            break; 
          case "U":
            acc=74;
            break;
          case "V":
            acc=83;
            break;
          case "W":
            acc=21;
            break;
          case "X":
            acc=3;
            break;
          case "Y":
            acc=12;
            break;
          case "Z":
            acc=30;
            break;
        }
        if(acc=='0') return false;
        var checksum = acc+8*d1+7*d2+6*d3+5*d4+4*d5+3*d6+2*d7+1*d8+1*d9;
        var check1 = parseInt(checksum/10);
        var check2 = checksum/10;
        var check3 = (check2-check1)*10;
        if((checksum!=check1*10)&&(d9!=(10-check3))) return false;
      }else return false;
    }
  	return true;
  }  
  function KeyWord(FileNmame,KeyWordType){
    var keyword=FileNmame.toLowerCase();
    if(KeyWordType=='content'){
    	var AryKey = new Array('1=1','a=a','--','./');
    }else if(KeyWordType=='tel'){
      var AryKey = new Array('`','~','!','$','%','^','&','*','+','=','[',']','{','}','\\','|',';',':','\'','""','<','>','//','1=1','a=a','--',';','./');
    }else if(KeyWordType=='url'){
      var AryKey = new Array('`','~','!','#','$','%','^','*','(',')','+','[',']','{','}','\\','|',';',':','\'','""','<','>','//','1=1','a=a','--',';','./');
    }else if(KeyWordType=='money'){
      var AryKey = new Array('`','~','!','#','%','^','&','*','(',')','+','=','[',']','{','}','\\','|',';',':','\'','""','<','>','//','1=1','a=a','--',';','./');
    }else{
      var AryKey = new Array('`','~','!','#','$','%','^','&','*','(',')','+','=','[',']','{','}','\\','|',';',':','\'','"','<','>','//','1=1','a=a','--',';','./');
    }
    for(var i=0;i<=AryKey.length-1;i++){
      if(keyword.indexOf(AryKey[i])!=-1) return false;
    }
    return true;
  }    
  function KeepWord(FileNmame){
    var keyword=FileNmame.toLowerCase();
    var AryKey = new Array('.asp','.bak','.cfm','.css','.dos','.htm','.inc','.ini','.js','.php','.txt','/*','*/','.@','@.','@@','@C','@T','@P','a href','admin','acunetix','alert','and','application','begin','cast','cache','config','char','chr','cookie','count','create','css','cursor','deallocate','declare','delete','dir','drop','echo','end','eval','exec','exists','execute','fetch','from','hidden','iframe','insert','into','is_','join','js','kill','left','manage','master','mid','nchar','ntext','nvarchar','open','password','right','script','set','select','session','src','sys','sysobjects','syscolumns','table','text','truncate','url','user','update','varchar','where','while','xtype','%2527','address.tst');
    for(var i=0;i<=AryKey.length-1;i++){
      if(keyword.indexOf(AryKey[i])!=-1) return false;
    }
    return true;
  }
  function CheckBox(FileNmame){
    for(i=0;i<FileNmame.length;i++){
      if(FileNmame[i].checked) return true;
    }
    return false;
  }
  function CheckBoxLen(FileNmame,MaxLen){
    var cnt=0;
    for(i=0;i<FileNmame.length;i++){
    	if(FileNmame[i].checked){
    		for(var j=0;j<FileNmame[i].value.length;j++ ){
    		  if(escape(FileNmame[i].value.charAt(j)).length >= 4) cnt+=2;
          else cnt++;
    		}
    		cnt+=2;
    	}
    }
    if(cnt>MaxLen) return false;
    else return true;
  }
  function CheckboxOther(FileNmame){
    for(i=0;i<FileNmame.length;i++){
    	if(FileNmame[i].value=='其他'&&FileNmame[i].checked||FileNmame[i].value=='其它'&&FileNmame[i].checked) return true;
    }
    return false;
  }
  function Radio(FileNmame){
    for(i=0;i<FileNmame.length;i++){
      if(FileNmame[i].checked) return true;
    }
    return false;
  }
  function RadioLen(FileNmame,MaxLen){
    var cnt=0;
    for(i=0;i<FileNmame.length;i++){
    	if(FileNmame[i].checked){
    		for(var j=0;j<FileNmame[i].value.length;j++ ){
    		  if(escape(FileNmame[i].value.charAt(j)).length >= 4) cnt+=2;
          else cnt++;
    		}
    	}
    }
    if(cnt>MaxLen) return false;
    else return true;
  }
  function RadioOther(FileNmame){
    for(i=0;i<FileNmame.length;i++){
    	if(FileNmame[i].value=='其他'&&FileNmame[i].checked) return true;
    }
    return false;
  }  
  function Ext(FileNmame,ExtType){
    if(FileNmame.length>4){
      if(FileNmame.indexOf(".")==-1||FileNmame.indexOf(".")==1||FileNmame.indexOf(".")==FileNmame.length) return false;
      var ExtName=FileNmame.substr(FileNmame.lastIndexOf('.')+1,FileNmame.length).toLowerCase();
      if(ExtType=='img'){
      	var AllExt='jpg,gif';
      }else	if(ExtType=='doc'){
        var AllExt='doc,pdf,xls';
      }else	if(ExtType=='wmv'){
        var AllExt='wmv,asf,avi,mpg';
      }else	if(ExtType=='flash'){
        var AllExt='swf';
      }else	if(ExtType=='flv'){
        var AllExt='flv';
      }
      var Ary_Ext=AllExt.split(',');
      for(var i=0;i<=Ary_Ext.length-1;i++){
        if(Ary_Ext[i]==ExtName) return true;
      }
      return false;
    }else return false;
  }    
  function ExtDIY(FileNmame,ExtType){
    if(FileNmame.length>4){
      if(FileNmame.indexOf(".")==-1||FileNmame.indexOf(".")==1||FileNmame.indexOf(".")==FileNmame.length) return false;
      var ExtName=FileNmame.substr(FileNmame.lastIndexOf('.')+1,FileNmame.length).toLowerCase();
      var Ary_Ext=ExtType.split(',');
      for(var i=0;i<=Ary_Ext.length-1;i++){
        if(Ary_Ext[i]==ExtName) return true;
      }
      return false;
    }else return false;
  }
  function LeapYear(Year){
    if(parseInt(Year)%4==0){
    	if(parseInt(Year)%100==0){
    		if(parseInt(Year)%400==0) return true;
    	  return false;	
      }else return true;
    }else return false;
  }
  function FormatNumber(Str,Dot,Round){ 
    var ReturnStr='';
    if(isNaN(Number(Str))==false&&isNaN(Number(Dot))==false){
      var Str_Dot='';
      if(Round=='Y'){
        if(parseInt(Dot)>0){
          for(var i=0;i<=parseInt(Dot);i++){
            if(i==parseInt(Dot)){
              Str_Dot=Str_Dot+'5';
            }else{
              Str_Dot=Str_Dot+'0';
            }
          }
          Str_Dot='0.'+Str_Dot;
        }else Str_Dot='0.5';    
      }else Str_Dot='0.0';
      var Ary_Str=(Number(Str)+Number(Str_Dot)).toString().split('.');
      var iMod=0;
      if(Ary_Str[0].length>3){
        iMod=Ary_Str[0].length%3;
        if(iMod!=0) ReturnStr=Ary_Str[0].substring(0,iMod)+',';
      }
      var j=0;
      for(var i=iMod;i<Ary_Str[0].length;i++){
        ReturnStr=ReturnStr+Ary_Str[0].substring(i,i+1); 
        j++;
        if((j+3)%3==0&&i!=Ary_Str[0].length-1) ReturnStr=ReturnStr+',';
      }    
      if(parseInt(Dot)>0&&Ary_Str.length>1){
        ReturnStr=ReturnStr+'.';
        for(var i=0;i<Ary_Str[1].length;i++){
          if(i<parseInt(Dot)) ReturnStr=ReturnStr+Ary_Str[1].substring(i,i+1);
        }
      }
    }
    return ReturnStr;
  }   
--></Script>