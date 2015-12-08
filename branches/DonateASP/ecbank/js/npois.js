function column(columnname,columntype,essential,keywords,maxlen,msg){
  /*必要欄位判斷*/
  if(essential=='Y'){
    ck_blank=new Boolean(false);
    if(columntype=='checkbox'||columntype=='radio'){
      for(i=0;i<columnname.length;i++){
        if(columnname[i].checked){ 
          ck_blank=true;
          i=columnname.length;
        }
      }
    }else{
      if(columnname.value.replace(/(^\s*)|(\s*$)/g, "")!='') ck_blank=true;
    }
    if(ck_blank==false){
      if(columntype=='checkbox'||columntype=='radio'){
        alert(''+msg+' 欄位請選擇！');
      }else{
        alert(''+msg+' 欄位不可為空白！');
        columnname.focus();
      }
      return false;
    }
  }

  /*欄位長度判斷*/
  if(columnname.value!=''){
    if(Number(maxlen)>0){
      var itotal=0;
      if(columntype=='checkbox'||columntype=='radio'){
        for(i=0;i<columnname.length;i++){
    	    if(columnname[i].checked){
    		    for(var j=0;j<columnname[i].value.length;j++ ){
    		      if(escape(columnname[i].value.charAt(j)).length >= 4) itotal+=2;
              else itotal++;
    		    } 
    		    itotal+=2;
    	    }
        }
      }else{	
        for(var i=0;i<columnname.value.length;i++){
          if(escape(columnname.value.charAt(i)).length>=4) itotal+=2;
          else itotal++;
        }
      }
      if(Number(itotal)>Number(maxlen)){
        alert(''+msg+' 欄位長度超過限制！');
        if(columntype!='checkbox'&&columntype!='radio') columnname.focus();
        return false;
      }
    }    
  }
  
  /*欄位型態判斷*/
  if(columnname.value!=''){
    if(columntype=='number'||columntype=='verify'){
      if(isNaN(Number(columnname.value))==true){
        alert(''+msg+' 欄位必須為數值型態！');
        columnname.focus();
        return false;
      }
      if(columntype=='verify'){
        if(columnname.value.length!=4){
          alert(''+msg+'錯誤！');
          columnname.focus();
          return false;
        }
      }
    }else if(columntype=='email'){
      if(columnname.value.indexOf("@")==-1||columnname.value.indexOf(".")==-1||columnname.value.indexOf("@")==0||columnname.value.indexOf("@")==columnname.value.length-1||columnname.value.indexOf(".")==columnname.value.length-1){
        alert(''+msg+' 欄位格式錯誤！');
        columnname.focus();
        return false;
      }
  	  if(!(/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(columnname.value))){
        alert(''+msg+' 欄位格式錯誤！');
        columnname.focus();
        return false;
  	  }
  	  var iLast=columnname.value.split(".")[columnname.value.split(".").length-1].length;
      if(iLast!=2&&iLast!=3){
        alert(''+msg+' 欄位格式錯誤！');
        columnname.focus();
        return false;
      }
    }else if(columntype=='tel'){
      var checkkey='()-#*/';
      for(var i=0;i<columnname.value.length;i++){
    	  if(columnname.value.substr(i,1).toLowerCase()!=''){
    	    if(isNaN(Number(columnname.value.substr(i,1).toLowerCase()))==true){
    	      if(checkkey.indexOf(columnname.value.substr(i,1).toLowerCase())==-1){
              alert(''+msg+' 欄位格式錯誤！');
              columnname.focus();
              return false;
    	      }
    	    }
    	  }
      }
    }else if(columntype=='cell'){
      if(isNaN(Number(columnname.value))==true){
        alert(''+msg+' 欄位格式錯誤！(09XXXXXXXX)');
        columnname.focus();
        return false;
      }
      if(columnname.value.length!=10){
        alert(''+msg+' 欄位格式錯誤！(09XXXXXXXX)');
        columnname.focus();
        return false;
      }
      if(columnname.value.substr(0,2).toString()!='09'){
        alert(''+msg+' 欄位格式錯誤！(09XXXXXXXX)');
        columnname.focus();
        return false;
      }
    }else if(columntype=='url'){
      if(columnname.value.indexOf(".")==-1||columnname.value.indexOf(".")==1||columnname.value.indexOf(".")==columnname.value.length){
        alert(''+msg+' 欄位格式錯誤！');
        columnname.focus();
        return false;
      }
    }else if(columntype=='date'){
      if(columnname.value.indexOf("/")==-1||columnname.value.indexOf("/")==1||columnname.value.indexOf("/")==columnname.value.length){
        alert(''+msg+' 欄位格式錯誤！(西元年/月/日)');
        columnname.focus();
        return false;
      }
  	  Ary_Date=columnname.value.split("/");
  	  if(Ary_Date.length!=3) return false;
  	  for(i=0;i<3;i++){
  	    if(isNaN(Number(Ary_Date[i]))==true){
          alert(''+msg+' 欄位格式錯誤！(西元年/月/日)');
          columnname.focus();
          return false;
  	    }
  	  }
  	  if(Ary_Date[0].length!=4){
        alert(''+msg+' 欄位格式錯誤！(西元年/月/日)');
        columnname.focus();
        return false;
  	  }
  	  if(parseInt(Number(Ary_Date[0]))<1000){
        alert(''+msg+' 欄位格式錯誤！(西元年/月/日)');
        columnname.focus();
        return false;
  	  }
  	  if(parseInt(Number(Ary_Date[1]))<1||parseInt(Number(Ary_Date[1]))>12){
        alert(''+msg+' 欄位格式錯誤！(西元年/月/日)');
        columnname.focus();
        return false;
  	  }
  	  var YYYY=parseInt(Number(Ary_Date[0]))
  	  var MM=parseInt(Number(Ary_Date[1]))
  	  var DD=parseInt(Number(Ary_Date[2]))    	  
  	  if(MM==1||MM==3||MM==5||MM==7||MM==8||MM==10||MM==12){
  	    if(DD<1||DD>31) return false;
  	  }else if(MM==4||MM==6||MM==9||MM==11){
  	    if(DD<1||DD>30){
          alert(''+msg+' 欄位格式錯誤！(西元年/月/日)');
          columnname.focus();
          return false;
  	    }
  	  }else if(MM==2){
  	    if(leapyear(YYYY)){
  	      if(DD<1||DD>29){
            alert(''+msg+' 欄位格式錯誤！(西元年/月/日)');
            columnname.focus();
            return false;
  	      }
  	    }else{
  	      if(DD<1||DD>28){
            alert(''+msg+' 欄位格式錯誤！(西元年/月/日)');
            columnname.focus();
            return false;
  	      }
  	    }
  	  }else{
        alert(''+msg+' 欄位格式錯誤！(西元年/月/日)');
        columnname.focus();
        return false;
  	  }
    }else if(columntype=='idno'){
      if(columnname.value.length==10){	
        if(isNaN(Number(columnname.value.substring(0,1)))==false){
          alert(''+msg+' 欄位格式錯誤！');
          columnname.focus();
          return false;
        }
        if(isNaN(Number(columnname.value.substring(1,10)))==true){
          alert(''+msg+' 欄位格式錯誤！');
          columnname.focus();
          return false;
        }
        var alpha = columnname.value.substring(0,1).toUpperCase();
        var d1 = columnname.value.substring(1,2);
        var d2 = columnname.value.substring(2,3);
        var d3 = columnname.value.substring(3,4);
        var d4 = columnname.value.substring(4,5);
        var d5 = columnname.value.substring(5,6);
        var d6 = columnname.value.substring(6,7);
        var d7 = columnname.value.substring(7,8);
        var d8 = columnname.value.substring(8,9);
        var d9 = columnname.value.substring(9,10);
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
        if(acc=='0'){
          alert(''+msg+' 欄位格式錯誤！');
          columnname.focus();
          return false;
        }
        var checksum = acc+8*d1+7*d2+6*d3+5*d4+4*d5+3*d6+2*d7+1*d8+1*d9;
        var check1 = parseInt(checksum/10);
        var check2 = checksum/10;
        var check3 = (check2-check1)*10;
        if((checksum!=check1*10)&&(d9!=(10-check3))){
          alert(''+msg+' 欄位格式錯誤！');
          columnname.focus();
          return false;
        }
      }else{
         alert(''+msg+' 欄位格式錯誤！');
        columnname.focus();
        return false;
      }
    }
  }
  
  /*關鍵字判斷*/
  if(keywords=='Y'&&columntype!='checkbox'&&columntype!='radio'){
  	if(columnname.value!=''){
      var AryKey = new Array('.asp','.bak','.cfm','.css','.dos','.htm','.inc','.ini','.js','.php','.txt','/*','*/','.@','@.','@@','@C','@T','@P','a href','acunetix','alert','application','begin','cast','cache','config','char','chr','cookie','count','create','css','cursor','deallocate','declare','delete','dir','drop','echo','end','eval','exec','exists','execute','fetch','from','hidden','iframe','insert','into','is_','join','js','kill','left','manage','master','mid','nchar','ntext','nvarchar','open','password','right','script','set','select','session','src','sys','sysobjects','syscolumns','table','text','truncate','url','user','update','varchar','where','while');
      for(var i=0;i<=AryKey.length-1;i++){
        if(columnname.value.indexOf(AryKey[i])!=-1){
          alert(''+msg+' 欄位請勿輸入保留字 [ '+AryKey[i]+' ]');
          columnname.focus();
          return false;
        }
      }
    }  
  }
  return true;
}

/*潤年判斷*/
function leapyear(year){
  if(parseInt(year)%4==0){
   if(parseInt(year)%100==0){
    if(parseInt(year)%400==0) return true;
    	return false;	
   }else return true;
  }else return false;
}

/*副檔名判斷*/
function column_ext(columnname,exttype,extother){
  ck_ext=new Boolean(false);
  if(columnname.value.length>4){
    if(columnname.value.indexOf('.')!=-1&&columnname.value.indexOf('.')!=1&&columnname.value.indexOf('.')!=columnname.value.length){
      var extname=columnname.value.substr(columnname.value.lastIndexOf('.')+1,columnname.value.length).toLowerCase();
      if(extother!=''){
        var strext=extother;
      }else if(exttype!=''){
        if(exttype=='img'){
          var strext='jpg,gif';
        }else	if(exttype=='doc'){
          var strext='doc,pdf,xls';
        }else	if(exttype=='wmv'){
          var strext='wmv,asf,avi,mpg';
        }else	if(exttype=='swf'){
          var strext='swf';
        }else	if(exttype=='flv'){
          var strext='flv';
        }else{
          var strext=extother;
        }
      }
      if(strext!=''){
        var aryext=strext.split(',');
        for(var i=0;i<=aryext.length-1;i++){
          if(aryext[i]==extname){ 
            ck_ext=true;
            i=aryext.length-1;
          }
        }
      }
    }
  }
  if(ck_ext==false){
    alert('上傳文件格式非系統指定！');
    columnname.focus();
    return false;
  }else return true;
}

/*submit*/
function act_submit(formname,act_value,confirm_msg){
  if(confirm_msg!==''){
    if(confirm('您是否確定要'+confirm_msg+'？')){
      formname.action.value=act_value;
      formname.submit();
    }
  }else{
    if(act_value!='') formname.action.value=act_value;
    formname.submit();
  }
}

/*縣市鄉鎮連動選單*/
function ChgCity(CityCode,AreaObj,ZipObj){
  ZipObj.value='';
  ClearOption(AreaObj);
  AreaObj.options[0] = new Option('鄉鎮市區','');
  var AryAreaCode = new Array('A',['100','中正區','103','大同區','104','中山區','105','松山區','106','大安區','108','萬華區','110','信義區','111','士林區','112','北投區','114','內湖區','115','南港區','116','文山區'],'B',['207','萬里區','208','金山區','220','板橋區','221','汐止區','222','深坑區','223','石碇區','224','瑞芳區','226','平溪區','227','雙溪區','228','貢寮區','231','新店區','232','坪林區','233','烏來區','234','永和區','235','中和區','236','土城區','237','三峽區','238','樹林區','239','鶯歌區','241','三重區','242','新莊區','243','泰山區','244','林口區','247','蘆洲區','248','五股區','249','八里區','251','淡水區','252','三芝區','253','石門區'],'C',['200','仁愛區','201','信義區','202','中正區','203','中山區','204','安樂區','205','暖暖區','206','七堵區'],'D',['260','宜蘭市','261','頭城鎮','262','礁溪鄉','263','壯圍鄉','264','員山鄉','265','羅東鎮','266','三星鄉','267','大同鄉','268','五結鄉','269','冬山鄉','270','蘇澳鎮','272','南澳鄉','290','釣魚台'],'E',['320','中壢市','324','平鎮市','325','龍潭鄉','326','楊梅市','327','新屋鄉','328','觀音鄉','330','桃園市','333','龜山鄉','334','八德市','335','大溪鎮','336','復興鄉','337','大園鄉','338','蘆竹鄉'],'F',['300','東區','3001','北區','3002','香山區'],'G',['302','竹北市','303','湖口鄉','304','新豐鄉','305','新埔鎮','306','關西鎮','307','芎林鄉','308','寶山鄉','310','竹東鎮','311','五峰鄉','312','橫山鄉','313','尖石鄉','314','北埔鄉','315','峨眉鄉'],'H',['350','竹南鎮','351','頭份鎮','352','三灣鄉','353','南庄鄉','354','獅潭鄉','356','後龍鎮','357','通霄鎮','358','苑裡鎮','360','苗栗市','361','造橋鄉','362','頭屋鄉','363','公館鄉','364','大湖鄉','365','泰安鄉','366','銅鑼鄉','367','三義鄉','368','西湖鄉','369','卓蘭鎮'],'I',['400','中區','401','東區','402','南區','403','西區','404','北區','406','北屯區','407','西屯區','408','南屯區','411','太平區','412','大里區','413','霧峰區','414','烏日區','420','豐原區','421','后里區','422','石岡區','423','東勢區','424','和平區','426','新社區','427','潭子區','428','大雅區','429','神岡區','432','大肚區','433','沙鹿區','434','龍井區','435','梧棲區','436','清水區','437','大甲區','438','外埔區','439','大安區'],'K',['500','彰化市','502','芬園鄉','503','花壇鄉','504','秀水鄉','505','鹿港鎮','506','福興鄉','507','線西鄉','508','和美鎮','509','伸港鄉','510','員林鎮','511','社頭鄉','512','永靖鄉','513','埔心鄉','514','溪湖鎮','515','大村鄉','516','埔鹽鄉','520','田中鎮','521','北斗鎮','522','田尾鄉','523','埤頭鄉','524','溪州鄉','525','竹塘鄉','526','二林鎮','527','大城鄉','528','芳苑鄉','530','二水鄉'],'L',['540','南投市','541','中寮鄉','542','草屯鎮','544','國姓鄉','545','埔里鎮','546','仁愛鄉','551','名間鄉','552','集集鎮','553','水里鄉','555','魚池鄉','556','信義鄉','557','竹山鎮','558','鹿谷鄉'],'M',['630','斗南鎮','631','大埤鄉','632','虎尾鎮','633','土庫鎮','634','褒忠鄉','635','東勢鄉','636','台西鄉','637','崙背鄉','638','麥寮鄉','640','斗六市','643','林內鄉','646','古坑鄉','647','莿桐鄉','648','西螺鎮','649','二崙鄉','651','北港鎮','652','水林鄉','653','口湖鄉','654','四湖鄉','655','元長鄉'],'N',['600','東區','6001','西區'],'O',['602','番路鄉','603','梅山鄉','604','竹崎鄉','605','阿里山鄉','606','中埔鄉','607','大埔鄉','608','水上鄉','611','鹿草鄉','612','太保市','613','朴子市','614','東石鄉','615','六腳鄉','616','新港鄉','621','民雄鄉','622','大林鎮','623','溪口鄉','624','義竹鄉','625','布袋鎮'],'P',['700','中西區','701','東區','702','南區','704','北區','708','安平區','709','安南區','710','永康區','711','歸仁區','712','新化區','713','左鎮區','714','玉井區','715','楠西區','716','南化區','717','仁德區','718','關廟區','719','龍崎區','720','官田區','721','麻豆區','722','佳里區','723','西港區','724','七股區','725','將軍區','726','學甲區','727','北門區','730','新營區','731','後壁區','732','白河區','733','東山區','734','六甲區','735','下營區','736','柳營區','737','鹽水區','741','善化區','742','大內區','743','山上區','744','新市區','745','安定區'],'R',['800','新興區','801','前金區','802','苓雅區','803','鹽埕區','804','鼓山區','805','旗津區','806','前鎮區','807','三民區','811','楠梓區','812','小港區','813','左營區','814','仁武區','815','大社區','817','東沙群島','819','南沙群島','820','岡山區','821','路竹區','822','阿蓮區','823','田寮區','824','燕巢區','825','橋頭區','826','梓官區','827','彌陀區','828','永安區','829','湖內區','830','鳳山區','831','大寮區','832','林園區','833','鳥松區','840','大樹區','842','旗山區','843','美濃區','844','六龜區','845','內門區','846','杉林區','847','甲仙區','848','桃源區','849','那瑪夏區','851','茂林區','852','茄萣區'],'T',['900','屏東市','901','三地門鄉','902','霧台鄉','903','瑪家鄉','904','九如鄉','905','里港鄉','906','高樹鄉','907','鹽埔鄉','908','長治鄉','909','麟洛鄉','911','竹田鄉','912','內埔鄉','913','萬丹鄉','920','潮州鎮','921','泰武鄉','922','來義鄉','923','萬巒鄉','924','崁頂鄉','925','新埤鄉','926','南州鄉','927','林邊鄉','928','東港鎮','929','琉球鄉','931','佳冬鄉','932','新園鄉','940','枋寮鄉','941','枋山鄉','942','春日鄉','943','獅子鄉','944','車城鄉','945','牡丹鄉','946','恆春鎮','947','滿州鄉'],'U',['950','台東市','951','綠島鄉','952','蘭嶼鄉','953','延平鄉','954','卑南鄉','955','鹿野鄉','956','關山鎮','957','海端鄉','958','池上鄉','959','東河鄉','961','成功鎮','962','長濱鄉','963','太麻里鄉','964','金峰鄉','965','大武鄉','966','達仁鄉'],'V',['970','花蓮市','971','新城鄉','972','秀林鄉','973','吉安鄉','974','壽豐鄉','975','鳳林鎮','976','光復鄉','977','豐濱鄉','978','瑞穗鄉','979','萬榮鄉','981','玉里鎮','982','卓溪鄉','983','富里鄉'],'W',['880','馬公市','881','西嶼鄉','882','望安鄉','883','七美鄉','884','白沙鄉','885','湖西鄉'],'X',['890','金沙鎮','891','金湖鎮','892','金寧鄉','893','金城鎮','894','烈嶼鄉','896','烏坵鄉'],'Y',['209','南竿鄉','210','北竿鄉','211','莒光鄉','212','東引鄉'])
  for(var i=0;i<=AryAreaCode.length-1;i=i+2){
    if(AryAreaCode[i]==CityCode){
      var theArea = AryAreaCode[i+1];
      var k=1;
      for(var j=0;j<=theArea.length-1;j=j+2){
        AreaObj.options[k] = new Option(theArea[j+1],theArea[j]);
        k++;
      }
    }
  }
}
function ClearOption(SelObj){
  var i;
  for(i=SelObj.length-1;i>=0;i--){
    SelObj.options[i] = null;
  }
}
function ChgArea(AreaCode,ZipObj){
  ZipObj.value=AreaCode.substring(0,3);
}
function ChgAddress(AddressObj){
  if(AddressObj=='請輸入路段號樓'){
    AddressObj='';
  }
}