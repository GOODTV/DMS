//耀溘逄晟婦宒腔潠楛蛌遙髡夔脣璃ㄐ
//Edited by Stardy --2005-04-16 , Web :http://www.stardy.com , QQ:2885465

var Default_isFT = 0		//蘇岆瘁楛极ㄛ0-潠极ㄛ1-楛极
var StranIt_Delay = 10 //楹祒晊奀瑭鏃ㄗ扢涴跺腔醴腔岆厙珜珂霜釧腔珆珋堤懂ㄘ

//ㄜㄜㄜㄜㄜㄜㄜ測鎢羲宎ㄛ眕狟梗蜊ㄜㄜㄜㄜㄜㄜㄜ
//蛌遙恅掛
function StranText(txt,toFT,chgTxt)
{
	if(txt==""||txt==null)return ""
	toFT=toFT==null?BodyIsFt:toFT
	if(chgTxt)txt=txt.replace((toFT?"潠":"楛"),(toFT?"楛":"潠"))
	if(toFT){return Traditionalized(txt)}
	else {return Simplized(txt)}
}
//蛌遙勤砓ㄛ妏蚚菰寥ㄛ紨脯婀善恅掛
function StranBody(fobj)
{
	if(typeof(fobj)=="object"){var obj=fobj.childNodes}
	else 
	{
		var tmptxt=StranLink_Obj.innerHTML.toString()
		if(tmptxt.indexOf("潠")<0)
		{
			BodyIsFt=1
			StranLink_Obj.innerHTML=StranText(tmptxt,0,1)
			StranLink.title=StranText(StranLink.title,0,1)
		}
		else
		{
			BodyIsFt=0
			StranLink_Obj.innerHTML=StranText(tmptxt,1,1)
			StranLink.title=StranText(StranLink.title,1,1)
		}
		setCookie(JF_cn,BodyIsFt,7)
		var obj=document.body.childNodes
	}
	for(var i=0;i<obj.length;i++)
	{
		var OO=obj.item(i)
		if("||BR|HR|TEXTAREA|".indexOf("|"+OO.tagName+"|")>0||OO==StranLink_Obj)continue;
		if(OO.title!=""&&OO.title!=null)OO.title=StranText(OO.title);
		if(OO.alt!=""&&OO.alt!=null)OO.alt=StranText(OO.alt);
		if(OO.tagName=="INPUT"&&OO.value!=""&&OO.type!="text"&&OO.type!="hidden")OO.value=StranText(OO.value);
		if(OO.nodeType==3){OO.data=StranText(OO.data)}
		else StranBody(OO)
	}
}
function JTPYStr()
{
	return '馬高鬼乾倏偯兜商啦啊啖唬域堅堆堂夠娶婀悼惘惆惚捲探接捧掘措掄捩救教敕晚晤晨曹梁梯梱械梭梆梅梔欲畢異痊眶眺硃統紮紹紼絀細紳組累終翎耜聊聆脯莎莢莖莽莫莒莊莓莉莠荷部郭酗野釵釦釧陵陬章鳥麻傍傅備凱剴勞博喧喊喝喘喋喳單唾喻喬喉喙堯堰報堝堠插揣揖揭換敞斑斐晴曾朝棗棵棹棲棣棋植椒椎欽渣湛湍湃童等策筆筐筋筑粟絞結絨絕紫絲絢絰絳聒腑腌菩萍菰萌萸菜萇蛭袱覃詞証詁隊階陽隆陲雄集雲須馭黃黍傭傲傯剿剷募勦勤勣嗨嗦嗤塑塗塚填塒嫁嫌媾楚楷楠概楨楝楣歲毓毽溯溼溺滄煙煤煌煥煖猷獅猿瑯瑟瑞瑁瑛瑜當瘀痺盞葷落董葩蛹蜈蜀蛾蜆裟裔補裒解該詳試詩詼詣詭訾賊賈跨跳跺跤麂僧僥僚像僱凳劃劂嘍嘈塵壽夤奩嫦嫗嫘寞寡寥實寨寢察嶄幛幣幕幗廓弊徹漣澈犖碳禍種筵箔箄綻綴綿誘誚誧貍貌賒赫趕輒輓辣遠遜遣遙遢遝鄙鄘酴銑閨閩障鞅韶頗領颯颱餃餅餌骰鳴鳳億儀僻儂儅慰憧憐憎憬憤憮撰撥撓撒撙撳數暮暱樟槨標樊槳樂槭歎毅漿潼澄潦潔潸澎潺潰澗潘滕潯潠潟熬熱熨牖犛獎獗瑩耦膛膜膝膠膚蓮蔓蔑蔣蔡蓬蝶蝸蝨蝙蝌蝓衛衝褐褓褕褊諄誕調論誹墨儐冪勳噙噩噤噸噥噱壁學導戰撼擂撾曉曄暸橫豌賞賦趟趣踡踞躺輝輛輩輜鴃麩麾橇樵機橈氅濂縞羲翮衡褲諺謀諜諭踱踴蹂輸辨遵鄴錠錶鋸錯錢嚏壓孺屨嶸幫應懂懇懦戲戴擎擊擘擠擰擦擬擱擢擭斂檜櫛殮氈濱濫濬濡燦燭爵牆獰璨癆療癌盪薑薔薨薊虧蟀蟑蟒蟆螫螻螺蟈蟋褻褶襄褸謗謙謝谿賽趨轂還邁邂鄹醣鍵鍊錘鍾鍛鍰闆隸韓顆颶餵瞽瞿瞻瞼礎禮穡竄竅簫簧簪簞簣簡糧繡罈翹翻聶臍臏舊藍藐藉薦蟯蟬蟲蟠覆觴謬謫豐贅蹙蹣蹦蹤蹕軀轉轍邇醬釐鎔鎖鎢鎳鎮鎬鎘鎗闔闖闐闕離雜雙雛獺癡礙穩籀繫羅羶臘藷蟻蠅襟襞譏譆贊蹴邋鏜鏢鏨關霪霧韻覺譯飄馨麵鼯齟齣齡囁櫻瓖瓔籐纏蘗蘚蠣蠢襪襬譴禳籠聽臟襲襯觼鑄霾韁顫饕驍鬚攫籣弊乾靨驗灞氶犮玊禸优伀伝伓伄匢匟厊圪夼奼尥庄弚扜扤扢朹机朿朳汆汒汋吨吤呇囧坁坅坌奀妠妗妎妡岊庋庉弝彸忭忤忺怀抎抏扽扲攷旳杅邟邧阰阭佼佽侀佶佪侞侕佫佮冾刵剆劼咂咈呥呦呡呤囷囹坫坱坶怴怊怬怉怜戔戽抭抪抶抩抸昍旽昐枆杴枟枙极歾沭泂沺泆泭泲肣苀芛芞芨芶虰虭迕迗邴邯邳俍侲俔俜侻勀厙峇峊峓峈峆峎峟帡帢帠彖怹恲恓恇恛恀怤恄恘恦恮挎挃拫挏拸拶挔昶昲昢昫柈枺炷炾炡牁牉牬牮狤狨狫狪珌珅玹玶玵玴玿珆玸珋瓬瓮畇畈疧盄眃眄盺砆砒砐祊种窀苭衎衪衩觓赲迣迡郕郅郈陊倓倵凄凎剞剟剕勍唚哿唄唑恁悢悀悝悗戙扆拲捄捅挶揤捋捊挳捚捑捀捈敆旆晇晑栻栖栱栫桎桄栒栦栨桍栠欭欱欴毤牷牶猀狺狴狻玼珜珛珔瓟瓵疰疻痀眝眐眙砬砪砱祏祜祓祒祑秠秭秝窅窊茛茪茈茼荍茖茤茠茷茯荓荋茧荈虒蚖蚑蚇蚥蚡蚘蚎蚝蚔袀豗赶趷軓迵适逄郚郘郜酐啎啈唭唻埢埶埜埴埽堈堋埮埲埥埬埡埼堐堁堌埱埩堍堄奜婠婘娸婐婥婗婃婝婒婄娾娹婜孮寁寀屙崞崌崨崥崏捸掅晥晛晢梇桮梮梫楖梬桵梏桲梀梛梖梠梊殎殌淀涴淔渀淈淖淜淝淴淊淭痋痑痐眽眥硒祧祪祣秺窐笘笝笱笫笲粘粣紶絅紬紿絊紻羕羛翊翏翉耟蚺蚳蚸蛌蚻蛃蚽蚾衒袕袨袪袑袡袟袘觙觕訧赹趿軘軞軝軜逡郪郰祡笘繫峈硐倜袧啣爵魒豻蟈邿';
}
function FTPYStr()
{
	return '陣坨焰叢涂糞薡煇瞌?潣琀???饎??踇杰?遁?寯??惾?煝煁?訊揝??窫鎤艿娗耇?圢??洒??鴛蔩宴熼?墿?攘滱??烒想惝攎訓腷鎉洙??鄆膏鄱墏鉺褰渙茦??熸?墣?鞍煠啅揈?伻?稔嶒絻愃???悊???跚寂?綒???捼詁?浮??嫽??櫮?莮??????蕺殟鱨謞?甓?璾愓追簌懅柒琵榰???偰?鵁愐???嘯傔鄣餫?鎌穆烿韅籌?慏敥恚??瘴?滭屳猧?渧??銼?顥??憨綟?緻???殠??罘??燊罠禭??朏?懁???惸??綎?璯徼掔溾?廠摽荾???虐???猊揱揥暖焌惍霑??漞陂??搪?脭褧????杍??蕣餞廢螮撝?鞨?屠頒嶲鑪鴆霓??沎?鋐菀敜?潬?矂阣?琌嫠糯??綹彪寶?頰篇溓廌愚蕆?譑?蟊髺蔉留?幟黯霥麝?讙??磐惈駟嫹?筳專紒陜稷漚?攉嶱褷掝?帴?猩?硿???躍?芅罰澢帗???罜嚐?踜?蜥褩夎?跎??琥螈??窣??????墮戤??腶?青霹?譿揎?馞薆繟?鞭蓋螗演姁崇韕??掰???贆踂糕鄭???霄?怲縫??蜘??篌??徨?醪糠惁?敤腯犢烯翷耷??嬇箍扊欒潽嬡??嶠猞??鞏??懌頀??牾?歊慮褣嚶瀼甔??讋?????撣漃憂?煢?焂屜?攲?齺瀖髼螿?妱啐?薖?舴?櫐靪???槊箌腛?珼嫶箑?莛??沄??澍齋稟節笒愩???睢?隇?顁趲??錐璲庲徵緬?誫嶰?灺饖?礝?旖嬐???揗畿掾偞??鯧?霎癖?泯?姝?舑?????亭?渹胍傛?褞掞??挲幅緲???酏悿妶偍顏螜?酬窺?熵信??軏舳?磧瓃猝豩觳闚惿踚?翞緙螹螴??鷨鎛?睥烝?薉鋒璗綃?鋤??牼苾?慶??娩?睜獘惛?閼悷屧烔???姵?杯彉鰎儡敪?嶪鞞蕔政侵瞥礌膿???跐????誏?爍淽?珴踊?鶸蓽錕??偝髯鴦?腢起糊?鎧??綄鷸???????傃??焀????訖璵?綌矮畣???嫢?薐朡???膽?荽掜莩?頤佧?緉??龜???搭軔萉??赩?沮絿悾?礩鷂惤裳買沬荿蕼鎵鼆毹?罫?廣郿?????輇蕄綅姖奫??褭徭??湠镽??窠?揊?????幨厜密?霉?旒?猣柰溟潐稛??偅???沱?聃???斮潝皺意壤胐猗嵞??豷喭焆崮?愊鄯慷霣戟滖?娊嘸??鋰??揃?遨?築軿唵瞎?濃?螖?孬頯癓?';
}
function Traditionalized(cc){
	var str='',ss=JTPYStr(),tt=FTPYStr();
	for(var i=0;i<cc.length;i++)
	{
		if(cc.charCodeAt(i)>10000&&ss.indexOf(cc.charAt(i))!=-1)str+=tt.charAt(ss.indexOf(cc.charAt(i)));
  		else str+=cc.charAt(i);
	}
	return str;
}
function Simplized(cc){
	var str='',ss=JTPYStr(),tt=FTPYStr();
	for(var i=0;i<cc.length;i++)
	{
		if(cc.charCodeAt(i)>10000&&tt.indexOf(cc.charAt(i))!=-1)str+=ss.charAt(tt.indexOf(cc.charAt(i)));
  		else str+=cc.charAt(i);
	}
	return str;
}

function setCookie(name, value)		//cookies扢离
{
	var argv = setCookie.arguments;
	var argc = setCookie.arguments.length;
	var expires = (argc > 2) ? argv[2] : null;
	if(expires!=null)
	{
		var LargeExpDate = new Date ();
		LargeExpDate.setTime(LargeExpDate.getTime() + (expires*1000*3600*24));
	}
	document.cookie = name + "=" + escape (value)+((expires == null) ? "" : ("; expires=" +LargeExpDate.toGMTString()));
}

function getCookie(Name)			//cookies黍
{
	var search = Name + "="
	if(document.cookie.length > 0) 
	{
		offset = document.cookie.indexOf(search)
		if(offset != -1) 
		{
			offset += search.length
			end = document.cookie.indexOf(";", offset)
			if(end == -1) end = document.cookie.length
			return unescape(document.cookie.substring(offset, end))
		 }
	else return ""
	  }
}

var StranLink_Obj=document.getElementById("StranLink")
if (StranLink_Obj)
{
	var JF_cn="ft"+self.location.hostname.toString().replace(/\./g,"")
	var BodyIsFt=getCookie(JF_cn)
	if(BodyIsFt!="1")BodyIsFt=Default_isFT
	with(StranLink_Obj)
	{
		if(typeof(document.all)!="object") 	//準IE銡擬
		{
			href="javascript:StranBody()"
		}
		else
		{
			href="#";
			onclick= new Function("StranBody();return false")
		}
		title=StranText("萸僻眕楛极笢恅源宒銡擬",1,1)
		innerHTML=StranText(innerHTML,1,1)
	}
	if(BodyIsFt=="1"){setTimeout("StranBody()",StranIt_Delay)}
}