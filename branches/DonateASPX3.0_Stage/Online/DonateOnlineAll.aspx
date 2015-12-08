<%@ Page Language="C#" EnableEventValidation="false" CodeFile="DonateOnlineAll.aspx.cs" Inherits="Online_DonateOnlineAll" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>線上奉獻首頁</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <style type="text/css">
        body
        {
            background-color: #CCC;
            margin-top: 0px;
            text-align: center;
        }
        .alignleft
        {
            font-size: 12px;
            color: #000;
            text-align: left;
        }
        body, td, th
        {
            font-size: 12px;
            color: #000;
            text-align: left;
        }        

        #header, #container{
	        width:955px;
	        clear:both;
        }
        #header{
	        height:70px;
	        background:url("images/main_header_background.jpg") repeat-x;
            position:relative;
        }
        #logo{
            float:left;
	        margin:0px;
	        margin-left:21px;
        }
        #logo img{
            border:0;
        }
        #newsletter-btn{
	        float:left;
	        display:block;
	        width:166px;
	        height:28px;
	        background:url("images/newsletter-btn.png") no-repeat;
	        margin-left:570px;
	        cursor:pointer;
            position:absolute;
            right:0;
            top:0;
        }
        #tv_station{
	        list-style:none;
	        float:left;
	        margin:0;
	        padding:0;
	        margin-left:21px;
	        margin-top:0px;
        }
        #tv_station li{
	        float:left;
        }
        #tv_station li > .btn{
	        display:block;
	        /*width:78px;
	        height:49px;*/
	        cursor:pointer;
        }
        #tv_station li > .oneChannel_btn{
	        background: url("images/a.btn.oneChannel_btn_o.jpg") no-repeat;
            width:98px;
	        height:68px;
        }
        #tv_station li > .twoChannel_btn{
	        background: url("images/a.btn.twoChannel_btn_o.jpg") no-repeat;
            width:115px;
	        height:68px;
        }

        #tv_station li > .nothASite_btn{
	        background: url("images/a.btn.nothASite_btn_o.jpg") no-repeat;
            width:94px;
	        height:68px;
        }

        #tv_station li > .hkSite_btn{
	        background: url("images/btn_HKSite_o.jpg") no-repeat;
        }

        #tv_station li > .prodSite_btn{
	        background:url("images/a.btn.prodSite_btn_o.jpg") no-repeat;
            width:91px;
	        height:68px;
        }
        
         #tv_station li > .fb_btn{
            width:30px;    
            height:68px;  
	        background: url("images/icon_face_16x16.jpg") 65% 65% no-repeat;
        }
        
        #tv_station li > .g_btn{
            width:20px;
            height:68px;  
	        background: url("images/icon_Google.jpg") 65% 65% no-repeat;
        }

        #tv_station a:LINK, #tv_station a:VISITED{
	        text-decoration:none;
	        color:#848689;
	        font-size:14px;
	        letter-spacing:3px;
        }
        #tv_station a:HOVER, #tv_station a:ACTIVE{
	        text-decoration:none;
	        color:#ea3c20;
        }

        /* search form */
        #search_form{
	        float:right;
	        margin-top:8px;
	        margin-right:20px;
        }

        #search_form input{
	        padding:3px;
        }
        #navigator{
	        width:100%;
	        height:32px;
	        background:url("images/main_navigator_background.png") repeat-x;
        }

        #navigator_line{
	        margin:0 0 0 42px;
	        line-height:32px;
	        vertical-align: middle;
        }

        #navigator_line a{
	        display: block;
	        float: left;
	        font-size: 12px;
	        text-decoration:none;
        }

        #navigator_line .nav_item{
	        color: #333;
        }

        #navigator_line .nav_add{
	        width: 12px;
	        height: 12px;
	        background: url("images/main_navigator_add.png") no-repeat;
	        margin: 9px 5px;
        }
        #top_menu{
	        float:right;
	        list-style:none;
	        margin:0px;
	        margin-right:20px;
        }

        #top_menu li{
	        float:left;
	        margin:0 6px;
	        list-style:none;
        }

        #top_menu a{
	        line-height:32px;
	        vertical-align:middle;
	        font-size:12px;
	        color:#44507b;
        }

        #top_menu a:LINK,
        #top_menu a:VISITED{
	        text-decoration:none;
        }

        #top_menu a:HOVER,
        #top_menu a:ACTIVE{
	        text-decoration:underline;
        }

        #top_menu a.active{
	        color:#e40448;
        }
        
        /* donate_div */
        #donate_banner{
            background-color:#FFFFFF;
        }
        #donate_title{
            margin-top:20px;	        
	        margin-left:25px;
        }
        #address
        {
            width:100%;
            text-align:center;
        }
        #contact
        {
            width:100%;
            text-align:center;
        }
        #copyright
        {
            width:100%;
            text-align:center;
        }
        #footer
        {
            clear:both;
            height:62px;
            padding-top:30px;
            background:url(images/main_footer_background.jpg) #fff repeat-x 0 30px;
            font-size:0.8em;
            color:#8a9ca6;
            font:12px/1.231 arial,helvetica,calen,sans-serif;
        }

        #tContent td {
            font-size: 18px;
        }

    </style>
    <script type="text/javascript">

        function MM_preloadImages() { //v3.0
            var d = document; if (d.images) {
                if (!d.MM_p) d.MM_p = new Array();
                var i, j = d.MM_p.length, a = MM_preloadImages.arguments; for (i = 0; i < a.length; i++)
                    if (a[i].indexOf("#") != 0) { d.MM_p[j] = new Image; d.MM_p[j++].src = a[i]; }
            }
        }
        function MM_swapImgRestore() { //v3.0
            var i, x, a = document.MM_sr; for (i = 0; a && i < a.length && (x = a[i]) && x.oSrc; i++) x.src = x.oSrc;
        }
        function MM_findObj(n, d) { //v4.01
            var p, i, x; if (!d) d = document; if ((p = n.indexOf("?")) > 0 && parent.frames.length) {
                d = parent.frames[n.substring(p + 1)].document; n = n.substring(0, p);
            }
            if (!(x = d[n]) && d.all) x = d.all[n]; for (i = 0; !x && i < d.forms.length; i++) x = d.forms[i][n];
            for (i = 0; !x && d.layers && i < d.layers.length; i++) x = MM_findObj(n, d.layers[i].document);
            if (!x && d.getElementById) x = d.getElementById(n); return x;
        }

        function MM_swapImage() { //v3.0
            var i, j = 0, x, a = MM_swapImage.arguments; document.MM_sr = new Array; for (i = 0; i < (a.length - 2); i += 3)
                if ((x = MM_findObj(a[i])) != null) { document.MM_sr[j++] = x; if (!x.oSrc) x.oSrc = x.src; x.src = a[i + 2]; }
        }

    </script>
    <script type="text/javascript">

        $(document).ready(function () {
            
            var host = $(location).attr('host'); //site url domain and host 沒有 http://

            var url_homepage = host.replace(/donate3/g, "http://www"); //首頁
            var url_second = host.replace(/donate3/g, "http://www") + '/?second';  // 二台
            var url_video = host.replace(/donate3/g, "http://www") + '/video';  //節目總覽
            var url_schedule = host.replace(/donate3/g, "http://www") + '/video/index.php/schedule';  //節目時間表
            var url_daily = host.replace(/donate3/g, "http://www") + '/daily';  // 每日靈糧
            var url_news = host.replace(/donate3/g, "http://www") + '/good-news';  //好消息月刊
            var url_about = host.replace(/donate3/g, "http://www") + '/about';  //關於好消息
            var url_service = host.replace(/donate3/g, "http://www") + '/service';  //線上客服
            //替換掉link的屬性
            $('#url_logo').attr("href", url_homepage);
            $('#setStationOneBtn').attr("href", url_homepage);
            $('#url_navigator').attr("href", url_homepage);
            $('#url_2').attr("href", url_second);
            $('#url_video').attr("href", url_video);
            $('#url_schedule').attr("href", url_schedule);
            $('#url_daily').attr("href", url_daily);
            $('#url_news').attr("href", url_news);
            $('#url_about').attr("href", url_about);
            $('#url_service').attr("href", url_service);

            //所有TextBox的開頭是txt都限制輸入數字，除了txtDataDate和txtDeptName
            $('input:text[id*=txt]').not('').bind("keyup", function (e) {
                var val = $(this).val();
                //if (isNaN(val)) {
                    val = val.replace(/[^0-9\.]/g, '');
                    $(this).val(val);
                //}
            });

            //選擇捐款方式
            $('#Item').bind("change", function (e) {
                $('#HFD_chkItem').val($(e.target).val());
                if ($('#Item').val() == "") {
                    $('#lblItemPoint').text('必選');
                    $('#lblItemPoint').addClass('blink');
                }
                else {
                    $('#lblItemPoint').removeClass('blink');
                    $('#lblItemPoint').text('');
                }
            });


            //捐款金額
            $('#txtAmountOnce').bind("blur", function (e) {
                if ($('#txtAmountOnce').val() != "") {
                    if ($('#txtAmountOnce').val() < 100) {
                        $('#sAmount').addClass('blink');
                        $('#sAmount').css('color', 'red');
                        $('#lblAmountPoint').text('');
                        $('#lblAmountPoint').removeClass('blink');
                    }
                    else {
                        $('#lblAmountPoint').text('');
                        $('#lblAmountPoint').removeClass('blink');
                        $('#sAmount').removeClass('blink');
                        $('#sAmount').css('color', 'blue');
                    }
                }
                else {
                    $('#lblAmountPoint').text('必填');
                    $('#lblAmountPoint').addClass('blink');
                    $('#sAmount').removeClass('blink');
                    $('#sAmount').css('color', 'blue');
                }
            });



        });

        function Foolproof() {

            var error_cut = 0;

            //捐款方式
            if ($('#HFD_chkItem').val() == "") {
                error_cut++;
                $('#lblItemPoint').text('必選');
                $('#lblItemPoint').addClass('blink');
                $('#HFD_chkItem').focus();
            }
            else {
                $('#lblItemPoint').removeClass('blink');
                $('#lblItemPoint').text('');
            }
            
            //捐款金額
            if ($('#txtAmountOnce').val() != "") {
                if ($('#txtAmountOnce').val() < 100) {
                    error_cut++;
                    $('#sAmount').addClass('blink');
                    $('#sAmount').css('color', 'red');
                    $('#lblAmountPoint').text('');
                    $('#lblAmountPoint').removeClass('blink');
                    $('#txtAmountOnce').focus();
                }
                else {
                    $('#lblAmountPoint').text('');
                    $('#lblAmountPoint').removeClass('blink');
                    $('#sAmount').removeClass('blink');
                    $('#sAmount').css('color', 'blue');
                }
            }
            else {
                error_cut++;
                $('#lblAmountPoint').text('必填');
                $('#lblAmountPoint').addClass('blink');
                $('#sAmount').removeClass('blink');
                $('#sAmount').css('color', 'blue');
            }

            if (error_cut > 0) {
                return false;
            }
            else {
                $('#HFD_Amount').val($('#txtAmountOnce').val())
                document.forms[0].action = "ShoppingCart.aspx";
            }

        }

    </script>
</head>
<body onload="MM_preloadImages('images/btn_其他奉獻_down.jpg','images/btn_金融機構奉獻_down.jpg','images/btn_授權書奉獻_down.jpg','images/btn_線上徵信_down.jpg')">
    <form id="form1" runat="server">
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_Amount" runat="server" />
    <table width="949" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td colspan="5">
                <div id="header">
                    <h1 id="logo">
    	                <a id="url_logo">
                            <img height="70" width="162" id="good_tv_logo" title="好消息電視台" alt="好消息電視台" src="../Online/images/logo.jpg" />
                        </a>
                    </h1>
                    <a id="newsletter-btn" target="_blank" href="http://w2.goodtv.org/newsletter/"></a>
                    <ul id="tv_station">
    	                <li><a id="setStationOneBtn" class="btn oneChannel_btn"></a></li>
                        <li><a id="url_2" class="btn twoChannel_btn"></a></li>
                        <li><a target="_blank" href="http://www.goodtvusa.tv/Site/index.php/home" class="btn nothASite_btn"></a></li>
                        <li><a target="_blank" href="http://shop.goodtv.tv/" class="btn prodSite_btn"></a></li>
                        <li><a target="_blank" href="https://www.facebook.com/goodtv" class="btn fb_btn"></a></li>
                        <li><a target="_blank" href="https://plus.google.com/112408150725302065226" class="btn g_btn" rel="publisher"></a></li>
                    </ul>
                </div>

                <div id="container">
	                <div id="navigator">
    	                <div id="navigator_line">
                            <a id="url_navigator" class="nav_item btn">首頁</a>
                            <a class="nav_add"></a><a class="nav_item">支持與奉獻</a>
                        </div>
                        <ul id="top_menu">
        	                <li> <a id="url_video" class="btn">節目總覽</a> </li>
                            <li> <a id="url_schedule" class="btn">節目時間表</a> </li>
                            <li> <a id="url_daily" class="btn">每日靈糧</a> </li>
                            <li> <a id="url_news" class="btn">好消息月刊</a> </li>
                            <li> <a id="url_about" class="btn">關於好消息</a> </li>
                            <li> <a href="donateonlineall.aspx" class="btn active">支持與奉獻</a> </li>
                            <li> <a id="url_service" class="btn">線上客服</a> </li>
                        </ul>
                    </div>
                </div> 

                <div id="donate_banner">
                    <img id="donate_title" class="title_image" src="images/donate_title.jpg" alt="支持與奉獻" title="支持與奉獻"/>
                </div>  
            </td>
        </tr>
        <tr>
            <td width="3" valign="top" bgcolor="#FFFFFF">
                &nbsp;
            </td>
            <td width="47" bgcolor="#FFFFFF">
                &nbsp;
            </td>
            <td width="204" valign="top" bgcolor="#FFFFFF">
                <table width="200" border="0" align="left" cellpadding="1" cellspacing="1">
                    <tr>
                        <td align="left">
                            <img src="images/btn_線上奉獻_down.jpg" name="donate_online" alt="線上奉獻頁面" width="198" height="36" border="0" id="donate_online" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <a href="DonateQry.aspx" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('DonateQry','','images/btn_線上徵信_down.jpg',1)">
                                <img src="images/btn_線上徵信_O.jpg" alt="線上徵信" name="DonateQry" width="198" height="36"
                                    border="0" id="DonateQry" /></a>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <a href="Authorization_Letter.html" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Authorization Letter','','images/btn_授權書奉獻_down.jpg',1)">
                                <img src="images/btn_授權書奉獻_O.jpg" alt="下載授權書" name="Authorization Letter" width="198"
                                    height="91" border="0" id="Authorization Letter" /></a>
                        </td>
                    </tr>                   
                    <tr>
                        <td>
                            <a href="financial.html" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('financial institution','','images/btn_金融機構奉獻_down.jpg',1)">
                                <img src="images/btn_金融機構奉獻_O.jpg" alt="金融機構" name="financial institution" width="198"
                                    height="91" border="0" id="financial institution" /></a>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <a href="donate_other.html" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image9','','images/btn_其他奉獻_down.jpg',1)">
                                <img src="images/btn_其他奉獻_O.jpg" alt="其他奉獻" name="Image9" width="198" height="91" border="0" id="Image9" />
                            </a>
                        </td>
                    </tr>
                    <tr>
                        <td height="25" valign="bottom">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td valign="bottom">
                            <img alt="" src="images/奉獻頁_leftDown.jpg" width="195" height="214" />
                        </td>
                    </tr>
                </table>
            </td>
            <td width="16" align="left" bgcolor="#FFFFFF">
                &nbsp;
            </td>
            <td width="686" align="left" valign="top" bgcolor="#FFFFFF">
                <table id="tContent" width="650" border="0" cellspacing="2" cellpadding="2">
                    <tr>
                        <td colspan="2" align="right" valign="bottom">
                            <!--<div align="right"><strong><font color="#003399">Language：</font><a href="../English/DonateOnlineAll.aspx" target="_self"><font size="3">English</font></a></strong></div>-->
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <img alt="" src="images/Titlebar_線上奉獻2.jpg" width="650" height="81" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <!--span>感謝神！透過衛星直播和網路科技的便利，我們讓佈道會在電視上播出，也透過網路直播的方式，讓教會成為另一個佈道會據點，擴大參與人數。例如「林書豪佈道會」，雖然在現場只有六千人參加，但透過全省直播的七百間教會統計，決志總人數超過一萬人，這就是科技帶來的突破和改變。
                            GOOD TV成立的異象，是以傳福音為使命。<br />
                            2012年，神賞賜給我們嶄新的宣教大樓，擁有一切先進的設備，我們清楚明白神為靈魂得救迫切的心，GOOD TV一定要回應神的呼召和感動，傳福音的責任不可推卸。
                            </!--span><br /><br /-->
                            <br />
                            <table width="100%" border="0" cellspacing="2" cellpadding="4">
                                <tr>
                                    <td><div style="text-align:center; font-size: 20px; color:crimson; font-family:Microsoft JhengHei; font-weight: bold;">你需要好消息，好消息也需要你</div>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <div style="text-align:left; color: brown; font-family:Microsoft JhengHei; font-weight: bold; font-size: 16px; line-height:25px;">
                                            Q.GOOD TV的營運及需要？
                                            </div>
                                        <div style="text-align:left; color: black; font-family:Microsoft JhengHei; font-weight: bold; font-size: 16px; line-height:25px;">
                                            GOOD TV是基督教超宗派的電視台，兩個頻道都是靠點滴奉獻來維持營運，是一個非營利的基督教家庭頻道。<br />
                                                我們一年營運費用至少需要新台幣四億元，以支應以下用途：<br />
                                                •	製作預工、福音、佈道、造就等類的優質節目、購買海外優質影片<br />
                                                •	高畫質製播設備、衛星租金、網站及傳輸費用，發展網路宣教事工<br />
                                                •	設立家庭關懷熱線中心，培訓專業專線志工及夥伴教會<br />
                                                •	製作教育類、家庭類教材及查經材料<br />
                                                •	人事管銷費用<br />
                                                </div>
                                        <div style="text-align:left; color: brown; font-family:Microsoft JhengHei; font-weight: bold; font-size: 16px; line-height:25px;">
                                            Q.何謂「GOOD TV天使」？如何奉獻支持GOOD TV？
                                            </div>
                                        <div style="text-align:left; color: blue; font-family:Microsoft JhengHei; font-weight: bold; font-size: 16px; line-height:25px;">
                                            媒體宣教需要的支持是長期持續的認獻。多年來若您能看到我們的節目，是因為背後有許多GOOD TV天使的奉獻祝福，他們獻上的五餅二魚，被神擘開就餵飽許多靈魂。邀請您加入祝福人的GOOD TV天使行列。就是「同心同行」地回應GOOD TV全球福音大平台的異象，您的五餅二魚，與GOOD TV一同宣教到地極。<br />
                                                </div>
                                        <div style="text-align:left; color: black; font-family:Microsoft JhengHei; font-weight: bold; font-size: 16px; line-height:25px;">
                                            歡迎您以認獻方式，於每月、每季、每半年或一年<u>採定期定額奉獻</u>，加入GOOD TV奉獻天使行動，與GOOD TV一同收割莊稼。 年度認獻金額，可訂定每年1萬、10萬、50萬、100萬、500萬、1000萬或每月至少1000元。<br />
                                                </div>
                                    </td>
                                </tr>
                            </table>
                            <!--font size="4" color="#FF5511"><b>您的奉獻將使用在以下兩大用途：</b></!--font><br />
                            <table width="100%" border="0" cellspacing="2" cellpadding="4">
                                <tr>
                                    <td style="vertical-align: top"><span>一、</span></td>
                                    <td><span>「優質媒體發展預算」，期望在新的季節呈現新的節目，擴展各類媒體通路，發展優質全方位電視台的福音影響力。</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="vertical-align: top"><span>二、</span></td>
                                    <td><span>「經常費用」除原有管銷、人事費用之外，另增加營運所需器材、資訊機房、主副控設備、工程維護保固等費用支出。</span>
                                    </td>
                                </tr>
                            </table-->
                       </td>
                    </tr>
                   <!--tr>
                        <td colspan="2">
                            <span><b>我們傳揚他，是用諸般的智慧，勸戒各人，教導各人，要把各人在基督裡完完全全的引到神面前。(西一28)</b></span>
                        </td>
                    </!--tr-->
                    <tr><td colspan="2"> </td></tr>
                    <tr>
                        <td style="text-align: right;">
                            <font size="4" color="#FF5511"><b>請選擇捐款方式：</b></font>
                        </td>
                        <td>
                            <asp:DropDownList ID="Item" name="Item" runat="server" Font-Size="18px" Width="40mm"></asp:DropDownList>
                            <asp:Label ID="lblItemPoint" runat="server"></asp:Label>
                            </td>
                    </tr>
                    <tr>
                        <td style="text-align: right;">
                            <font size="4" face='新細明體'>新台幣：</font>
                        </td>
                        <td>
                            <asp:TextBox ID="txtAmountOnce" runat="server" Width="40mm" Style="text-align: right" Font-Size="18px"></asp:TextBox>
                            <asp:Label ID="lblAmountPoint" runat="server"></asp:Label>
                            <font size="4" face='新細明體'>元整</font>　<span id="sAmount" style="color: blue; font-weight: bold; font-size: 14px;">最低金額為100元</span>
                        </td>
                    </tr>
                    <tr>
                        <td height="10" colspan="2" align="center">
                            <div class="function">
                                <asp:ImageButton ID="btnNext" class="" runat="server" ImageUrl="images/next.jpg" OnClientClick="return Foolproof();" />
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                                <img alt="" src="images/pic_faith.jpg" width="650" height="100" border="0" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td height="100" colspan="5" align="right" valign="bottom" bgcolor="#FFFFFF">
                <div id="footer">
                    <div id="address">
			            <a>財團法人加百列福音傳播基金會</a>
			            <a>　　TEL:02-8024-3911</a>
			            <a>　　FAX:02-8024-3933</a> 
		            </div>
		            <div id="contact">
			            <a>Gabriel Broadcasting Foundation. All Rights Reserved.</a>
		            </div>
		            <div id="copyright">
			            <a>版權所有，請勿擅自轉載節錄重製。</a>
		            </div>
                </div>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
