<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateOnlineAll.aspx.cs"
    Inherits="Online_DonateOnlineAll" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <%--<meta http-equiv="X-UA-Compatible" content="IE=edge" />--%>
    <title>Donation Home</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/jquery.address.js"></script>
    <script type="text/javascript" src="../include/jquery.field.js"></script>
    <script type="text/javascript" src="../include/caseutil.js"></script>
    <%--<script type="text/javascript" src="../include/jquery-1.7.2.min.js"></script>--%>
    <script type="text/javascript" src="../include/jquery-ui-1.8.18.custom/development-bundle/ui/jquery.ui.core.js"></script>
    <script type="text/javascript" src="../include/jquery-ui-1.8.18.custom/development-bundle/ui/jquery.ui.datepicker.js"></script>
    <script type="text/javascript" src="../include/jquery-ui-1.8.18.custom/development-bundle/ui/i18n/jquery.ui.datepicker-zh-TW.js"></script>
    <link type="text/css" rel="stylesheet" media="all" href="../include/jquery-ui-1.8.18.custom/development-bundle/themes/ui-lightness/jquery.ui.all.css" />
    <script type="text/javascript">
        $(document).ready(function () {

            //所有TextBox的開頭是txt都限制輸入數字，除了txtDataDate和txtDeptName
            $('input:text[id*=txt]').not('').bind("keyup", function (e) {
                var val = $(this).val();
                if (isNaN(val)) {
                    val = val.replace(/[^0-9\.]/g, '');
                    $(this).val(val);
                }
            });

            //            $('input:checkbox[id*=CHK_Item]').bind("click", function (e) {
            //                $('#HFD_chkItem').val($('input:checkbox:checked[name="Item"]').getValue());
            //                var itemList = GetItemList();
            //                $('#lblGridList').html(itemList);
            //                var s = $(e.target).attr('ID');
            //                if (!$(e.target).attr('checked')) {
            //                    $('#txtAmount' + s.substr(s.length - 1, 1)).val('');
            //                }
            //            });
            $('input:radio[id*=RDO_Item]').bind("click", function (e) {
                $('#HFD_chkItem').val($(e.target).val());
                var itemList = GetItemList();
                $('#lblGridList').html(itemList);
                if ($(e.target).val() == '單筆奉獻') {
                    $('#txtAmountPeriod').val('');
                    $('#txtAmountPeriod').attr('disabled', true);
                    $('#txtAmountOnce').attr('disabled', false);
                }
                else {
                    $('#txtAmountOnce').val('');
                    $('#txtAmountOnce').attr('disabled', true);
                    $('#txtAmountPeriod').attr('disabled', false);
                }
            });

            $('input:radio[id*=rdoPayType]').bind("click", function (e) {
                PayTypeChange();
            });

            $('input:text[id*=txt]').bind("focusout", function (e) {
//                $('#HFD_chkItem').val($('input:checkbox:checked[name="Item"]').getValue());
                //$('#HFD_chkItem').val($('input:radio[name="RDO_Item"]').getValue());
                var itemList = GetItemList();
                $('#lblGridList').html(itemList);
            });

            $('[id*=ddl]').bind("change", function (e) {
                var itemList = GetItemList();
                $('#lblGridList').html(itemList);
            });

            if ($('#HFD_chkItem').val() == "") {
                $('#btnNext').hide();
            }
            else {
                var itemList = GetItemList();
                $('#lblGridList').html(itemList);
                $('#btnNext').show();
            }

            if ($('#HFD_NeedCheck').val() == 'Y') {
                $('#btnCancel').hide();
            }
            else {
                $('#btnCancel').show();
            }

            //$('#divMonth').hide();
            //$('#divQtr').hide();
            // 季繳提示訊息 by Hilty 2013/12/17
            //$('#divQtrText').hide();
            //$('#divYear').hide();
            $('#txtAmountOnce').attr('disabled', true);
            $('#txtAmountPeriod').attr('disabled', true);
        });                          //end of ready()

        window.history.forward(1);
    </script>
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
    </style>
    <style type="text/css">
        #header, #container{
	        width:955px;
	        clear:both;
        }
        #header{
	        height:70px;
	        background:url("images/main_header_background.jpg") repeat-x;
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
        }
        #tv_station{
	        list-style:none;
	        float:left;
	        margin:0;
	        padding:0;
	        margin-left:21px;
	        margin-top:-7px;
        }
        #tv_station li{
	        float:left;
        }
        #tv_station li > .btn{
	        display:block;
	        width:78px;
	        height:49px;
	        cursor:pointer;
        }
        #tv_station li > .oneChannel_btn{
	        background: url("images/btn_oneChannel_o.jpg") no-repeat;
        }
        #tv_station li > .twoChannel_btn{
	        background: url("images/btn_twoChannel_o.jpg") no-repeat;
        }

        #tv_station li > .nothASite_btn{
	        background: url("images/btn_NothASite_o.jpg") no-repeat;
        }

        #tv_station li > .hkSite_btn{
	        background: url("images/btn_HKSite_o.jpg") no-repeat;
        }

        #tv_station li > .prodSite_btn{
	        width:93px;
	        background:url("images/btn_prodSite_o.jpg") no-repeat;
        }
        
         #tv_station li > .fb_btn{
            width:30px;           
	        background: url("images/icon_face_16x16.jpg") no-repeat;
        }
        
        #tv_station li > .g_btn{
            width:20px;
	        background: url("images/icon_Google.jpg") no-repeat;
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
        function Foolproof() {
            /*if (!$('#rdoPurpose1').prop('checked') && !$('#rdoPurpose2').prop('checked')) {
                alert('請選擇奉獻用途！');
                return false;
            }*/
            if ($('#RDO_Item1').attr('checked') && $('#txtAmountOnce').val() < 100) {
                alert('【' + $('#RDO_Item1').val() + '】Minimum: NT$ 100！');
                $('#txtAmountOnce').focus();
                return false;
            }
            if ($('#RDO_Item2').attr('checked') && $('#txtAmountPeriod').val() < 100) {
                alert('【' + $('#RDO_Item2').val() + '】Minimum: NT$ 100！');
                $('#txtAmountPeriod').focus();
                return false;
            }
            /*if ($('#RDO_Item2').attr('checked') && !$('#rdoPayType1').prop('checked') && !$('#rdoPayType2').prop('checked') && !$('#rdoPayType3').prop('checked')) 
            {
                alert('選擇奉獻週期！');
                return false;
            }*/

        }       
    </script>
    <script type="text/javascript">
        $(document).ready(function () {

            /*$('.oneChannel_btn').hover(
		        function () {
		            $(this).css('background-image', 'url(images/btn_oneChannel_down.jpg)');
		        },
		        function () {
		            $(this).css('background-image', 'url(images/btn_oneChannel_o.jpg)');
		        }
	        );

            $('.twoChannel_btn').hover(
		        function () {
		            $(this).css('background-image', 'url(images/btn_twoChannel_down.jpg)');
		        },
		        function () {
		            $(this).css('background-image', 'url(images/btn_twoChannel_o.jpg)');
		        }
	        );

            $('.nothASite_btn').hover(
		        function () {
		            $(this).css('background-image', 'url(images/btn_NothASite_down.jpg)');
		        },
		        function () {
		            $(this).css('background-image', 'url(images/btn_NothASite_o.jpg)');
		        }
	        );

            $('.hkSite_btn').hover(
		        function () {
		            $(this).css('background-image', 'url(images/btn_HKSite_down.jpg)');
		        },
		        function () {
		            $(this).css('background-image', 'url(images/btn_HKSite_o.jpg)');
		        }
	        );

            $('.prodSite_btn').hover(
		        function () {
		            $(this).css('background-image', 'url(images/btn_prodSite_down.jpg)');
		        },
		        function () {
		            $(this).css('background-image', 'url(images/btn_prodSite_o.jpg)');
		        }
	        );*/

            $("#date").datepicker({
                changeYear: true,
                changeMonth: true,
                showOn: "button",
                buttonImage: "images/calendar.gif",
                buttonImageOnly: true,
                yearRange: "2000:2015",
                dateFormat: 'yy-mm-dd'
            });
            $("#date").focus(function () {
                $(this).val('');
                $("#date").datepicker("show");
            });

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
        });
    </script>
</head>
<body onload="MM_preloadImages('images/btn_其他奉獻_down.jpg','images/btn_金融機構奉獻_down.jpg','images/btn_授權書奉獻_down.jpg','images/btn_線上徵信_down.jpg')">
    <form id="form1" runat="server">
    <asp:HiddenField ID="HFD_NeedCheck" runat="server" />
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_ItemList" runat="server" />
    <asp:HiddenField ID="HFD_PayType" runat="server" />
    <asp:Panel ID="Panel1" runat="server" DefaultButton="btnNext">
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
                        <li><a target="_blank" href="http://www.goodtvusa.tv/site/" class="btn nothASite_btn"></a></li>
                        <li><a target="_blank" href="http://www.goodtvhk.tv/" class="btn hkSite_btn"></a></li>
                        <li><a target="_blank" href="http://shop.goodtv.tv/" class="btn prodSite_btn"></a></li>
                        <li><a target="_blank" href="https://www.facebook.com/goodtv" class="btn fb_btn"></a></li>
                        <li><a target="_blank" href="https://plus.google.com/112408150725302065226" class="btn g_btn" rel="publisher"></a></li>
                    </ul>
                    <div id="search_form">
    	                <%--<form action="http://www.goodtv.tv/video/index.php/video/search" method="post" accept-charset="utf-8" id="searchVideoForm" name="searchVideoForm">
                            <input type="text" name="date" value="選擇日期" size="6" id="date"  />
                            <input type="text" name="keyWord" value="" size="9" id="keyWord"  />
                            <asp:Button ID="submit" PostBackUrl="http://www.goodtv.tv/video/index.php/video/search" Text="節目搜尋" runat="server" />                               
                        </form>--%>
                    </div>
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
                            <li> <a href="donate_index.html" class="btn active">支持與奉獻</a> </li>
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
                            <img alt="" src="images/pic_index_3.jpg" width="195" height="121" />
                        </td>
                    </tr>
                </table>
            </td>
            <td width="16" align="left" bgcolor="#FFFFFF">
                &nbsp;
            </td>
            <td width="686" align="left" valign="top" bgcolor="#FFFFFF">
                <table width="650" border="0" cellspacing="2" cellpadding="2">
                    <tr>
                        <td colspan="2" align="right" valign="bottom">
                            <div align="right"><strong><font color="#003399">Language：</font><a href="../Online/DonateOnlineAll.aspx" target="_self"><font size="3">中文</font></a></strong></div>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <img alt="" src="images/Titlebar_線上奉獻.jpg" width="650" height="25" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <span style="font-size:16px;">Praise God, through satellite live broadcast and the convenience of Internet technologies, we are now able to broadcast evangelistic rally on TV; and through online live broadcast, enable churches to be evangelism venue to maximize attendance. Take [Jeremy Lin Taiwan Evangelistic Rally] for example, the stadium in Taipei had six thousand attendees, but with live broadcast partnership with over 700 churches in Taiwan, over 10,000 filled out forms to devote their lives to Christ. This is the breakthrough and change brought forth by technologies.<br /><br />
                            Good TV was established with an evangelistic mission. In the year 2012, God gracefully gave us brand new mission building with advanced equipment. We are fully aware of God’s pressing desire for souls to be saved, Good TV must respond promptly to God’s calling and preach the gospel.
                            </span><p />
                            <font size="3" color="#FF5511"><b>Your donation will be used toward these two major area:</b></font>
                            <%-- <asp:RadioButton ID="rdoPurpose1" GroupName="Purpose" runat="server" Text="經常費" ForeColor="Crimson" Font-Size="Small" />--%>
                            <table width="100%" border="0" cellspacing="2" cellpadding="2">
                                <tr>
                                    <td><span style="font-size:14px;">1.</span></td>
                                    <td><br /><span style="font-size:14px;"> [Operating Expense] in addition to Sales and administration, and human resource expense, also need to add operating equipment, Internet data center, Master Control and Sub Control equipment, and engineering maintenance expense.</span>
                                    </td>
                                </tr>
                                <%-- <asp:RadioButton ID="rdoPurpose2" GroupName="Purpose" runat="server" Text="優質媒體" ForeColor="Crimson" Font-Size="Small" />--%>
                                <tr>
                                    <td><span style="font-size:14px;">2.</span></td>
                                    <td><br /><span style="font-size:14px;"> [Media Development Budget] for new program development, expand diverse media channels and develop comprehensive Television Media Evangelistic influence power.</span>
                                    <%--<font style="background-color:#BDBDBD">■經常費</font><br />
                                    GOOD TV進駐宣教大樓後，經常費用除原有之管銷、人事費用外，並增加器材、機房、主控、工程維護保固等費用支出。
                                    <p />
                                    <font style="background-color:#BDBDBD">■優質媒體</font><br />
                                    GOOD TV為邁向全新的發展階段而推動「優質媒體發展預算」，期望在新的季節呈現新的節目，擴展各類媒體通路，發展優質全方位電視台的福音影響力。--%>
                                    </td>
                                </tr>
                            </table>
                       </td>
                    </tr>
                   <tr>
                        <td colspan="2">
                            <span style="font-size:13px;"><b>He is the one we proclaim, admonishing and teaching everyone with all wisdom, so that we may present everyone fully mature in Christ. (Colossians 1:28)</b></span>
                        </td>
                    </tr>
                    <%--<tr>
                        <td bgcolor="#FFCC00">
                            <b>&nbsp;&nbsp;奉&nbsp;獻&nbsp;用&nbsp;途&nbsp;：</b>
                        </td>
                        <td>
                            <asp:RadioButtonList ID="rblPurpose" runat="server" 
                                RepeatDirection="Horizontal" RepeatLayout="Flow">
                                <asp:ListItem>經常費</asp:ListItem>
                                <asp:ListItem>優質媒體</asp:ListItem>
                            </asp:RadioButtonList>
                        </td>
                    </tr>--%>
                    <tr>
                        <td colspan="2" align="center" valign="middle">
                            <font size="3" color="#FF5511"><b>You can choose to make one time donation or recurring donation.</b></font>
                        </td>
                    </tr>
                    <%--                    <tr>
                        <td colspan="2" style="color: #8B0000;">
                            <asp:Label ID="lblTitle" runat="server">平安שלום שלום！歡迎您使用線上奉獻，請填寫奉獻內容：</asp:Label><br />
                        </td>
                    </tr>--%>
                    <tr>
                        <td width="170" align="center" valign="middle">
                            <asp:Label runat="server" ID="lblItemOnce"></asp:Label>：
                        </td>
                        <td width="581">
                            <font size="2" face='新細明體'>NTD
                            <asp:TextBox ID="txtAmountOnce" runat="server" Width="40mm" Style="text-align: right"></asp:TextBox>
                            </font>　<span style="color: blue; font-weight: bold;">Minimum amount is $100 NTD</span>
                        </td>
                    </tr>
<%--                    <tr>
                        <td height="10" colspan="2" align="center">
                            <hr width="100%" size="1" />
                        </td>
                    </tr>--%>
                    <tr>
                        <td align="center" valign="top">
                            <asp:Label runat="server" ID="lblItemPeriod"></asp:Label>：
                        </td>
                        <td style="line-height:25px">
                            <font size="2" face='新細明體'>NTD
                            <asp:TextBox ID="txtAmountPeriod" runat="server" Width="40mm" Style="text-align: right"></asp:TextBox>
                            </font> <!--img alt="" src="images/New.png" height="20" /-->
                            
<%--                            <br />奉獻週期：
                            <asp:Label runat="server" ID="lblPayType"></asp:Label>--%>
                            <br />
                           <%-- <div id="divMonth">
                                開始年月：
                                <asp:DropDownList ID="ddlYearMS" runat="server">
                                </asp:DropDownList>
                                年
                                <asp:DropDownList ID="ddlMonthS" runat="server">
                                </asp:DropDownList>
                                月。
                                <br/>
                                結束年月：
                                <asp:DropDownList ID="ddlYearME" runat="server">
                                </asp:DropDownList>
                                年
                                <asp:DropDownList ID="ddlMonthE" runat="server">
                                </asp:DropDownList>
                                月。
                            </div>
                            <div id="divQtr">
                                開始年季：
                                <asp:DropDownList ID="ddlYearQS" runat="server">
                                </asp:DropDownList>
                                年
                                <asp:DropDownList ID="ddlQtrS" runat="server">
                                </asp:DropDownList>
                                季。
                                <br/>
                                結束年季：
                                <asp:DropDownList ID="ddlYearQE" runat="server">
                                </asp:DropDownList>
                                年
                                <asp:DropDownList ID="ddlQtrE" runat="server">
                                </asp:DropDownList>
                                季。
                            </div>--%>
                            <%-- 季繳提示訊息 by Hilty 2013/12/17 --%>
                            <%--<div id="divQtrText">
                                <font color="red">若需終止定期定額奉獻「季繳」請電洽-奉獻服務專線(02)8024-3911 捐款服務組</font>
                            </div>
                            <div id="divYear">
                                開始年：
                                <asp:DropDownList ID="ddlYearS" runat="server">
                                </asp:DropDownList>
                                年。
                                <br/>
                                結束年：
                                <asp:DropDownList ID="ddlYearE" runat="server">
                                </asp:DropDownList>
                                年。
                            </div>--%>
                        </td>
                    </tr>
<%--                    <tr>
                        <td height="10" colspan="2" align="center">
                            <div class="function">
                                <asp:Label runat="server" ID="lblGridList"></asp:Label>
                            </div>
                        </td>
                    </tr>--%>
                    <tr>
                        <td height="10" colspan="2" align="center">
                            <div class="function">
                                <asp:Button ID="btnNext" class="Online Online_Button3" runat="server" Width="90px"
                                    Text="Add to cart" OnClientClick="return Foolproof();" OnClick="btnNext_Click" />
                                <asp:Button ID="btnCancel" class="Online Online_Button5" runat="server" Width="174px"
                                    Text="Continue to donation items" OnClick="btnCancel_Click" />
                            </div>
                        </td>
                    </tr>
<%--                    <tr>
                        <td>
                            &nbsp;
                        </td>
                        <td>
                            &nbsp;
                        </td>
                    </tr>--%>
                   <%-- <tr>
                        <td align="right" bgcolor="#FFCC00">
                            &nbsp;&nbsp;&nbsp;備&nbsp;&nbsp;&nbsp;註：
                        </td>
                        <td>
                            1、線上刷卡、WEB ATM，每筆奉獻金額最低 100 元，最高60000元。
                        </td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                        <td>
                            2、7-11、ibon 、便利商店等電子帳單，每筆金額最低 500 元。
                        </td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                        <td>
                            3、奉獻金額請填寫純數字如 10000
                        </td>
                    </tr>--%>
                    <tr>
                        <td colspan="2">
                            <a href="donate_index.html">
                                <img alt="" src="images/pic_faith.jpg" width="650" height="100" border="0" /></a>
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
                <!--img alt="" src="images/page_botten.jpg" width="956" height="75" /-->
                <div id="footer">
                    <div id="address">
			            <a>財團法人加百列福音傳播基金會</a>
			            <a>　　TEL:02-8024-3911</a>
			            <a>　　FAX:02-8024-3938</a> 
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
    </table></asp:Panel>
    </form>
</body>
</html>
