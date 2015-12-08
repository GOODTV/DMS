<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateQry.aspx.cs" Inherits="Online_DonateQry" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <%--<meta http-equiv="X-UA-Compatible" content="IE=edge" />--%>
    <title>GOODTV線上奉獻徵信錄</title>
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
    <script type="text/javascript" src="../include/jquery-1.7.2.min.js"></script>    
    <script type="text/javascript" src="../include/jquery-ui-1.8.18.custom/development-bundle/ui/jquery.ui.datepicker.js"></script>
    <script type="text/javascript" src="../include/jquery-ui-1.8.18.custom/development-bundle/ui/jquery.ui.core.js"></script>
    <script type="text/javascript" src="../include/jquery-ui-1.8.18.custom/development-bundle/ui/i18n/jquery.ui.datepicker-zh-TW.js"></script>
    <link type="text/css" rel="stylesheet" media="all" href="../include/jquery-ui-1.8.18.custom/development-bundle/themes/ui-lightness/jquery.ui.all.css" />
    
    <script type="text/javascript">
        $(document).ready(function () {

            //            //利用 js 阻擋滑鼠右鍵的預設行為
            //            $(this).bind("contextmenu", function(e) {
            //                    e.preventDefault();
            //                });
        });                     //end of ready()               
    </script>    
    <style type="text/css">
        body
        {
            background-color: #CCC;
            margin-top: -20px;
			padding-top: 20px;
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
        .alignright
        {
            text-align: right;
        }
        .aligncenter
        {
            text-align: center;
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
<%--<body class="body" oncut="return false" oncopy="return false" onselectstart="return false" style="align:center">--%>
<body oncut="return false" oncopy="return false" onselectstart="return false" onload="MM_preloadImages('images/btn_其他奉獻_down.jpg','images/btn_金融機構奉獻_down.jpg','images/btn_授權書奉獻_down.jpg','images/btn_線上奉獻_down.jpg')">
	<form id="Form1" runat="server"> 
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="btnQuery">
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
                                <input type="text" name="keyWord" value="" size="7" id="keyWord"  />
                                <asp:Button ID="submit" PostBackUrl="http://www.goodtv.tv/video/index.php/video/search" Text="節目搜尋" runat="server" />                               
                            </form>    --%>
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
        </table>
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <table width="955" border="0" align="center" cellpadding="0" cellspacing="0">                
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
                                    <a href="donateonlineall.aspx" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('donate_online','','images/btn_線上奉獻_down.jpg',1)">
                                        <img src="images/btn_線上奉獻_O.jpg" name="donate_online" alt="線上奉獻頁面" width="198" height="36" border="0" id="donate_online" /></a>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                     <img src="images/btn_線上徵信_down.jpg" alt="線上徵信" name="DonateQry" width="198" height="36" border="0" id="DonateQry" />
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
                    <td width="16" bgcolor="#FFFFFF">
                        &nbsp;
                    </td>
                    <td width="686" valign="top" bgcolor="#FFFFFF">
                        <table width="650" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                                <td colspan="4">
                                    <img alt="" src="images/Titlebar_線上徵信.jpg" width="650" height="25" />
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" bgcolor="#B5B09F">
                                    <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td colspan="2">
                                                &nbsp;
                                            </td>
                                        </tr>
                                        <tr>
                                            <td height="25" colspan="2" bgcolor="#EAE6D7">
                                                <font color="#FF0000" size="3px"><b>&nbsp;&nbsp;&nbsp;&nbsp;那賜種給撒種的，賜糧給人吃的，必多多加給你們種地的種子，</b></font>
                                            </td>
                                        </tr>
                                         <tr>
                                            <td height="25" colspan="2" bgcolor="#EAE6D7">
                                                <font color="#FF0000" size="3px"><b>&nbsp;&nbsp;&nbsp;&nbsp;又增添你們仁義的果子；叫你們凡事富足，可以多多施捨，</b></font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td height="25" colspan="2" bgcolor="#EAE6D7">
                                                <font color="#FF0000" size="3px"><b>&nbsp;&nbsp;&nbsp;&nbsp;就藉著我們使感謝歸於神。</b></font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td width="70%" align="right">
                                                &nbsp;
                                            </td>
                                            <td width="30%" height="20" align="right" bgcolor="#D2CAAA">
                                                <font color="#FF0000" size="3px">&nbsp;&nbsp;哥林多後書 9:10-11 </font>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" bgcolor="#B5B09F">
                                    &nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4">
                                    <table width="100%" border="1" cellpadding="2" cellspacing="2">
                                        <%--                                        <tr>
                                            <td bgcolor="#EAE6D7" align="right">
                                                類別 :
                                            </td>
                                            <td bgcolor="#EAE6D7" align="left">
                                                <input type="radio" name="radio" id="radio2" value="radio" />
                                                捐 款 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                <input type="radio" name="radio" id="radio3" value="radio" />
                                                物資捐贈
                                            </td>
                                        </tr>--%>
                                        <tr>
                                            <td width="14%" bgcolor="#EAE6D7" align="right">
                                                &nbsp;&nbsp;&nbsp;&nbsp;捐&nbsp;款&nbsp;期&nbsp;間：
                                            </td>
                                            <td width="86%" bgcolor="#EAE6D7" align="left">
                                                <asp:DropDownList ID="ddlYearMS" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlYearMS_SelectedIndexChanged">
                                                </asp:DropDownList>
                                                年， 請選擇
                                                <asp:DropDownList ID="ddlMonthS" runat="server">
                                                </asp:DropDownList>
                                                月 。
                                            </td>
                                        </tr>
                                        <tr>
                                            <td bgcolor="#EAE6D7" align="right">
                                                &nbsp;&nbsp;捐&nbsp;款&nbsp;人&nbsp;姓&nbsp;名：
                                            </td>
                                            <td bgcolor="#EAE6D7" align="left">
                                                <asp:TextBox ID="txtDonorName" runat="server" Width="60mm"></asp:TextBox>
                                                &nbsp;(至少需輸入2個字)
                                            </td>
                                        </tr>
                                        <%--                                        <tr>
                                            <td bgcolor="#EAE6D7" align="right">
                                                捐款人電話:
                                            </td>
                                            <td bgcolor="#EAE6D7" align="left">
                                                <input name="money" type="text" id="money" size="50" />
                                            </td>
                                        </tr>--%>
                                        <tr>
                                            <td colspan="2" bgcolor="#EAE6D7">
                                                <table width="5%" border="0" align="right" cellpadding="2" cellspacing="2">
                                                    <tr>
                                                        <td>
                                                            <asp:Button ID="btnQuery" runat="server" class="npoButton npoButton_Search" OnClick="btnQuery_Click"
                                                                Text="查 詢" Width="80px" />
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" bgcolor="#B5B09F" align="center">
                                    <asp:GridView ID="gdvQry" runat="server" CellPadding="4" Font-Size="16px" ForeColor="#333333"
                                        GridLines="None" OnDataBound="gdvQry_DataBound" Width="100%" ShowFooter="True"
                                        EmptyDataText="查無奉獻資料！">
                                        <AlternatingRowStyle BackColor="White" />
                                        <EditRowStyle BackColor="#2461BF" />
                                        <EmptyDataRowStyle ForeColor="Red" />
                                        <EmptyDataTemplate>
                                            查無奉獻資料！
                                        </EmptyDataTemplate>
                                        <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                        <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                        <RowStyle BackColor="#EFF3FB" />
                                        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                                        <SortedAscendingCellStyle BackColor="#F5F7FB" />
                                        <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                                        <SortedDescendingCellStyle BackColor="#E9EBEF" />
                                        <SortedDescendingHeaderStyle BackColor="#4870BE" />
                                    </asp:GridView>
                                </td>
                            </tr>
                            <tr>
                                <td width="13%" height="65">
                                    &nbsp;
                                </td>
                                <td width="87%" height="65" colspan="3">
                                    &nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td align="right">
                                    &nbsp;
                                </td>
                                <td colspan="3">
                                    <hr width="100%" />
                                </td>
                            </tr>
                            <tr>
                                <td align="right">
                                    &nbsp;
                                </td>
                                <td colspan="3">
                                    <b>財團法人加百列福音傳播基金會</b>
                                </td>
                            </tr>
                            <tr>
                                <td align="right">
                                    &nbsp;&nbsp;&nbsp;
                                </td>
                                <td colspan="3">
                                    奉獻服務專線：(02)8024-3911分機7601陳姐妹 &nbsp;&nbsp;&nbsp;&nbsp; 傳真: (02)8024-3933
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    &nbsp;
                                </td>
                                <td colspan="3">
                                    e-mail: ds@goodtv.tv
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    &nbsp;
                                </td>
                                <td colspan="3">
                                    &nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4">
                                    <a href="donate_index.html">
                                        <img alt="" src="images/pic_Authorization_botton.jpg" width="650" height="100" border="0" /></a>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4">
                                </td>
                            </tr>
                        </table>
                    </td></table>
        </ContentTemplate>
    </asp:UpdatePanel></asp:Panel>
    <%--</td> </tr>--%>
    <table width="949" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="100" colspan="5" align="right" valign="bottom" bgcolor="#FFFFFF">
            <img alt="" src="images/page_botten.jpg" width="956" height="75" />
        </td>
    </tr>
    </table>
    </form>
</body></html>
<%--<body oncut="return false" oncopy="return false" onselectstart="return false" onload="MM_preloadImages('images/btn_經常奉獻_don.jpg','images/btn_優質媒體奉獻_don.jpg','images/btn_首頁_don.jpg','images/btn_其他_don.jpg','images/btn_線上查詢_don.jpg')">
    <form id="Form1" runat="server">
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_ItemList" runat="server" />
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <table width="937" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td colspan="5">
                <img src="images/title_1.jpg" width="955" height="222" alt="title" />
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF">
                &nbsp;
            </td>
            <td width="30" align="left" bgcolor="#FFFFFF">
                &nbsp;
            </td>
            <td width="200" align="left" valign="top" bgcolor="#FFFFFF">
                <table width="200%" border="0" cellspacing="1" cellpadding="1">
                    <tr>
                        <td>
                            <a href="donate_index.html" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image6','','images/btn_首頁_don.jpg',1)">
                                <img src="images/btn_首頁_on.jpg" name="Image6" width="120" height="26" border="0"
                                    id="Image6" /></a>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <a href="DonateOnlineUsual.aspx" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('btn_2','','images/btn_經常奉獻_don.jpg',1)">
                                <img src="images/btn_經常奉獻_on.jpg" alt="經常費頁面" name="btn_2" width="120" height="36"
                                    border="0" id="btn_2" /></a>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <a href="DonateOnlineMedia.aspx" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('btn_3','','images/btn_優質媒體奉獻_don.jpg',1)">
                                <img src="images/btn_優質媒體奉獻_on.jpg" alt="優質媒體" name="btn_3" width="120" height="35"
                                    border="0" id="btn_3" /></a>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <a href="DonateOnlineOther.aspx" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image7','','images/btn_其他_don.jpg',1)">
                                <img src="images/btn_其他_on.jpg" alt="other" name="Image7" width="120" height="35"
                                    border="0" id="Image7" /></a>
                        </td>
                    </tr>
                    <tr>
                        <td height="36">
                            <img src="images/btn_線上查詢_don.jpg" width="120" height="36" />
                        </td>
                    </tr>
                    <tr>
                        <td valign="top">
                            &nbsp;
                        </td>
                    </tr>
                    <tr class="alignleft">
                        <td height="213" valign="top">
                            <table width="180" border="0" cellspacing="0" cellpadding="5">
                                <tr>
                                    <td height="5" colspan="2" align="right" bgcolor="#80AEA6">
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td width="60" align="right" bgcolor="#E1ECEA">
                                        奉獻熱線：
                                    </td>
                                    <td width="100" bgcolor="#E1ECEA">
                                        02-8024-3911
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" bgcolor="#E1ECEA">
                                        郵撥帳號：
                                    </td>
                                    <td bgcolor="#E1ECEA">
                                        19113367
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" valign="top" bgcolor="#E1ECEA">
                                        戶 名：
                                    </td>
                                    <td bgcolor="#E1ECEA">
                                        財團法人加百列福音傳播基金會
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" valign="top" bgcolor="#E1ECEA">
                                        地 址：
                                    </td>
                                    <td bgcolor="#E1ECEA">
                                        新北市中和區中正路911號6樓
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" bgcolor="#E1ECEA">
                                        傳 真：
                                    </td>
                                    <td bgcolor="#E1ECEA">
                                        (02) 8024-3938
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" bgcolor="#E1ECEA">
                                        &nbsp;
                                    </td>
                                    <td bgcolor="#E1ECEA">
                                        &nbsp;
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
            <td width="700" align="left" valign="top" bgcolor="#FFFFFF">
                <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                    <ContentTemplate>
                        <table width="700" border="0" align="left" cellpadding="1" cellspacing="1">
                            <tr>
                                <td colspan="2" align="center" valign="middle" bgcolor="#5F8FB3">
                                    <font color="#FFFFFF" size="+1">線　上　徵　信</font>
                                </td>
                            </tr>
                            <tr class="alignleft">
                                <td height="50" colspan="2" align="center" valign="middle" bgcolor="#FFFFFF">
                                    <font color="#FF0000" size="medium">&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&#8226; 你們要給人，就必有給你們的，並且用十足的升斗，連搖帶按，上尖下流的倒在你們懷裡；因為你們用什麼量器量給人，也必用什麼量器量給你們。 - 路加福音六章38節</font>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="center" valign="middle">
                                    &nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td width="290" colspan="2" >
                                    奉獻期間:
                                    <asp:DropDownList ID="ddlYearMS" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlYearMS_SelectedIndexChanged">
                                    </asp:DropDownList>
                                    年， 請選擇
                                    <asp:DropDownList ID="ddlMonthS" runat="server">
                                    </asp:DropDownList>
                                    月
                                </td>
                            </tr>
                            <tr>
                                <td height="10" colspan="2" align="center">
                                    <hr width="100%" size="1" />
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" bgcolor="#E8EAEA">
                                    &nbsp;&nbsp;&nbsp; 奉獻姓名/團體:
                                    <asp:TextBox ID="txtDonorName" runat="server" Width="60mm"></asp:TextBox>
                                    <asp:Button ID="btnQuery" runat="server" class="npoButton npoButton_Search" OnClick="btnQuery_Click"
                                        Text="奉獻徵信" Width="100px" />
                                </td>
                            </tr>
                            <tr>
                                <td height="10" colspan="2" align="center">
                                    <hr width="100%" size="1" />
                                </td>
                            </tr>
                            <tr>
                                <td height="10" colspan="2" align="center">
                                    &nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2" >
                                    <asp:GridView ID="gdvQry" runat="server" AllowPaging="False" CellPadding="4" Font-Size="16px"
                                        ForeColor="#333333" GridLines="None" OnDataBound="gdvQry_DataBound" 
                                        Width="90%" ShowFooter="True" EmptyDataText="查無奉獻資料！">
                                        <AlternatingRowStyle BackColor="White" />
                                        <EditRowStyle BackColor="#2461BF" />
                                        <EmptyDataRowStyle ForeColor="Red" />
                                        <EmptyDataTemplate>
                                            查無奉獻資料！
                                        </EmptyDataTemplate>
                                        <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                        <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                        <RowStyle BackColor="#EFF3FB" />
                                        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                                        <SortedAscendingCellStyle BackColor="#F5F7FB" />
                                        <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                                        <SortedDescendingCellStyle BackColor="#E9EBEF" />
                                        <SortedDescendingHeaderStyle BackColor="#4870BE" />
                                    </asp:GridView>
                                </td>
                            </tr>
                            <tr>
                                <td height="176" colspan="2" align="center">
                                    <img alt="" src="images/pic_check.jpg" width="700" height="100" />
                                </td>
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center" bgcolor="#FFFFFF">
                &nbsp;
            </td>
            <td bgcolor="#FFFFFF">
                &nbsp;
            </td>
            <td bgcolor="#FFFFFF">
                &nbsp;
            </td>
            <td bgcolor="#FFFFFF">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="5" bgcolor="#FFFFFF">
                &nbsp;
            </td>
        </tr>
    </table>
    </form>
</body>--%>
<%--</html>--%>
