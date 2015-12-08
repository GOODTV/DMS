<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ShoppingCart.aspx.cs" Inherits="Online_ShoppingCart" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>線上奉獻清單</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            //有奉獻項目才顯示訊息
            if ($('#HFD_Mode').val() == '' &&
                        ($('#lblGridOnce').html() != '' || $('#lblGridPeriod').html() != '') && 
                            // 用來判斷到是否有新增新的奉獻(有文字=有新增奉獻)，到ShoppingCart.aspx頁面時是否要跳出"已加入奉獻清單"的提示訊息 by Hilty 2013/12/18
                            $('#AddShopping').val() != '') {
                $('#reg').show();
                $('#regMsg').show();
                if ($('#HFD_DelOnce').val() != '' || $('#HFD_DelPeriod').val() != '') {
                    $('#regMsg').html('<br/><br/><br/><br/>奉獻項目已刪除。');
                }
                //背景半透明效果
                $("#reg").fadeTo("fast", 0.6);
                //1.0秒後關閉
                selfAlert(1000);
            }

            $('#ItemPeriod td:nth-child(2)').css('text-align', 'center');
            $('#ItemPeriod td:nth-child(3)').css('text-align', 'right');
            $('#ItemPeriod td:nth-child(4)').css('text-align', 'center');

            $("#ItemOnce tr:first").css('background-color', '#CDE1E2');
            $('#ItemOnce td:nth-child(2)').css('text-align', 'right');
            $('#ItemOnce td:nth-child(3)').css('text-align', 'center');
            $("#ItemOnce tr:last>td:last").hide();
            $("#ItemOnce tr:last").css('background-color', '#CDE1E2');
        });            //end of ready()

        function selfAlert(timer) {
            remove = function () {
                $('#reg').hide();
                $('#regMsg').hide();
            }
            setTimeout("remove()", timer);
        }

    </script>
    <style type="text/css">
<!--
body {
	background-color: #EEE;
	text-align: center;
}
a {
	font-size: 12px;
}
body,td,th {
	font-size: 14px;
	color: #000;
	font-weight: bold;
}
-->
</style>
    <script type="text/javascript">
<!--
        function MM_swapImgRestore() { //v3.0
            var i, x, a = document.MM_sr; for (i = 0; a && i < a.length && (x = a[i]) && x.oSrc; i++) x.src = x.oSrc;
        }
        function MM_preloadImages() { //v3.0
            var d = document; if (d.images) {
                if (!d.MM_p) d.MM_p = new Array();
                var i, j = d.MM_p.length, a = MM_preloadImages.arguments; for (i = 0; i < a.length; i++)
                    if (a[i].indexOf("#") != 0) { d.MM_p[j] = new Image; d.MM_p[j++].src = a[i]; }
        }
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
//-->
    </script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <%--<body class="body">--%>
    <div id="reg" style="position: fixed; z-index: 100; top: 0px; left: 0px; height: 100%;
        width: 100%; background: #000; display: none;">
    </div>
    <div id="regMsg" style="background: #ffffff; border-radius: 5px 5px 5px 5px; color: blue;
        font-size: large; display: none; padding-bottom: 2px; width: 500px; height: 200px;
        z-index: 11001; left: 30%; position: fixed; text-align: center; top: 200px;">
        <br />
        <br />
        <br />
        <br />
        已加入奉獻清單
    </div>
    <form id="Form1" runat="server">
    <asp:HiddenField ID="HFD_ItemList" runat="server" />
    <asp:HiddenField ID="HFD_DelOnce" runat="server" />
    <asp:HiddenField ID="HFD_DelPeriod" runat="server" />
    <asp:HiddenField ID="HFD_Mode" runat="server" />
    <%-- 用來判斷到是否有新增新的奉獻(有文字=有新增奉獻)，到ShoppingCart.aspx頁面時是否要跳出"已加入奉獻清單"的提示訊息 by Hilty 2013/12/18 --%>
    <asp:HiddenField ID="AddShopping" runat="server" />
    <table width="860" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td>
                &nbsp;
            </td>
        </tr>       
        <tr>
            <td width="436">
                <img alt="aa" src="images/bar_logo.jpg" width="390" height="40" />
            </td>
            <td width="424" style="background-image:url('images/bar_item.jpg')">
                <table width="95%" border="0" align="center" cellpadding="5" cellspacing="5">
                    <tr>
                        <td align="left" valign="bottom">
                            <strong><font color="#003399">我的奉獻明細</font></strong>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td align="left" valign="bottom" height="30" colspan="3">
            <asp:Label ID="lblCart" runat="server"></asp:Label>
			</td>
        </tr>
        <tr valign="top">
            <td height="405" colspan="3" style="background-image:url('images/bk_fish.jpg')" align="center">
                            <br />
                            
                <table width="70%">
                    <tr>
                        <td>
                            <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="Brown" />
                            <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="Chocolate" />
                            <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="Chocolate" />
                            <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="Chocolate" />
                            <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="Chocolate" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <br />
                            <asp:Label ID="lblGridOnce" runat="server"></asp:Label>
                            <br />
                            <asp:Label ID="lblGridPeriod" runat="server"></asp:Label>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td height="10" colspan="2">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                        <td width="133">
							<asp:Button ID="btnContinue" runat="server" class="npoButton npoButton_Submit" Text="繼續奉獻"
                                OnClick="btnContinue_Click" Width="120px" height="30px" />
                            <%--<asp:ImageButton ID="btnContinue" runat="server" OnClick="btnContinue_Click" ImageUrl="images/btn_繼續奉獻_O.jpg"
                                onmouseover="this.src='images/btn_繼續奉獻_down.jpg'" onmouseout="this.src='images/btn_繼續奉獻_O.jpg'" />--%>
                        </td>
                         <td width="10">
                            &nbsp;
                        </td>
                        <td width="154">
                            <asp:Button ID="btnCheckOut" runat="server" class="npoButton npoButton_Submit" Text="下一步"
                                OnClick="btnCheckOut_Click" Width="120px" height="30px" />
                            <%--<asp:ImageButton ID="btnCheckOut" runat="server" OnClick="btnCheckOut_Click" ImageUrl="images/btn_完成奉獻_O.jpg"
                                onmouseover="this.src='images/btn_完成奉獻_down.jpg'" onmouseout="this.src='images/btn_完成奉獻_O.jpg'" />--%>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
