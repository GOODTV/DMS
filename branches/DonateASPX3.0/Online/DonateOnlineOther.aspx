<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateOnlineOther.aspx.cs" Inherits="Online_DonateOnlineOther" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>線上奉獻</title>
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

            $('input:checkbox[id*=CHK_Item]').bind("click", function (e) {
                $('#HFD_chkItem').val($('input:checkbox:checked[name="Item"]').getValue());
                var itemList = GetItemList();
                $('#lblGridList').html(itemList);
                var s = $(e.target).attr('ID');
                if (!$(e.target).attr('checked')) {
                    $('#txtAmount' + s.substr(s.length - 1, 1)).val('');
                }
            });

            $('input:radio[id*=rdoPayType]').bind("click", function (e) {
                PayTypeChange();
            });

            $('input:text[id*=txt]').bind("focusout", function (e) {
                $('#HFD_chkItem').val($('input:checkbox:checked[name="Item"]').getValue());
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

            $('#divMonth').hide();
            $('#divQtr').hide();
            $('#divYear').hide();
        });                     //end of ready()

    </script>
    <style type="text/css">
        body
        {
            background-color: #CCC;
            margin-top: 0px;
            text-align: center;
        }
        .alignright
        {
            font-size: 12px;
            color: #000;
            text-align: right;
        }
        body, td, th
        {
            font-size: 12px;
            color: #000;
            text-align: left;
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
</head>
<body onload="MM_preloadImages('images/btn_經常奉獻_don.jpg','images/btn_優質媒體奉獻_don.jpg','images/btn_首頁_don.jpg','images/btn_其他_don.jpg','images/btn_線上查詢_don.jpg')">
    <form id="Form1" runat="server">
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_ItemList" runat="server" />
    <asp:HiddenField ID="HFD_PayType" runat="server" />
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
                                <img src="images/btn_首頁_on.jpg" alt="首頁" name="Image6" width="120" height="26" border="0"
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
                            <img src="images/btn_其他_don.jpg" width="120" height="35" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <a href="DonateQry.aspx" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image7','','images/btn_線上查詢_don.jpg',1)">
                                <img alt="線上查詢頁面" src="images/btn_線上查詢_on.jpg" name="Image7" width="120" height="35" border="0"
                                    id="Image7" /></a>
                        </td>
                    </tr>
                    <tr>
                        <td valign="top">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td valign="top">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
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
                <table width="700" border="0" align="left" cellpadding="1" cellspacing="1">
                    <tr>
                        <td style="text-align:center" colspan="2" align="center" valign="middle" bgcolor="#5F8FB3">
                            <font color="#FFFFFF" size="+1">其　他</font>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" align="center" valign="middle" bgcolor="#CCCCCC">
                        <font color="#FF0000">
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;「他又對他們說，你們往普天下去，傳福音給萬民聽。」（馬十六15）
                        </font>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" align="left" valign="top" bgcolor="#DDE8F9">
                            ※我對
                            <input type="text" name="item" id="item" />
                            有負擔，說明：
                            <textarea name="word2" id="word2" cols="50" rows="3"></textarea>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" align="center" valign="middle">
                            &nbsp;
                        </td>
                    </tr>
<%--                    <tr>
                        <td colspan="2" style="color: #8B0000;">
                            <asp:Label ID="lblTitle" runat="server">平安שלום שלום！歡迎您使用線上奉獻，請填寫奉獻內容：</asp:Label><br />
                        </td>
                    </tr>--%>
                    <tr>
                        <td width="112" align="center" valign="middle">
                            <asp:Label runat="server" ID="lblItemOnce"></asp:Label>：
                        </td>
                        <td width="581">
                            新台幣
                            <asp:TextBox ID="txtAmountOnce" runat="server" Width="40mm" Style="text-align: right"></asp:TextBox>
                            元
                        </td>
                    </tr>
                    <tr>
                        <td height="10" colspan="2" align="center">
                            <hr width="100%" size="1" />
                        </td>
                    </tr>
                    <tr>
                        <td align="center">
                            <asp:Label runat="server" ID="lblItemPeriod"></asp:Label>：
                        </td>
                        <td>
                            新台幣
                            <asp:TextBox ID="txtAmountPeriod" runat="server" Width="40mm" Style="text-align: right"></asp:TextBox>
                            元，繳費方式：
                            <asp:Label runat="server" ID="lblPayType"></asp:Label>
                            <br />
                            <div id="divMonth">
                                開始年月：
                                <asp:DropDownList ID="ddlYearMS" runat="server">
                                </asp:DropDownList>
                                年
                                <asp:DropDownList ID="ddlMonthS" runat="server">
                                </asp:DropDownList>
                                月，結束年月：
                                <asp:DropDownList ID="ddlYearME" runat="server">
                                </asp:DropDownList>
                                年
                                <asp:DropDownList ID="ddlMonthE" runat="server">
                                </asp:DropDownList>
                                月
                            </div>
                            <div id="divQtr">
                                開始年季：
                                <asp:DropDownList ID="ddlYearQS" runat="server">
                                </asp:DropDownList>
                                年
                                <asp:DropDownList ID="ddlQtrS" runat="server">
                                </asp:DropDownList>
                                季，結束年季：
                                <asp:DropDownList ID="ddlYearQE" runat="server">
                                </asp:DropDownList>
                                年
                                <asp:DropDownList ID="ddlQtrE" runat="server">
                                </asp:DropDownList>
                                季
                            </div>
                            <div id="divYear">
                                開始年：
                                <asp:DropDownList ID="ddlYearS" runat="server">
                                </asp:DropDownList>
                                年，結束年：
                                <asp:DropDownList ID="ddlYearE" runat="server">
                                </asp:DropDownList>
                                年
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td align="center">
                            &nbsp;
                        </td>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td height="10" colspan="2" align="center">
                            <hr width="100%" size="1" />
                        </td>
                    </tr>
                    <tr>
                        <td align="right" bgcolor="#FCB84B">
                            &nbsp;&nbsp;&nbsp;備&nbsp;&nbsp;&nbsp;註：
                        </td>
                        <td height="5" align="center">
                            1、線上刷卡、WEB ATM，每筆金額最低 200 元，最高60000元。
                        </td>
                    </tr>
                    <tr>
                        <td height="0" align="right">
                            &nbsp;
                        </td>
                        <td height="10" align="center">
                            2、7-11、ibon 、便利商店等電子帳單，每筆金額最低 500 元。
                        </td>
                    </tr>
                    <tr>
                        <td height="10" align="right">
                            &nbsp;
                        </td>
                        <td height="10" align="center">
                            3、奉獻金額請填寫純數字如 10000
                        </td>
                    </tr>
                    <tr>
                        <td height="10" colspan="2" align="center">
                            <div class="function">
                                <%--<asp:Button ID="btnPrev" class="npoButton npoButton_PrevStep" runat="server" Text="上一步"
                                    OnClick="btnPrev_Click" />--%>
                                <asp:Button ID="btnNext" class="npoButton npoButton_NextStep" runat="server" Width="121px"
                                    Text="立即奉獻" OnClientClick="return CheckAmount();" OnClick="btnNext_Click" />
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td height="176" colspan="2" align="center">
                            <img src="images/pic_Other.jpg" width="700" height="100" />
                        </td>
                    </tr>
                </table>
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
</body>
</html>
