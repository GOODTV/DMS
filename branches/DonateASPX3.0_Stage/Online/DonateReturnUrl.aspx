<%@ Page Language="C#" EnableEventValidation="false" CodeFile="DonateReturnUrl.aspx.cs" Inherits="Online_DonateReturnUrl" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>信用卡奉獻交易確認</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            if ($("#HFD_Question").val() != "") {
                $('.Questionaire').show();
            }
        });
        </script>
</head>
<body style="background-color: #EEE" onkeydown="if(event.keyCode==116){return false;}">
    <form id="Form1" runat="server" >
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_ItemList" runat="server" />

    <asp:HiddenField ID="storeid" runat="server"  value=""/>
    <asp:HiddenField ID="password" runat="server" value=""/>
    <asp:HiddenField ID="orderid" runat="server" value=""/>
    <asp:HiddenField ID="account" runat="server" value=""/>
    <asp:HiddenField ID="remark" runat="server" value=""/>
    <asp:HiddenField ID="param" runat="server" value=""/>
    <asp:HiddenField ID="customer" runat="server" value=""/>
    <asp:HiddenField ID="cellphone" runat="server" value=""/>
    <asp:HiddenField ID="charset" runat="server" value=""/>
    <asp:HiddenField ID="invoiceflag" runat="server" value=""/>
    <asp:HiddenField ID="paytype" runat="server" value=""/>
    <asp:HiddenField ID="purpose" runat="server" value=""/>
    <asp:HiddenField ID="HFD_UID" runat="server" value=""/>
    <asp:HiddenField ID="HFD_OrderId" runat="server" value=""/>
    <asp:HiddenField ID="HFD_Question" runat="server" value=""/>

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
            <td width="424" style="background-image: url('images/bar_item.jpg')">
                <table width="95%" border="0" align="center" cellpadding="5" cellspacing="5">
                    <tr>
                        <td align="left" valign="bottom">
                            <strong><font color="#003399">信用卡奉獻交易確認 Your Donation Confirmation</font></strong>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td height="20" colspan="3">
            </td>
        </tr>
        <tr valign="top">
            <td height="405" colspan="3" align="center">
                <br />
                 <font color="red" size="4"><b>感謝您的奉獻！ Thank you for your support.</b></font>
                    <br/><br/><font color='black'><b>以下是您本次刷卡授權記錄的結果，您可以列印保留 
                    <br/>The following message is the status of your payment , you may print this page for your reference.
					<br/><br/>願神大大賜福與您! May the Lord bless you greatly.</b></font><asp:Label ID="lblTitle" runat="server" Style="color: #000000; font-weight:bold;"></asp:Label>
                <br />
                <br />
                <font color='black'><b>您有任何捐款問題請和我們聯絡：
                    <br />
                    If you have any questions regarding your donation, Please give us a call:
                            <br />
                    <br />
                    總機 Tel：02-8024-3911<br />
                    <br />
                    傳真 Fax：02-8024-3933
                            <br />
                    <br />
                    地址：235 新北市中和區中正路911號6樓 捐款服務組
                    <br />
                    6F., No.911 Zhongzheng Rd., Zhonghe Dist., New Taipei City 23544, Taiwan, R.O.C.</b></font>
            </td>
        </tr>
        <tr>
            <td colspan="2">
               &nbsp;
            </td>
        </tr>
        <tr class="Questionaire" style="display: none;" align="center">
            <td colspan="2">
               <table width="530" border="0" cellpadding="0" cellspacing="1" class="table_v2">
                    <tr>
                        <th colspan="3"><font color="blue">為更有效服務捐款人，請務必勾選下列問題，或惠賜寶貴意見</font></th>
                    </tr>
                    <tr>
                        <th align="center" style="width:150px;">奉獻動機(可複選)：</th>
                        <td>
                            <asp:CheckBox ID="DonateMotive1" runat="server" Text="支持媒體宣教大平台，可廣傳福音" /><br />
                            <asp:CheckBox ID="DonateMotive2" runat="server" Text="個人靈命得造就" /><br />
                            <asp:CheckBox ID="DonateMotive3" runat="server" Text="支持優質節目製作" /><br />
                            <asp:CheckBox ID="DonateMotive4" runat="server" Text="支持GOOD TV家庭事工" /><br />
                            <%--<asp:CheckBox ID="DonateMotive5" runat="server" Text="抵扣稅額" /><br />--%>
                            <asp:CheckBox ID="DonateMotive5" runat="server" Text="感恩奉獻" /><br />
                            <asp:CheckBox ID="DonateMotive6" runat="server" Text="其他" />(如有請寫入〝想對GOODTV說的話〞)
                        </td>
                        </tr>
                        <tr>
                        <th align="center" style="width:150px;">收看管道(可複選)：</th>
                        <td>
                            <asp:CheckBox ID="WatchMode1" runat="server" Text="GOOD TV電視頻道" />
                            <br />
                            <asp:CheckBox ID="WatchMode9" runat="server" Text="報章雜誌" />
                            <br />新媒體平台：
                            <asp:CheckBox ID="WatchMode2" runat="server" Text="官網" />
                            <asp:CheckBox ID="WatchMode3" runat="server" Text="Facebook" />
                            <asp:CheckBox ID="WatchMode4" runat="server" Text="Youtube" />
                            <br />平面：　　　
                            <asp:CheckBox ID="WatchMode5" runat="server" Text="好消息月刊" />
                            <asp:CheckBox ID="WatchMode6" runat="server" Text="GOOD TV簡介刊物" />
                            <br />親友推薦：　
                            <asp:CheckBox ID="WatchMode7" runat="server" Text="教會牧者" />
                            <asp:CheckBox ID="WatchMode8" runat="server" Text="親友" />
                        </td>
                    </tr>
                   <tr>
                        <th align="right" nowrap="nowrap">想對GOOD TV說的話：
                        </th>
                        <td>
                            <asp:TextBox ID="txtToGoodTV" runat="server" Width="90%" TextMode="MultiLine" Rows="5" Font-Size="16px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" align="center">
                            <asp:Button ID="btnSubmit" runat="server" Text="感謝您提供寶貴意見，送出"
                                OnClick="btnSubmit_Click" />
                        </td>
                    </tr>
                </table>         
            </td>
        </tr>
        <tr>
            <td colspan="2">
               &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center">
                <asp:Button ID="btnBackPay" class="css_btn_class" runat="server" Text="重新授權付款"
                    OnClick="btnBackPay_Click" Visible="false" />
                <asp:Button ID="btnConfirm" class="css_btn_class" runat="server" Text="回到奉獻首頁"
                    OnClick="btnConfirm_Click" />
                <%--<asp:Button ID="btnQuestion" class="css_btn_class" runat="server" Text="填寫奉獻問卷"
                    OnClick="btnQuestion_Click" Visible="false" />--%>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
