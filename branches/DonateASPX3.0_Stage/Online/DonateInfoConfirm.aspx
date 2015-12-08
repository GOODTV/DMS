<%@ Page Language="C#" EnableEventValidation="false" CodeFile="DonateInfoConfirm.aspx.cs" Inherits="Online_DonateInfoConfirm" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>線上明細確認</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">

        function btnGoBack_Click() {
            document.forms[0].action = "CheckOut.aspx";
            document.forms[0].submit();
            return false;
            //history.back();
            //return false;
        }

    </script>
    <style type="text/css">

        body {
            font-size: 18px;
        }

        .table_h tr {
            color: black;
            height: 30px;
        }

        .table_h tr:hover {
            color: black;
            background-color: white;
        }

    </style>
</head>
<body style="margin: 0px; background-color: #EEE">    
    <form id="Form1" runat="server" >
    <asp:HiddenField ID="HFD_DonorID" runat="server" />
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_Amount" runat="server" />
    <asp:HiddenField ID="HFD_PayType" runat="server" />
    <asp:HiddenField ID="HFD_DonorName" runat="server" />
    <asp:HiddenField ID="HFD_LoginOK" runat="server" />

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
                            <strong><font color="#003399">確認您的奉獻明細</font></strong>
                        </td>
                        <td align="right" valign="bottom">
                            <!--<strong><font color="#003399">Language：</font><a href="../English/DonateInfoConfirm.aspx" target="_self"><font size="3">English</font></a></strong>-->
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
            <td colspan="3" align="center">
                <br />
                <table align="center" border="0" cellpadding="0" cellspacing="1">
                <tr>
                    <td align="center" style="white-space: nowrap;">
                        <span>我目前的位置：</span>
                        <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="MediumBlue" />
                        <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="MediumBlue" />
                        <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="Red" Font-Bold="True" />
                        <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="MediumBlue" />
                        <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="MediumBlue" />
                    </td>
                </tr>
                </table>
                <br />               
                
                <table align="center" border="0" cellpadding="0" cellspacing="1" class="table_h" width="80%">
                    <tr>
                        <td align="left">
                            <asp:Label ID="lblDonorName" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Label ID="lblPhoneNum" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Label ID="lblEmail" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Label ID="lblDonateContent" runat="server"></asp:Label>
                        </td>
                    </tr>
                </table><br/>
                <asp:Panel ID="msgOnce" runat="server" Visible="false">
                    <table align="center" style="padding-top: 10px; background-color: #FFFFFF;" border="0" width="70%" cellspacing="2" cellpadding="2">
                        <tr>
                            <td colspan="2">
                                <span style="width: 70%">信用卡是與國泰世華銀行合作，接受VISA,Master,JCB及銀聯信用卡，使用思遠資訊之iePay金流閘道，並同時提供其他捐款管道選擇。</span></td>
                        </tr>
                        <tr>
                            <td>
                                <img alt="" src="../images/payment_gateway.jpg" />
                            </td>
                            <td>
                                <img style="height: 40px;" alt="" src="../images/visa.jpg" />
                                <img style="height: 40px;" alt="" src="../images/MasterCard.png" />
                                <img style="height: 40px;" alt="" src="../images/UnionPay.png" />
                                <img style="height: 40px;" alt="" src="../images/JCB_Cards_svg.png" />
                                <img style="height: 40px;" alt="" src="../images/small_amex.gif" />
                                <img style="height: 40px;" alt="" src="../images/7-11.png" />
                                <img style="height: 40px;" alt="" src="../images/family.png" />
                                <img style="height: 40px;" alt="" src="../images/hi_life.png" />
                                <img style="height: 40px;" alt="" src="../images/ok_mark.png" />
                                <img style="height: 40px;" alt="" src="../images/1111.png" />
                                <img style="height: 40px;" alt="" src="../images/ibon.png" />
                                <img style="height: 40px;" alt="" src="../images/famipost.png" />
                            </td>
                        </tr>
                    </table>
                    <br />
                </asp:Panel>
                <asp:Panel ID="msgPeriod" runat="server" Visible="false">
                        <table width="80%">
                            <tr>
                                <td valign="top">1.</td>
                                <td>定期定額捐款扣款方式：捐款金額月繳是每月固定25日進行授權扣款，季繳為每三個月扣款一次，以此類推。</td>
                            </tr>
                            <tr>
                                <td valign="top">2.</td>
                                <td>若需終止定期定額奉獻，請電洽-奉獻服務專線(02)8024-3911 捐款服務組，謝謝您的支持。</td>
                            </tr>
                        </table>
                </asp:Panel>
            </td>
        </tr>
        <tr>
            <td height="10" colspan="2">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center">
                <table width="80%">
                    <tr>
                        <td>
                            <asp:Button ID="btnGoHome" class="css_btn_class" runat="server" Text="回上一頁"
                                OnClientClick="btnGoBack_Click();" /></td>
                        <td align="right">
                            <asp:Button ID="btnConfirm" class="css_btn_class" runat="server" Text="確認，開始付款"
                                OnClick="btnConfirm_Click" /></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
