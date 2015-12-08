<%@ Page Language="C#" EnableEventValidation="false" AutoEventWireup="true" CodeFile="VerificationCode.aspx.cs" Inherits="Online_VerificationCode" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>線上奉獻取得驗證碼</title>
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <style type="text/css">

        body {
            font-size: 18px;
        }

        .table_v {
            font-size: 18px;
        }

        .table_v th {
            line-height: 30px;
        }

        .table_v td {
            line-height: 30px;
        }

        .show {
            color: blue;
        }

    </style>
    <script type="text/javascript">

        $(document).ready(function () {

            $('#txtEmail').focus(function () {
                $('#SendOK').hide();
            });

        });

        function GoBack() {

            $("#HFD_DonorID").val('')
            document.forms[0].action = "ShoppingCart.aspx";
            document.forms[0].submit();

        }

        function getToDay() {

            var Today = new Date();
            return Today.getFullYear().toString() + (Today.getMonth() + 1).toString() + Today.getDate().toString();

        }

        function CheckValidate(sender, e) {
            if (e.Value == $("#HFD_code").val() / getToDay())
                e.IsValid = true;
            else
                e.IsValid = false;
        }

    </script>
</head>
<body style="background-color: #EEE">
    <form id="form1" runat="server">
    <asp:HiddenField ID="HFD_DonorID" runat="server" />
    <asp:HiddenField ID="HFD_DonorName" runat="server" />
    <asp:HiddenField ID="HFD_code" runat="server" />
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_Amount" runat="server" />
    <asp:HiddenField ID="HFD_PayType" runat="server" />
    <asp:HiddenField ID="HFD_Email" runat="server" />
       <table width="860" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
               <td>&nbsp;</td>
           </tr>
            <tr>
                <td width="436">
                    <img alt="aa" src="images/bar_logo.jpg" width="390" height="40" />
                </td>
                <td width="424" style="background-image: url('images/bar_item.jpg')">
                    <table width="95%" border="0" align="center" cellpadding="5" cellspacing="5">
                        <tr>
                            <td align="left" valign="bottom">
                                <strong><font color="#003399">取得驗證碼</font></strong>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <br />
        <div align="center" style="white-space: nowrap;">
            <span>我目前的位置：</span>
            <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="Red" Font-Bold="True" />
            <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="MediumBlue" />
            <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="MediumBlue" />
            <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="MediumBlue" />
            <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="MediumBlue" />
        </div>
        <br />
        <div style="text-align: center">
        請填入您的 E-mail，我們會把驗證碼寄送到帳號(Email)信箱，您得知驗證碼後，再輸入驗證碼與新密碼。
        </div>
        <br />
        <table border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td>
                    <table border="0" cellpadding="0" cellspacing="1" class="table_v">
                        <tr>
                            <td colspan="2" align="center">第一步：輸入E-mail，取得驗證碼。
                            </td>
                        </tr>
                        <tr>
                            <th align="right">
                                <span class="necessary">帳號(Email)：</span>

                            </th>
                            <td>
                                <asp:TextBox ID="txtEmail" runat="server" Width="300px" Font-Size="18px"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="rfvEmail" runat="server" ControlToValidate="txtEmail" ErrorMessage="必填" CssClass="blink" Display="Dynamic"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="revEmail" runat="server" ErrorMessage="格式錯誤" ControlToValidate="txtEmail" CssClass="blink" Display="Dynamic" ValidationExpression="[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" align="center">
                                <asp:Label ID="SendOK" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" align="center">
                                <asp:Button ID="SendCode" runat="server" class="css_btn_class" Text="取得驗證碼" OnClick="SendCode_Click"  />
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" align="center">第二步：從Email得知驗證碼，輸入驗證碼，輸入新密碼，完成更新密碼。
                            </td>
                        </tr>
                        <tr>
                            <th align="right">
                                <span class="necessary">輸入驗證碼：</span>

                            </th>
                            <td>
                                <asp:TextBox ID="txtCode" runat="server" Width="300px" Font-Size="18px" Enabled="False"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="rfvCode" runat="server" ControlToValidate="txtCode" ErrorMessage="必填" CssClass="blink" Display="Dynamic" Enabled="False"></asp:RequiredFieldValidator>
                                <asp:CustomValidator ID="cvCode" runat="server" ErrorMessage="不對" ClientValidationFunction="CheckValidate" ControlToValidate="txtCode" CssClass="blink" Enabled="False" Display="Dynamic"></asp:CustomValidator>
                            </td>
                        </tr>
                        <tr>
                            <th align="right"><span class="necessary">新密碼：</span></th>
                            <td>
                                <asp:TextBox ID="txtDonorPwd1" runat="server" Width="300px" TextMode="Password" Font-Size="18px" Enabled="False"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="rfvDonorPwd1" runat="server" ControlToValidate="txtDonorPwd1" ErrorMessage="必填" CssClass="blink" Display="Dynamic" Enabled="False"></asp:RequiredFieldValidator>
                                <asp:CompareValidator ID="cvDonorPwd1" runat="server" ErrorMessage="不相符" ControlToCompare="txtDonorPwd1" ControlToValidate="txtDonorPwd2" Display="Dynamic" CssClass="blink" Enabled="False"></asp:CompareValidator>
                            </td>
                        </tr>
                        <tr>
                            <th align="right"><span class="necessary">確認密碼(再次輸入新密碼)：</span></th>
                            <td>
                                <asp:TextBox ID="txtDonorPwd2" runat="server" Width="300px" TextMode="Password" Font-Size="18px" Enabled="False"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="rfvDonorPwd2" runat="server" ControlToValidate="txtDonorPwd2" ErrorMessage="必填" CssClass="blink" Display="Dynamic" Enabled="False"></asp:RequiredFieldValidator>
                                <asp:CompareValidator ID="cvDonorPwd2" runat="server" ErrorMessage="不相符" ControlToCompare="txtDonorPwd1" ControlToValidate="txtDonorPwd2" Display="Dynamic" CssClass="blink" Enabled="False"></asp:CompareValidator>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" align="center">
                                <asp:Label ID="UpdateOK" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" align="center">
                                <asp:Button ID="UpdatePwd" runat="server" class="css_btn_class" Text="更改密碼" OnClick="UpdatePwd_Click" Visible="False" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr><td>&nbsp;</td></tr>
            <tr>
                <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td>&nbsp;
                            </td>
                            <td width="100">
                            </td>
                            <td width="10">&nbsp;
                            </td>
                            <td width="100">
                                <asp:Button ID="Button2" runat="server" class="css_btn_class" Text="回上一頁" OnClientClick="GoBack();return false;" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

    </form>
</body>
</html>
