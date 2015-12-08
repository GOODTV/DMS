<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ResetPassword.aspx.cs" Inherits="Online_ResetPassword" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>線上奉獻重設密碼確認</title>
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <style type="text/css">

        .table_v {
            font-size: 14px;
        }

        .table_v th {
            line-height: 30px;
        }

    </style>
</head>
<body style="background-color: #EEE">
    <form id="form1" runat="server">
        <table width="860" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td width="436">
                    <img alt="aa" src="images/bar_logo.jpg" width="390" height="40" />
                </td>
                <td width="424" style="background-image: url('images/bar_item.jpg')">
                    <table width="95%" border="0" align="center" cellpadding="5" cellspacing="5">
                        <tr>
                            <td align="left" valign="bottom">
                                <strong><font color="#003399">重設密碼確認</font></strong>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <!--<div style="text-align: center">
            <ul>
                <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="Red" Font-Size="Medium" Font-Bold="true" />
                <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="MediumBlue" />
            </ul>
        </div>-->
        <br />
        <div style="text-align: center">
        <strong><font color="#003399">請填入您的 E-mail 及姓名，我們會把密碼送到帳號(Email)信箱。</font></strong>
        </div>
        <br />
        <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td>
                    <table width="600" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
                        <tr>
                            <th align="right">
                                <span class="necessary">姓名：</span>

                            </th>
                            <td>
                                <asp:TextBox ID="txtDonorName" runat="server" Width="300px"></asp:TextBox>

                            </td>
                        </tr>
                        <tr>
                            <th align="right">
                                <span class="necessary">帳號(Email)：</span>

                            </th>
                            <td>
                                <asp:TextBox ID="txtEmail" runat="server" Width="300px" TextMode="Email"></asp:TextBox>

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
                                <asp:Button ID="Button1" runat="server" class="Online Online_Button3" Text="送出" OnClick="Button1_Click" />
                            </td>
                            <td width="10">&nbsp;
                            </td>
                            <td width="100">
                                <asp:Button ID="Button2" runat="server" class="Online Online_Button5" Text="取消,關閉視窗" OnClientClick="window.close();" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

    </form>
</body>
</html>
