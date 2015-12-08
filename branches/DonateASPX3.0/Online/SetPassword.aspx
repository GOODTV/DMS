<%@ Page Language="C#" AutoEventWireup="true" CodeFile="SetPassword.aspx.cs" Inherits="Online_SetPassword" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>線上奉獻重設密碼輸入</title>
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
</head>
<body style="background-color: #EEE">
    <form id="form1" runat="server">
        <div id="setpwd" runat="server">
            <table width="860" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="436">
                        <img alt="aa" src="images/bar_logo.jpg" width="390" height="40" />
                    </td>
                    <td width="424" style="background-image: url('images/bar_item.jpg')">
                        <table width="95%" border="0" align="center" cellpadding="5" cellspacing="5">
                            <tr>
                                <td align="left" valign="bottom">
                                    <strong><font color="#003399">重設密碼輸入</font></strong>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
            <div style="text-align: center">
                <ul>
                    <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="MediumBlue" />
                    <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="Red" Font-Size="Medium" Font-Bold="true" />
                    <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="MediumBlue" />
                    <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="MediumBlue" />
                    <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="MediumBlue" />
                </ul>
            </div>

            <div>
                <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                        <td>
                            <table width="600" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
                                <tr>
                                    <th align="right">捐款人編號：</th>
                                    <td>
                                        <asp:Label ID="TxtDonorId" runat="server"></asp:Label>

                                    </td>
                                </tr>
                                <tr>
                                    <th align="right">姓名：</th>
                                    <td>
                                        <asp:Label ID="txtDonorName" runat="server"></asp:Label>

                                    </td>
                                </tr>
                                <tr>
                                    <th align="right">帳號(Email)：</th>
                                    <td>
                                        <asp:Label ID="txtEmail" runat="server"></asp:Label>

                                    </td>
                                </tr>
                                <tr>
                                    <th align="right"><span class="necessary">密碼(Password)：</span></th>
                                    <td>
                                        <asp:TextBox ID="txtDonorPwd1" runat="server" Width="200px" TextMode="Password"></asp:TextBox>

                                    </td>
                                </tr>
                                <tr>
                                    <th align="right"><span class="necessary">確認密碼(Confirm password)：</span></th>
                                    <td>
                                        <asp:TextBox ID="txtDonorPwd2" runat="server" Width="200px" TextMode="Password"></asp:TextBox>

                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
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
                                        <asp:Button ID="Button2" runat="server" class="Online Online_Button3" Text="取消" OnClick="Button2_Click" OnClientClick="javascript:if(!confirm('確定要取消嗎?'))return false;" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
    </form>
</body>
</html>
