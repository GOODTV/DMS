<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateOnlineDefault.aspx.cs"
    Inherits="Online_DonateOnlineDefault" %>

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
    <script type="text/javascript">
        $(document).ready(function () {

        });                     //end of ready()

    </script>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_ItemList" runat="server" />
    <h1>
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        <asp:Literal ID="litTitle" runat="server">Good TV 線上奉獻</asp:Literal>
    </h1>
    <table align="center" border="0" cellpadding="0" cellspacing="1" class="table_v"
        width="100%">
        <tr>
            <td align="center" colspan="2" style="background-color: #D3D3D3; color: Black; font-weight: bold">
                線 上 奉 獻
            </td>
        </tr>
        <tr>
            <td colspan="2" style="color: #8B0000;">
                <asp:Label ID="lblTitle" runat="server">平安שלום שלום！歡迎您使用線上奉獻，請選擇奉獻項目：</asp:Label>
            </td>
        </tr>
        <tr>
            <th align="right" style="width: 30mm">
                奉獻項目：
            </th>
            <td valign="middle">
                <table>
                    <tr>
                        <td>
                            <a href="DonateOnlineUsual.aspx">
                            <asp:Label runat="server" ID="lblItem1"></asp:Label>
                            </a>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <a href="DonateOnlineMedia.aspx">
                                <asp:Label runat="server" ID="lblItem2"></asp:Label>
                            </a>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <a href="DonateOnlineOther.aspx">
                            <asp:Label runat="server" ID="lblItem3"></asp:Label>
                            </a>
                        </td>
                    </tr>
                </table>
                <br />
<%--                <span style="color: #8B0000;">備註
                    <br />
                    1：線上刷卡、WEB ATM，每筆金額最低 200 元，最高60000元。
                    <br />
                    2：7-11、ibon 、便利商店等電子帳單，每筆金額最低 500 元。
                    <br />
                    3：奉獻金額請填寫純數字如 10000 </span>--%>
            </td>
        </tr>
        <tr>
            <td colspan="2">
                <asp:Label ID="lblGridList" runat="server"></asp:Label>
            </td>
        </tr>
    </table>
    <div class="function">
        <asp:Button ID="btnCheckOut" class="npoButton npoButton_Submit" runat="server" 
            Text="完成奉獻" onclick="btnCheckOut_Click" Width="121px"/>
    </div>
    </form>
</body>
</html>
