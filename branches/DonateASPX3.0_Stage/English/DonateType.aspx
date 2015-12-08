<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateType.aspx.cs"
    Inherits="Online_DonateType" %>

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
        <asp:Literal ID="litTitle" runat="server">Good TV 線上奉獻 - 確認奉獻方式</asp:Literal>
    </h1>
    <table align="center" border="0" cellpadding="0" cellspacing="1" class="table_v"
        width="100%">
        <tr>
            <td align="center" colspan="2" style="background-color: #D3D3D3; color: Black; font-weight: bold">
                選　擇／確　認　奉　獻　方　式
            </td>
        </tr>
        <tr>
            <td colspan="2" style="color: #8B0000;">
                <asp:Label ID="lblTitle" runat="server">平安שלום שלום！歡迎您使用線上奉獻，請確認奉獻方式：</asp:Label>
            </td>
        </tr>
        <tr>
            <th align="right" style="width: 30mm">
                線上刷信用卡：
            </th>
            <td>
                <asp:Label runat="server" ID="lblType1"></asp:Label>
            </td>
        </tr>
<%--        <tr>
            <th align="right" style="width: 30mm">
                Web ATM：
            </th>
            <td>
                <asp:Label runat="server" ID="lblType2"></asp:Label>
            </td>
        </tr>--%>
    </table>
    <div class="function">
        <asp:Button ID="btnPrev" class="npoButton npoButton_PrevStep" runat="server" 
            Text="上一步" onclick="btnPrev_Click"/>
        <asp:Button ID="btnConfirm" class="npoButton npoButton_Submit" runat="server" 
            Text="確認" onclick="btnConfirm_Click" Width="80px"/>
    </div>
    </form>
</body>
</html>
