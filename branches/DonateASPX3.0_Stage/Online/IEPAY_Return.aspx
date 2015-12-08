<%@ Page Language="C#" AutoEventWireup="true" CodeFile="IEPAY_Return.aspx.cs" Inherits="Online_IEPAY_Return" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>金流丟參數</title>
</head>
<body>
    <form id="Form1" runat="server">
    <div>
        訂單編號 orderid：<asp:TextBox ID="orderid" runat="server"></asp:TextBox><br />
        授權商店代碼 storeid：<asp:TextBox ID="storeid" runat="server"></asp:TextBox><br />
        捐款人編號 param：<asp:TextBox ID="param" runat="server"></asp:TextBox><br />
        狀態 status：<asp:TextBox ID="status" runat="server"></asp:TextBox><br />
        金額 amount：<asp:TextBox ID="amount" runat="server"></asp:TextBox><br />
        付款方式 paytype：<asp:TextBox ID="paytype" runat="server"></asp:TextBox><br />
        payformat：<asp:TextBox ID="payformat" runat="server"></asp:TextBox><br />
        authdate：<asp:TextBox ID="authdate" runat="server"></asp:TextBox><br />
        請款時間 date：<asp:TextBox ID="date" runat="server"></asp:TextBox><br />
        <asp:Button ID="submit" runat="server" Text="送出" PostBackUrl="DonateCheckPending.aspx" Height="20px" Width="60px" />
    </div>
    </form>
</body>
</html>
