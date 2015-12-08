<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateToPledge.aspx.cs" Inherits="Online_DonateToPledge" validateRequest="false" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="Form1" runat="server">
        <asp:HiddenField ID="HFD_DonorID" runat="server" />

        <asp:HiddenField ID="cavalue" runat="server" value=""/>
        <asp:HiddenField ID="storeid" runat="server"  value=""/>
        <asp:HiddenField ID="ordernumber" runat="server" value=""/>
        <asp:HiddenField ID="amount" runat="server" value=""/>
        <asp:HiddenField ID="authstatus" runat="server" value=""/>
        <asp:HiddenField ID="authcode" runat="server" value=""/>
        <asp:HiddenField ID="authtime" runat="server" value=""/>
        <asp:HiddenField ID="authmsg" runat="server" value=""/>
        <asp:HiddenField ID="postXML" runat="server" value=""/>
        <asp:HiddenField ID="strRqXML" runat="server" value=""/>
    </form>
</body>
</html>
