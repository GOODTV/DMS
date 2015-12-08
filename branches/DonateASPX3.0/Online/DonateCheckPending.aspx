<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateCheckPending.aspx.cs" Inherits="Online_DonateCheckPending" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>已請款通知</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/jquery.field.js"></script>
</head>
<body>
    <form id="Form1" runat="server">
        <asp:HiddenField ID="storeid" runat="server"  value=""/>
        <asp:HiddenField ID="orderid" runat="server" value=""/>
        <asp:HiddenField ID="amount" runat="server" value=""/>
        <asp:HiddenField ID="date" runat="server" value=""/>
    <div>
    
    </div>
    </form>
</body>
</html>
