﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Invoice_Print_Data.aspx.cs" Inherits="Invoice_Print_Data" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>收據列印</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            //InitMenu();
        });
    </script>
    <script type="text/javascript">
    </script>
        <style type="text/css">
        .style8
        {
            width: 130px;
        }
    </style>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <asp:Literal ID="GridList" runat="server" Visible="false"></asp:Literal>
    </form>
</body>
</html>
