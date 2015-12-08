<%@ Page Language="C#" AutoEventWireup="true" CodeFile="GiftShow.aspx.cs" Inherits="DonorMgr_Gift_Show" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>公關贈品品項查詢</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        function ReturnOpener(val) {
            var url = window.location.toString();
            var str = "";
            var HFD_Gifts_Id = "";
            var tbxGifts_Name = "";
            if (url.indexOf("?") != -1) {
                var ary = url.split("?")[1].split("&");
                for (var i in ary) {
                    str = ary[i].split("=")[0];
                    if (str == "Gifts_Id") {
                        HFD_Gifts_Id = decodeURI(ary[i].split("=")[1]);
                    }
                    if (str == "Gifts_Name") {
                        tbxGifts_Name = decodeURI(ary[i].split("=")[1]);
                    }
                }
            }
            var Gifts_Id = val.split("|")[0];
            var Gifts_Name = val.split("|")[1];

            opener.document.getElementById(HFD_Gifts_Id).value = Gifts_Id;
            opener.document.getElementById(tbxGifts_Name).value = Gifts_Name;
            window.close();
        }
    </script>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <div id="container">
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
        <tr>
            <th align="right" colspan="1">
                公關贈品品項性質：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlGifts_Property" CssClass="font9" ></asp:dropdownlist>
            </td>
            <td align="right" colspan="1" >
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" OnClick="btnQuery_Click"/>
            </td>
        </tr> 
        <tr>
            <td  align="center" width="100%" colspan="5">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
