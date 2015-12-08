<%@ Page Language="C#" AutoEventWireup="true" CodeFile="GoodsShow.aspx.cs" Inherits="ContributeMgr_Goods_Show" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>物品查詢</title>
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
            var HFD_Goods_Id = "";
            var tbxGoods_Name = "";
            var tbxGoods_Unit = "";
            if (url.indexOf("?") != -1) {
                var ary = url.split("?")[1].split("&");
                for (var i in ary) {
                    str = ary[i].split("=")[0];
                    if (str == "Goods_Id") {
                        HFD_Goods_Id = decodeURI(ary[i].split("=")[1]);
                    }
                    if (str == "Goods_Name") {
                        tbxGoods_Name = decodeURI(ary[i].split("=")[1]);
                    }
                    if (str == "Goods_Unit") {
                        tbxGoods_Unit = decodeURI(ary[i].split("=")[1]);
                    }
                }
            }
            var Goods_Id = val.split("|")[0];
            var Goods_Name = val.split("|")[1];
            var Goods_Unit = val.split("|")[2];

            opener.document.getElementById(HFD_Goods_Id).value = Goods_Id;
            opener.document.getElementById(tbxGoods_Name).value = Goods_Name;
            opener.document.getElementById(tbxGoods_Unit).value = Goods_Unit;
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
                物品性質：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlGoods_Property" CssClass="font9" 
                     AutoPostBack="True" 
                    onselectedindexchanged="ddlGoods_Property_SelectedIndexChanged"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                物品類別：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlGoods_Type" CssClass="font9" 
                     AutoPostBack="True" ></asp:dropdownlist>
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
