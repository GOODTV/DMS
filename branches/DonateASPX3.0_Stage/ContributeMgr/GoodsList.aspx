<%@ Page Language="C#" AutoEventWireup="true" CodeFile="GoodsList.aspx.cs" Inherits="ContributeMgr_GoodsList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>實物奉獻主檔維護</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            InitMenu();
        });
        function Print(PrintType) {
            window.open('../ContributeMgr/GoodsList_Print.aspx', 'GoodsList_Print', 'width=800,height=600,toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes', '');
        }
    </script>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        實物奉獻主檔維護 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1">
                物品代號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxGoods_Id" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right"colspan="1" >
                物品名稱：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxGoods_Name" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1" >
                 庫存管理：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlGoods_IsStock" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                物品性質：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlGoods_Property" CssClass="font9" 
                    AutoPostBack="True" 
                    onselectedindexchanged="ddlGoods_Property_SelectedIndexChanged"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                物品類別：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlGoods_Type" CssClass="font9"></asp:dropdownlist>
            </td> 
            <td align="left" colspan="1">
            <asp:checkbox runat="server" ID="cbxGoods_Qty" Text="現有庫存量>0 " CssClass="font9" ></asp:checkbox>
            </td> 
        </tr>
        <tr>
            <td align="right" colspan="6" class="style1">
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" OnClick="btnQuery_Click"/>
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm" Text="列表"
                 OnClientClick="if (confirm('您是否確定要將查詢結果匯出？')==false) {return false;} Print('');"/>
                <asp:Button ID="btnToxls" runat="server"  Width="20mm"
                    Text="匯出" OnClick="btnToxls_Click" CssClass="npoButton npoButton_Excel" OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');"/>
                <asp:Button ID="btnAdd" CssClass="npoButton npoButton_New" runat="server"  Width="25mm"
                    Text="新增資料" onclick="btnAdd_Click"/>
            </td>
        </tr>
        <tr>
            <td  align="center" width="100%" colspan="6">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
