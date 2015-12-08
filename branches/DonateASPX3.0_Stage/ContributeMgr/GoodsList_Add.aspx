<%@ Page Language="C#" AutoEventWireup="true" CodeFile="GoodsList_Add.aspx.cs" Inherits="ContributeMgr_GoodsList_Add" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>實物奉獻主檔維護【新增】</title>
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
    </script>
    <script type="text/javascript">
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var cnt = 0;
            var sName = "";
            var tbxGoods_Id = document.getElementById('tbxGoods_Id');
            var tbxGoods_Name = document.getElementById('tbxGoods_Name');
            var tbxGoods_Unit = document.getElementById('tbxGoods_Unit');
            var ddlGoods_Property = document.getElementById('ddlGoods_Property');
            var ddlGoods_Type = document.getElementById('ddlGoods_Type'); 
            if (tbxGoods_Id.value == "") {
                strRet += "物品代號 ";
            }
            if (tbxGoods_Name.value == "") {
                strRet += "物品名稱 ";
            }
            if (tbxGoods_Unit.value == "") {
                strRet += "庫存單位 ";
            }
            if (ddlGoods_Property.value == "") {
                strRet += "物品性質 ";
            }
            if (ddlGoods_Type.value == "") {
                strRet += "物品類別 ";
            }
            if (strRet!="") {
                strRet += "欄位不可為空白！";
                alert(strRet)
                return false;
            }
            cnt = 0;
            sName = tbxGoods_Id.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('物品代號 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxGoods_Name.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 50) {
                alert('物品名稱 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = ddlGoods_Property.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('物品性質 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = ddlGoods_Type.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('物品類別 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxGoods_Unit.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('物品單位 欄位長度超過限制！');
                return false;
            }

            return true;
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
        實物奉獻主檔維護【新增】
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                物品代號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxGoods_Id" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                物品名稱：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxGoods_Name" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                物品性質：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlGoods_Property" CssClass="font9" 
                    AutoPostBack="True" 
                    onselectedindexchanged="ddlGoods_Property_SelectedIndexChanged"></asp:dropdownlist>
                
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                物品類別：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlGoods_Type" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                庫存單位：
            </th>
            <td align="left" colspan="5">
                <asp:TextBox ID="tbxGoods_Unit" runat="server"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                是否庫存管理：
            </th>
            <td align="left" colspan="5">
                <asp:RadioButtonList ID="rblGoods_IsStock" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Selected="True" Value="Y">是</asp:ListItem>
                    <asp:ListItem Value="N">否</asp:ListItem>
                </asp:RadioButtonList>
            </td> 
        </tr>
        <tr>
            <th colspan="1" align="right">
                備註：<br />
            </th>
            <td colspan="5">
                <asp:Textbox runat="server" ID="tbxRemark" CssClass="font9" 
                    TextMode="MultiLine" Width="450px" Height="100px" Rows="4"></asp:Textbox> 
            </td>
         </tr>
    </table>
    </div>
    <div class="function">
        <asp:Button ID="btnAdd" class="npoButton npoButton_New" runat="server" OnClientClick= "return CheckFieldMustFillBasic(); "
            Text="存檔" onclick="btnAdd_Click"/>
        <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
            Text="取消" onclick="btnExit_Click"/>
    </div>
    </form>
</body>
</html>
