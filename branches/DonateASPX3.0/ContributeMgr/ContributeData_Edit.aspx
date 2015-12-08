<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ContributeData_Edit.aspx.cs" Inherits="ContributeMgr_ContributeData_Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>捐贈物品修改</title>
    <meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxGoods_DueDate",   // id of the input field
                button: "imgGoods_DueDate"     // 與觸發動作的物件ID相同
            });
        }
        function ReadOnly() {
            var tbxGoods_Name = document.getElementById('tbxGoods_Name');
            var tbxGoods_Unit = document.getElementById('tbxGoods_Unit');
            tbxGoods_Name.readOnly = true;
            tbxGoods_Name.style.backgroundColor = '#FFE1AF';
            tbxGoods_Unit.readOnly = true;
            tbxGoods_Unit.style.backgroundColor = '#FFE1AF';
        }
        function CheckFieldMustFillBasic() {
            var cnt = 0;
            var sName = '';
            var tbxGoods_Name = document.getElementById('tbxGoods_Name');
            var tbxGoods_Qty = document.getElementById('tbxGoods_Qty');
            var tbxGoods_Unit = document.getElementById('tbxGoods_Unit');
            var tbxGoods_Amt = document.getElementById('tbxGoods_Amt');
            var tbxGoods_DueDate = document.getElementById('tbxGoods_DueDate');
            var tbxGoods_Comment = document.getElementById('tbxGoods_Comment');

            if (tbxGoods_Name.value == "") {
                alert('物品名稱 欄位不可為空白');
                tbxGoods_Name.focus();
                return false;
            }
            sName = tbxGoods_Name.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 50) {
                alert('物品名稱 欄位長度超過限制！');
                return false;
            }
            if (tbxGoods_Qty.value == "") {
                alert('數量 欄位不可為空白');
                tbxGoods_Qty.focus();
                return false;
            }
            if (isNaN(Number(tbxGoods_Qty.value)) == true) {
                alert('數量 欄位必須為數字！');
                tbxGoods_Qty.focus();
                return false;
            }
            cnt = 0;
            sName = tbxGoods_Qty.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 7) {
                alert('數量 欄位長度超過限制！');
                return false;
            }
            if (tbxGoods_Unit.value == '') {
                alert('單位 欄位不可為空白！');
                tbxGoods_Unit.focus();
                return false;
            }
            cnt = 0;
            sName = tbxGoods_Unit.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('單位 欄位長度超過限制！');
                return false;
            }
            if (tbxGoods_Amt.value == '') {
                tbxGoods_Amt.value = '0';
            }
            if (isNaN(Number(tbxGoods_Amt.value)) == true) {
                alert('折合現金 欄位必須為數字！');
                tbxGoods_Amt.focus();
                return false;
            }
            cnt = 0;
            sName = tbxGoods_Amt.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 7) {
                alert('折合現金 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxGoods_Comment.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 100) {
                alert('備註 欄位長度超過限制！');
                return false;
            }
            if (confirm('您是否確定要修改？')==false) {
                return false;
            }
            else {
                return true;
            }
         }
    </script>
    <style type="text/css">
        .style1
        {
            color: #FF3300;
        }
    </style>
</head>
<body class="body">
    <form id="form1" runat="server">
    <asp:HiddenField ID="HFD_Ser_No" runat="server" />
    <asp:HiddenField ID="HFD_Contribute_Id" runat="server" />
    <asp:HiddenField ID="HFD_Donor_Id" runat="server" />
    <asp:HiddenField ID="HFD_Goods_Id" runat="server" />
    <asp:HiddenField ID="HFD_Goods_Qty" runat="server" />
    <asp:HiddenField ID="HFD_Goods_Amt" runat="server" />
        <div align="center">
            <center>
                <table border="0" width="100%" class="table_v">
                    <tr>
                        <td width="100%">
                            <div align="center">
                                <center>
                                    <table width="100%" border="1" cellspacing="0" cellpadding="2">
                                        <tr>
                                            <th align="right" colspan='1' height="22">
                                                物品名稱：
                                            </th>
                                            <td align="left" colspan='3'>
                                                <asp:TextBox ID="tbxGoods_Name" Width="300" runat="server" MaxLength="30"></asp:TextBox>
                                                <br />
                                                <asp:Label ID="lblWarm" runat="server" Text ="庫存物品不可更改物品名稱 / 單位 " 
                                                    style="color: #FF3300"/>
                                            </td>
                                        </tr>
                                        <tr>
                                            <th align="right" colspan='1' height="22">
                                                 數量：
                                            </th>
                                            <td align="left" colspan='1'>
                                                <asp:TextBox ID="tbxGoods_Qty"  CssClass="font9" runat="server"></asp:TextBox>
                                            </td>
                                            <th align="right" colspan='1' height="22">
                                                 單位：
                                            </th>
                                            <td align="left" colspan='1'>
                                                <asp:TextBox ID="tbxGoods_Unit"  CssClass="font9" runat="server"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <th align="right" colspan='1' height="22">
                                                 折合現金：
                                            </th>
                                            <td align="left" colspan='1'>
                                                <asp:TextBox ID="tbxGoods_Amt"  CssClass="font9" runat="server"></asp:TextBox>
                                            </td>
                                            <th align="right" colspan='1' height="22">
                                                 保存期限：
                                            </th>
                                            <td align="left" colspan='1'>
                                                <asp:TextBox ID="tbxGoods_DueDate" onchange="CheckDateFormat(this, '保存期限');" CssClass="font9" runat="server"></asp:TextBox>
                                                <img id="imgGoods_DueDate" alt="" src="../images/date.gif"/>
                                            </td>
                                        </tr>
                                        <tr>
                                            <th align="right" colspan='1' height="22">
                                                備註：
                                            </th>
                                            <td align="left" colspan='3'>
                                                <asp:TextBox ID="tbxGoods_Comment" Width="300" runat="server" MaxLength="30"></asp:TextBox>
                                            </td>
                                        </tr>
                                    </table>
                                    <div class="function">
                                        <asp:Button ID="btnEdit" runat="server" Text="修改" OnClick="btnEdit_Click" 
                                             CssClass="npoButton npoButton_Modify"  OnClientClick= "return CheckFieldMustFillBasic(); "/>
                                        <asp:Button ID="btnCancle" runat="server" Text="離開" OnClientClick="window.close();"
                                             CssClass="npoButton npoButton_Exit"/>
                                    </div>
                                </center>
                            </div>
                        </td>
                    </tr>
                </table>
            </center>
        </div>
    </form>
</body>
</html>
