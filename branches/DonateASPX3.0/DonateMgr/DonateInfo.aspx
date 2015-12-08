<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateInfo.aspx.cs" Inherits="DonateMgr_DonateInfo" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>收據維護紀錄</title>
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
        // menu控制
        $(document).ready(function () {
            InitMenu();
        });
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "txtDonateDateS",   // id of the input field
                button: "imgDonateDateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtDonateDateE",   // id of the input field
                button: "imgDonateDateE"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxAccoun_DateS",   // id of the input field
                button: "imgAccoun_DateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxAccoun_DateE",   // id of the input field
                button: "imgAccoun_DateE"     // 與觸發動作的物件ID相同
            });
        }
        function Print(PrintType) {
            window.open('../DonateMgr/DonateInfo_Print.aspx', 'DonateInfo_Print', 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
        }
        function CheckFieldMustFillBasic() {
            var Donor_Id = document.getElementById('tbxDonor_Id');
            var Donor_Id_Old = document.getElementById('tbxDonor_Id_Old');
            var Invoice_NoS = document.getElementById('tbxInvoice_NoS');
            var Invoice_NoE = document.getElementById('tbxInvoice_NoE');
            
            if (isNaN(Number(Donor_Id.value)) == true) {
                alert('『捐款人編號』 欄位必須為數字！');
                Donor_Id.focus();
                return false;
            }
            if (isNaN(Number(Donor_Id_Old.value)) == true) {
                alert('『舊編號』 欄位必須為數字！');
                Donor_Id_Old.focus();
                return false;
            }
            if (isNaN(Number(Invoice_NoS.value)) == true || isNaN(Number(Invoice_NoE.value)) == true) {
                alert('『收據編號』 欄位必須為數字！不需輸入前置碼');
                Invoice_NoS.focus();
                return false;
            }            
            return true;
        }
    </script>
    <style type="text/css">
        .style1
        {
            height: 35px;
        }
    </style>
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
        收據維護紀錄 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                捐款人：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right"colspan="1" >
                捐款人編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Id" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right"colspan="1" >
                舊編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Id_Old" CssClass="font9"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1" >
                 捐款方式：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Payment" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                捐款用途：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Purpose" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                捐款日期：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="txtDonateDateS" runat="server" 
                    onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                <img id="imgDonateDateS" alt="" src="../images/date.gif" /> ~
                <asp:TextBox ID="txtDonateDateE" runat="server" 
                    onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                <img id="imgDonateDateE" alt="" src="../images/date.gif" />
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                收據開立：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlInvoice_Type" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                收據編號：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxInvoice_NoS" CssClass="font9"></asp:TextBox>~
                <asp:TextBox runat="server" ID="tbxInvoice_NoE" CssClass="font9"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                經手人：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlCreate_User" CssClass="font9"></asp:dropdownlist>
            </td> 
            <td align="left" colspan="2">
                <asp:checkbox runat="server" ID="cbxInvoice_Pre" Text="手開收據" CssClass="font9" ></asp:checkbox>
                <asp:checkbox runat="server" ID="cbxExport" Text="作廢收據 " CssClass="font9" ></asp:checkbox>
            </td> 
        </tr>
        <tr>
            <td>

            </td>
            <td>

            </td>
            <td>

            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                傳票號碼：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonation_NumberNo" CssClass="font9"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                劃撥 / 匯款單號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonation_SubPoenaNo" CssClass="font9"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlAct_Id" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                入帳銀行：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlAccoun_Bank" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr> 
            <th align="right" colspan="1">
                沖帳日期：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxAccoun_DateS" runat="server" 
                    onchange="CheckDateFormat(this, '沖帳日期');" Width="90px"></asp:TextBox>
                <img id="imgAccoun_DateS" alt="" src="../images/date.gif" /> ~
                <asp:TextBox ID="tbxAccoun_DateE" runat="server" 
                    onchange="CheckDateFormat(this, '沖帳日期');" Width="90px"></asp:TextBox>
                <img id="imgAccoun_DateE" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                收據抬頭：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxInvoice_Title" CssClass="font9" Width="200px"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                會計科目：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlAccounting_Title" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <td align="left" colspan="2" class="style1">
                捐款金額合計：
                <asp:Label ID="lblAmt" runat="server" Text=""></asp:Label>
                元</td>
            <td align="right" colspan="7" class="style1">
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" OnClick="btnQuery_Click" OnClientClick="return CheckFieldMustFillBasic(); "/>
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm" Text="列表"
                 OnClientClick="if (confirm('您是否確定要將查詢結果匯出？')==false) {return false;} Print('');"/>
                <asp:Button ID="btnToxls" runat="server"  Width="20mm"
                    Text="匯出" OnClick="btnToxls_Click" CssClass="npoButton npoButton_Excel" OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');"/>
                <asp:Button ID="btnAdd" CssClass="npoButton npoButton_New" runat="server"  Width="25mm"
                    Text="新增資料" onclick="btnAdd_Click"/>
            </td>
        </tr>
        <tr>
            <td  align="center" width="100%" colspan="9">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>