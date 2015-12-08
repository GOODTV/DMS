<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ContributeList.aspx.cs" Inherits="ContributeMgr_ContributeList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>實物奉獻捐贈維護</title>
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
                inputField: "txtContribute_DateS",   // id of the input field
                button: "imgContribute_DateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtContribute_DateE",   // id of the input field
                button: "imgContribute_DateE"     // 與觸發動作的物件ID相同
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
            if (window.confirm('您是否確定要將查詢結果匯出？') == false) {
                return false; 
            }
            else {
                window.open('../ContributeMgr/ContributeList_Print.aspx', 'ContributeList_Print', 'width=1000,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
            }
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
        實物奉獻捐贈維護 
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
                捐贈人：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right"colspan="1" >
                捐贈人編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Id" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1" >
                捐贈方式：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlContribute_Payment" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐贈用途：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlContribute_Purpose" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                經手人：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlCreate_User" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                收據開立：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlInvoice_Type" CssClass="font9"></asp:dropdownlist>
            </td> 
            <td align="center" colspan="2">
            <asp:checkbox runat="server" ID="cbxInvoice_Pre" Text="手開收據" CssClass="font9" ></asp:checkbox>
            <asp:checkbox runat="server" ID="cbxExport" Text="作廢收據 " CssClass="font9" ></asp:checkbox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐贈日期：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="txtContribute_DateS" onchange="CheckDateFormat(this, '捐贈日期');" runat="server" Width="90px"></asp:TextBox>
                <img id="imgContribute_DateS" alt="" src="../images/date.gif" /> ~
                <asp:TextBox ID="txtContribute_DateE" onchange="CheckDateFormat(this, '捐贈日期');" runat="server" Width="90px"></asp:TextBox>
                <img id="imgContribute_DateE" alt="" src="../images/date.gif" />
            </td>
            
            <th align="right" colspan="1">
                收據編號：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxInvoice_NoS" CssClass="font9"></asp:TextBox>~
                <asp:TextBox runat="server" ID="tbxInvoice_NoE" CssClass="font9"></asp:TextBox>
            </td> 
            
        </tr>
        <tr> 
            <th align="right" colspan="1">
                沖帳日期：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxAccoun_DateS" onchange="CheckDateFormat(this, '沖帳日期');" runat="server" Width="90px"></asp:TextBox>
                <img id="imgAccoun_DateS" alt="" src="../images/date.gif" /> ~
                <asp:TextBox ID="tbxAccoun_DateE" onchange="CheckDateFormat(this, '沖帳日期');" runat="server" Width="90px"></asp:TextBox>
                <img id="imgAccoun_DateE" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                會計科目：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlAccounting_Title" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlAct_Id" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <td align="left" colspan="2" class="style1">
                折合現金合計：
                <asp:Label ID="lblAmt" runat="server" Text=""></asp:Label>
                元</td>
            <td align="right" colspan="6" class="style1">
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" OnClick="btnQuery_Click"/>
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  
                    Width="20mm" Text="列表" OnClientClick="Print('');"/>
                <asp:Button ID="btnToxls" runat="server"  Width="20mm"
                    Text="匯出" OnClick="btnToxls_Click" CssClass="npoButton npoButton_Excel" OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');"/>
                <asp:Button ID="btnAdd" CssClass="npoButton npoButton_New" runat="server"  Width="25mm"
                    Text="新增資料" onclick="btnAdd_Click"/>
            </td>
        </tr>
        <tr>
            <td  align="center" width="100%" colspan="8">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
