<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateMonthQry.aspx.cs" Inherits="Report_DonateMonthReport" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>捐款報表</title>
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
                button: "imgtxtDonateDateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtDonateDateE",   // id of the input field
                button: "imgtxtDonateDateE"     // 與觸發動作的物件ID相同
            });
        }
        function Print(PrintType) {
            window.open('../Report/DonateMonthQry_Print.aspx', 'DonateMonthQry_Print', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600', '');
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
        捐款報表</h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1">
                機構：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                收據編號：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="txtInvoice_NoB" runat="server"></asp:TextBox>~<asp:TextBox ID="txtInvoice_NoE" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款日期：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="txtDonateDateS" runat="server" 
                    onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                <img id="imgtxtDonateDateS" alt="" src="../images/date.gif" /> ~
                <asp:TextBox ID="txtDonateDateE" runat="server" 
                    onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                <img id="imgtxtDonateDateE" alt="" src="../images/date.gif" />
            </td>
            <!--th align="right"colspan="1">
                單筆捐款金額：
            </!--th>
            <td align="left" colspan="3">
                <asp:TextBox ID="tbxDonate_AmtB" runat="server"></asp:TextBox>~<asp:TextBox ID="tbxDonate_AmtE" runat="server"></asp:TextBox>
            </td>
            <th align="right" colspan="1" >
                募款活動：
            </th>
            <td align="left" colspan="3">
                <asp:dropdownlist runat="server" ID="ddlActName" CssClass="font9"></asp:dropdownlist>
            </td-->
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款方式：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblDonate_Payment" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                線上付款方式：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblPayment_type" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款用途：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblDonate_Purpose" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                區域：
            </th>
            <td align="left" colspan="5">
                 <asp:RadioButtonList ID="rblAddress" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Value="1">國內</asp:ListItem>
                    <asp:ListItem Value="2">海外 </asp:ListItem>
                    <asp:ListItem Value="3" Selected="True">全部 </asp:ListItem>
                </asp:RadioButtonList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                匯入台銀回覆檔檔名：
            </th>
            <td align="left" colspan="5">
                <asp:TextBox ID="tbPledgeBatchFileName" runat="server" Width="230px" MaxLength="30"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <td align="right" colspan="6" class="style1">
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="列表"   onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
                <asp:Button ID="btnToxls" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click" 
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="6">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
