<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateStatQry.aspx.cs" Inherits="Report_DonateStatQry" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>捐款人統計</title>
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
                ifFormat: "%Y/%m/%d",       // format of the input field
                showsTime: false,
                timeFormat: "24",
                button: "imgtxtDonateDateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtDonateDateE",   // id of the input field
                ifFormat: "%Y/%m/%d",       // format of the input field
                showsTime: false,
                timeFormat: "24",
                button: "imgtxtDonateDateE"     // 與觸發動作的物件ID相同
            });
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
        統計分析
    </h1>
    <table width="70%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1" >
                機構：
            </th>
            <td align="left" colspan="3" >
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1" >
                捐款日期：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox ID="txtDonateDateS" runat="server" 
                    onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                <img id="imgtxtDonateDateS" alt="" src="../images/date.gif" /> ~
                <asp:TextBox ID="txtDonateDateE" runat="server" 
                    onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                <img id="imgtxtDonateDateE" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1" >
                募款活動：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlActName" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1" >
            統計項目： 
            </th>
            <td align="left" colspan="3" >
                <asp:RadioButton ID="rbCategory1" runat="server" Text="類別" GroupName="1" />
                <asp:RadioButton ID="rbCategory2" runat="server" Text="性別" GroupName="1" />
                <asp:RadioButton ID="rbCategory3" runat="server" Text="年齡" GroupName="1" />
                <asp:RadioButton ID="rbCategory4" runat="server" Text="教育程度" GroupName="1" />
                <asp:RadioButton ID="rbCategory5" runat="server" Text="職業別" GroupName="1" />
                <asp:RadioButton ID="rbCategory6" runat="server" Text="婚姻狀況" GroupName="1" />
                <asp:RadioButton ID="rbCategory7" runat="server" Text="宗教信仰" GroupName="1" />
                <br />
                <asp:RadioButton ID="rbCategory8" runat="server" Text="通訊縣市" GroupName="1" />
                <asp:RadioButton ID="rbCategory9" runat="server" Text="收據縣市" GroupName="1" />
                <asp:RadioButton ID="rbCategory10" runat="server" Text="捐款方式" GroupName="1" />
                <asp:RadioButton ID="rbCategory11" runat="server" Text="捐款用途" GroupName="1" />
                <asp:RadioButton ID="rbCategory12" runat="server" Text="捐款金額" GroupName="1" />
                <asp:TextBox ID="tbxAmt" runat="server" Width="270px"  maxlength="100" value="1000,3000,5000,10000,30000,50000,100000"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td align="center"  colspan="4">
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" OnClientClick="return confirm('您是否確定要將查詢結果匯出？');" 
                    onclick="btnPrint_Click"/>
                <asp:Button ID="btnToxls" runat="server"  Width="20mm" OnClientClick="return confirm('您是否確定要將查詢結果匯出？');"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"/>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
