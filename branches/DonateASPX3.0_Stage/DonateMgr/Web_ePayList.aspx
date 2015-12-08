<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Web_ePayList.aspx.cs" Inherits="DonateMgr_Web_ePayList" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
    <title>>線上單筆奉獻查詢</title>
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
        }
        function Print(PrintType) {
            window.open('Web_ePayList_Print.aspx', 'Web_ePayList_Print', 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
        }
        function ErrorPrint(PrintType) {
            window.open('ErrorCode.aspx', 'ErrorCode', 'width=1200,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
        }
    </script>
</head>
<body class="body">
    <form id="Form1" runat="server">
        <div id="menucontrol">
            <a href="#">
                <img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
            <a href="#">
                <img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        </div>
        <div id="container">
        <asp:HiddenField runat="server" ID="HFD_Uid" />
            <h1 style="padding-bottom: 0px;">
                <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
                線上單筆奉獻查詢 
            </h1>
            <table width="100%" border="0" align="left" cellpadding="0" cellspacing="1" class="table_v">
                <tr>
                    <th align="right" colspan="1">捐款狀態：
                    </th>
                    <td align="left" colspan="1">
                        <asp:DropDownList ID="Status" runat="server" class="font9">
                            <asp:ListItem Value="">請選擇</asp:ListItem>
                            <asp:ListItem Value="0">捐款成功(授權成功)</asp:ListItem>
                            <asp:ListItem Value="1">繳款中(授權成功)</asp:ListItem>
                            <asp:ListItem Value="2">授權失敗</asp:ListItem>
                            <asp:ListItem Value="3">未完成刷卡流程</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                    <th align="right" colspan="1">付款方式：
                    </th>
                    <td><asp:dropdownlist runat="server" ID="ddlDonateIeType" CssClass="font9"></asp:dropdownlist></td>
                    <th align="right" colspan="1">捐款日期：
                    </th>
                    <td align="left" colspan="2">
                        <asp:TextBox ID="txtDonateDateS" runat="server"
                            onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                        <img id="imgDonateDateS" alt="" src="../images/date.gif" />
                        ~
                <asp:TextBox ID="txtDonateDateE" runat="server"
                    onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                        <img id="imgDonateDateE" alt="" src="../images/date.gif" />
                        　　　　
                        <asp:checkbox runat="server" ID="cbxHistoryRecord" Text="歷史記錄" CssClass="font9"></asp:checkbox>
                    </td>
                    <td align="right" colspan="3" class="style1">
                        <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server" Width="80px"
                            Text="查詢" OnClick="btnQuery_Click" />
                    <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="80px" Text="列表"
                 OnClientClick="if (confirm('您是否確定要將查詢結果列表？')==false) {return false;} Print('');"/>
                <asp:Button ID="btnToxls" runat="server"  Width="20mm"
                    Text="匯出" OnClick="btnToxls_Click" CssClass="npoButton npoButton_Excel" OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');"/>
                 
                </td>
                </tr>
                <tr>
                    <td colspan="7">
                        <div style="color: blue">1.點擊任一奉獻紀錄可以進入捐款人之捐款紀錄清單。2.點擊收據編號可以直接進入其收據維護紀錄頁。3.預設顯示近三個月記錄，歷史記錄請勾選後再按查詢。</div>
                    </td>
                    <td align="right"><asp:Button ID="btnErrorQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="120px" Text="錯誤代碼查詢" 
                     OnClientClick="ErrorPrint('');"/></td>
                </tr>
                <tr>
                    <td align="left" colspan="4">
                        <asp:Label ID="lblQueryCnt" runat="server"></asp:Label></td>
                    <td align="left" colspan="4">
                        <asp:Label ID="lblCheckCnt" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td align="center" width="100%" colspan="8">
                        <asp:Label ID="lblGridList" runat="server"></asp:Label>
                    </td>
                </tr>
            </table>

        </div>
    </form>
</body>
</html>
