<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ContributeIssueList.aspx.cs" Inherits="ContributeMgr_ContributeIssueList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>實物奉獻領用作業</title>
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
                inputField: "txtIssue_DateS",   // id of the input field
                button: "imgIssue_DateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtIssue_DateE",   // id of the input field
                button: "imgIssue_DateE"     // 與觸發動作的物件ID相同
            });
        }
        function Print(PrintType) {
            if (window.confirm('您是否確定要將查詢結果匯出？') == false) {
                return false;
            }
            else {
                window.open('../ContributeMgr/ContributeIssueList_Print.aspx', 'ContributeIssueList_Print', 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
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
        實物奉獻領用作業 
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
                領取人：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIssue_Processor" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                領取用途：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlIssue_Purpose" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                經手人：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlCreate_User" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            
             
            <th align="right" colspan="1">
                領取日期：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="txtIssue_DateS" onchange="CheckDateFormat(this, '領取日期');" runat="server" Width="90px"></asp:TextBox>
                <img id="imgIssue_DateS" alt="" src="../images/date.gif" /> ~
                <asp:TextBox ID="txtIssue_DateE" onchange="CheckDateFormat(this, '領取日期');" runat="server" Width="90px"></asp:TextBox>
                <img id="imgIssue_DateE" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                領取編號：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxIssue_NoS" CssClass="font9"></asp:TextBox>~
                <asp:TextBox runat="server" ID="tbxIssue_NoE" CssClass="font9"></asp:TextBox>
            </td> 
            <td align="center" colspan="2">
            <asp:checkbox runat="server" ID="cbxIssue_Pre" Text="手開" CssClass="font9" ></asp:checkbox>
            <asp:checkbox runat="server" ID="cbxExport" Text="作廢 " CssClass="font9" ></asp:checkbox>
            </td> 
        </tr>
        <tr>
            <td align="right" colspan="8" class="style1">
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
