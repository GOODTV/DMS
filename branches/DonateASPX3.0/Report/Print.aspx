<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Print.aspx.cs" Inherits="Report_Print" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>單次收據列印</title>
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
    </script>
    <script type="text/javascript">
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "txtDateFrom",   // id of the input field
                ifFormat: "%Y/%m/%d",       // format of the input field
                showsTime: false,
                timeFormat: "24",
                button: "imgDateFrom"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtDateTo",   // id of the input field
                ifFormat: "%Y/%m/%d",       // format of the input field
                showsTime: false,
                timeFormat: "24",
                button: "imgDateTo"     // 與觸發動作的物件ID相同
            });
        }
    </script>
        <style type="text/css">
        .style8
        {
            width: 130px;
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
        單次收據列印</h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
            <tr>
                <th align="center" colspan="2"  >查詢條件</th>
        </tr>
             <tr>
                <th align="right" class="style8"  >日期：</th>
                   <td align="left">
                       <asp:TextBox ID="txtDateFrom" runat="server" 
                           onchange="CheckDateFormat(this, '出生日期');" Width="90px"></asp:TextBox>
                <img id="imgDateFrom" alt="" src="../images/date.gif" />~
                       <asp:TextBox ID="txtDateTo" runat="server" onchange="CheckDateFormat(this, '出生日期');" Width="90px"></asp:TextBox>
                <img id="imgDateTo" alt="" src="../images/date.gif" />&nbsp; (兩邊都要有範圍)</td>
        </tr>
             <tr>
                <th align="right" class="style8"  >通訊地址：</th>
                   <td align="left">
                       <asp:DropDownList ID="ddlCity" runat="server" AutoPostBack="True" 
                           DataTextField="CaseName" DataValueField="Caseid" onselectedindexchanged="ddlCity_SelectedIndexChanged" 
                           >
                       </asp:DropDownList>
                       <asp:DropDownList ID="ddlArea" runat="server" AutoPostBack="True" 
                           DataTextField="CaseName" DataValueField="Caseid">
                       </asp:DropDownList>
                 </td>
             </tr>
             <tr>
                 <th align="right" class="style8">
                     姓名：</th>
                 <td align="left">
                     <asp:TextBox ID="txtName" runat="server" CssClass="font9"></asp:TextBox>
                 </td>
             </tr>
     </table>
    <asp:Button ID="btnWord" runat="server" 
            onclick="btnWord_Click" Text="匯出郵件格式" Width="130px" 
            CssClass="npoButton npoButton_Print" />
    </div>
    <asp:Literal ID="GridList" runat="server" Visible="false"></asp:Literal>
    </form>
</body>
</html>
