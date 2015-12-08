﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Log.aspx.cs" Inherits="SysMgr_Log" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>系統LOG查詢</title>
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
                inputField: "txtDateS",   // id of the input field
                button: "imgDateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtDateE",   // id of the input field
                button: "imgDateE"     // 與觸發動作的物件ID相同
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
        系統LOG查詢 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1">
                Log日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="txtDateS" onchange="CheckDateFormat(this, 'Log日期');" runat="server" ></asp:TextBox>
                <img id="imgDateS" alt="" src="../images/date.gif" /> ~
                <asp:TextBox ID="txtDateE" onchange="CheckDateFormat(this, 'Log日期');" runat="server" ></asp:TextBox>
                <img id="imgDateE" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                Log類別：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlLog_Type" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                帳號：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlUserID" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <td align="right" colspan="4" class="style1">
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" OnClick="btnQuery_Click"/>
            </td>
        </tr>
        <tr>
            <td  align="center" width="100%" colspan="4">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
