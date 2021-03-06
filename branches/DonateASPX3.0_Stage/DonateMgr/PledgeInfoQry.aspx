﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PledgeInfoQry.aspx.cs" Inherits="DonateMgr_PledgeQry" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>固定轉帳授權書輸入 </title>
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
        固定轉帳授權書輸入 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                捐款人：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
            </td>
             <th align="right" colspan="1">
                捐款人編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Id" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                會員編號：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxMemberNo" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1" >
                身份別：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonor_Type" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right"colspan="1" >
                身分證/統編：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIDNo" CssClass="font9"></asp:TextBox>
            </td>
             <th align="right" colspan="1">
                聯絡電話：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCellular_Phone" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                地址：
            </th>
            <td colspan="3" >
                <asp:dropdownlist runat="server" ID="ddlCity" CssClass="font9" 
                     AutoPostBack="True" onselectedindexchanged="ddlCity_SelectedIndexChanged"></asp:dropdownlist>
                <asp:dropdownlist runat="server" ID="ddlArea" CssClass="font9" 
                     AutoPostBack="True"></asp:dropdownlist>
                <asp:TextBox runat="server" ID="tbxAddress" CssClass="font9" Width="200px"></asp:TextBox>
            </td>
        </tr>    
        <tr>
            <td align="right"  colspan="8">
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" OnClick="btnQuery_Click"/>
                <asp:Button ID="btnExit" CssClass="npoButton npoButton_Cancel" runat="server"  Width="20mm"
                    Text="取消" onclick="btnExit_Click"/>
                <asp:Button ID="btnDonor_Add" CssClass="npoButton npoButton_New" runat="server"  Width="40mm"
                    Text="新增捐款人資料" OnClick="btnDonor_Add_Click" align="right"/>
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="8" align="center">
                 <asp:Label ID="lblGridList" runat="server"></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
