﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="GroupMapping.aspx.cs" Inherits="DonorGroupMgr_GroupMapping" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.field.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            InitMenu();
        });
    </script>
    <script type="text/javascript">
        function CheckData()
        {
            if ($('#ddlGroupItem').getValue() == '')
            {
                alert('請選擇群組類別及代表！');
                return false;
            }
            return true;
        }
        //---------------------------------------------------------------------------
    </script>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <asp:HiddenField runat="server" ID="HFD_Mode" />
    <div id="menucontrol">
        <a href="#">
            <img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}"
                onclick="SwitchMenu();return false;" /></a> <a href="#">
                    <img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}"
                        onclick="SwitchMenu();return false;" /></a>
    </div>
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <h1>
                <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
                捐款人-群組選取功能設定
            </h1>
            <table width="100%" class="table_v">
                <tr>
                    <th align="right" rowspan="2">
                        奉獻者姓名
                    </th>
                    <td width="25%" rowspan="2">
                        <asp:TextBox runat="server" ID="txtDonorName"></asp:TextBox>
                        <asp:Button ID="btnSearch" Text="搜尋" runat="server" OnClick="btnSearch_Click" />
                        <br />
                        <asp:ListBox ID="lstAvailableUser" runat="server" Height="300px" Width="180px" SelectionMode="Multiple">
                        </asp:ListBox>
                        <td align="left" rowspan="2">
                            <asp:Button ID="btnSelect" Text="加入" runat="server" OnClick="btnSelect_Click" OnClientClick="return CheckData()"/>
                        </td>
                    </td>
                    <th width="20%" align="right">
                        群組類別及群組代表：
                    </th>
                    <td width="25%">
                        <asp:DropDownList ID="ddlGroupClass" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlGroupClass_SelectedIndexChanged">
                        </asp:DropDownList>
                        <asp:DropDownList ID="ddlGroupItem" runat="server" AutoPostBack="True" 
                            onselectedindexchanged="ddlGroupItem_SelectedIndexChanged" >
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        群組代表所屬奉獻者
                    </th>
                    <td align="left">
                        <asp:ListBox ID="lstGroupItemUser" runat="server" Height="300px" Width="180px" SelectionMode="Multiple">
                        </asp:ListBox>
                    </td>
                    <td align="left" rowspan="2">
                        <asp:Button ID="btnRemove" Text="移出" runat="server" OnClick="btnRemove_Click" />
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    </form>
</body>
</html>
