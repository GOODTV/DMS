﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="GroupClass_Edit.aspx.cs"
    Inherits="DonorGroupMgr_GroupClass_Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
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
            if ($('#txtGroupClassName').getValue() == '')
            {
                alert('群組類別為必填欄位!');
                return false;
            }
            return true;
        }
        //---------------------------------------------------------------------------
    </script>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <asp:HiddenField runat="server" ID="HFD_Mode" />
    <asp:HiddenField runat="server" ID="HFD_GroupClassUID" />
    <h1>
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        <asp:Literal ID="litTitle" runat="server"></asp:Literal>
    </h1>
    <table class="table_v" width="100%">
        <tr>
            <th width="15%" align="right">
                群組類別：
            </th>
            <td>
                <asp:TextBox runat="server" ID="txtGroupClassName" Width="300px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right">
                備註：
            </th>
            <td>
                <asp:TextBox runat="server" ID="txtSupplement" Width="80mm"></asp:TextBox>
            </td>
        </tr>
    </table>
    <div class="function">
        <asp:Button ID="btnAdd" class="npoButton npoButton_New" runat="server" 
            Text="新增" onclick="btnAdd_Click" OnClientClick="return CheckData()" />
        <asp:Button ID="btnUpdate" class="npoButton npoButton_Modify" runat="server" 
            Text="修改" onclick="btnUpdate_Click" OnClientClick="return CheckData()"/>
        <asp:Button ID="btnDelete" class="npoButton npoButton_Del" runat="server" 
            Text="刪除" onclick="btnDelete_Click" OnClientClick="return window.confirm ('您是否確定要刪除該群組類別？\n若刪除該群組類別，連同關連群組項目和捐款人設定將會一併刪除！');"/>
        <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
            Text="離開" onclick="btnExit_Click"/>
    </div>
    </form>
</body>
</html>
