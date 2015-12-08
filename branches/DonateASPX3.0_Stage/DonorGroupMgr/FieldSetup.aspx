<%@ Page Language="C#" AutoEventWireup="true" CodeFile="FieldSetup.aspx.cs" Inherits="DonorGroupMgr_FieldSetup" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>欄位設定</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            InitMenu();
        });
    </script>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <asp:HiddenField runat="server" ID="HFD_Mode" />
    <h1>
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        欄位設定
    </h1>
    <table width="100%" class="table_v">
        <tr>
            <td>
                顯示欄位：
            </td>
            <td>
            </td>
            <td>
                所有欄位：
            </td>
        </tr>
        <tr>
            <td width="25%"  style="vertical-align:top" >
                <asp:ListBox ID="lstDisplayField" runat="server" Height="300px" Width="180px" 
                    SelectionMode="Multiple"></asp:ListBox>
            </td>
            <td  width="60px" align="center" style="vertical-align:top" >
                <br/><br/>
                <asp:Button ID="btnSelect" Text="加入 <<" runat="server" OnClick="btnSelect_Click" />
                <br/><br/>
                <asp:Button ID="btnRemove" Text="移出 >>" runat="server" OnClick="btnRemove_Click" />
                <br/><br/><asp:Button ID="btnUp" Text="上移" runat="server" 
                    onclick="btnUp_Click" />
                <br/><br/><asp:Button ID="Button1" Text="下移" runat="server" 
                    onclick="Button1_Click" />
                <br/><br/><br/><br/><br/><br/><br/><br/><br/>
            </td>
            <td align="left" style="vertical-align:top">
                <asp:ListBox ID="lstAllField" runat="server" Height="300px" Width="180px" SelectionMode="Multiple">
                </asp:ListBox>
            </td>
        </tr>
        <tr>
            <td colspan="3">
                命名：
                <asp:DropDownList ID="ddlConfigName" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlConfigName_SelectedIndexChanged"></asp:DropDownList>
                &nbsp;&nbsp;<asp:Button ID="btnSave" Text="儲存" runat="server" onclick="btnSave_Click" />
                &nbsp;&nbsp;<asp:Button ID="btnDelete" Text="刪除" runat="server" onclick="btnDelete_Click" OnClientClick ="return window.confirm ('您是否確定要刪除選擇的命名 ?');"/>
                &nbsp;&nbsp;<asp:Button ID="btnExit" runat="server" Text="離開" OnClick="btnExit_Click" />
                <br/><br/>新名稱：
                <asp:TextBox ID="txtNewConfigName" runat="server" Width="186px" ></asp:TextBox>
                <asp:Button ID="btnSaveAs" Text="另存新名稱" runat="server" 
                    onclick="btnSaveAs_Click" />
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
