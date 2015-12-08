<%@ Page Language="C#" AutoEventWireup="true" CodeFile="MainDefault.aspx.cs" Inherits="MainDefault" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
    <title></title>
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
</head>
<body class="body">
    <form id="form1" name="form" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        公告 
    </h1>
    <table width="100%" class="table_v">
        <tr>
            <td align="left">
            </td>
        </tr>
        <tr>
            <td width="100%" valign="top">
                <asp:Label ID="lblNews" runat="server"></asp:Label>
                <a href="../filemgr/News_Show_List.aspx">
                    <img border="0" src="../images/more.gif" alt="" width="63" height="21" align="right" />
                </a>
            </td>
        </tr>
        <tr>
            <td width="100%" align="left"   valign="top">
                <asp:Panel runat="server" ID="Panel1">
                    <fieldset id="Fieldset6" runat="server" style="border-color: Orange">
                        <legend id="Legend6" runat="server" class="necessary">寒士待處理案件列表</legend>
                        <asp:Label ID="lblStreetList" runat="server" Text=""></asp:Label>
                    </fieldset>
                </asp:Panel>
            </td>
        </tr>
        <tr>
            <td width="100%" align="left"   valign="top">
                <asp:Panel runat="server" ID="Panel2">
                    <fieldset id="Fieldset1" runat="server" style="border-color: Orange">
                        <legend id="Legend1" runat="server" class="necessary">單媽待處理案件列表</legend>
                        <asp:Label ID="lblWomans" runat="server" Text=""></asp:Label>
                    </fieldset>
                </asp:Panel>
            </td>
        </tr>
        <tr>
            <td width="100%" align="left"   valign="top">
                <asp:Panel runat="server" ID="Panel3">
                    <fieldset id="Fieldset2" runat="server" style="border-color: Orange">
                        <legend id="Legend2" runat="server" class="necessary">特殊事件處理待處理案件列表</legend>
                        <asp:Label ID="lblSpecialEvent" runat="server" Text=""></asp:Label>
                    </fieldset>
                </asp:Panel>
            </td>
        </tr>
        <tr>
            <td width="100%" align="left"   valign="top">
                <asp:Panel runat="server" ID="Panel4">
                    <fieldset id="Fieldset3" runat="server" style="border-color: Orange">
                        <legend id="Legend3" runat="server" class="necessary">夜宿申請表待處理案件列表</legend>
                        <asp:Label ID="lblNightSleep" runat="server" Text=""></asp:Label>
                    </fieldset>
                </asp:Panel>
            </td>
        </tr>
        <tr>
            <td width="100%" align="left"   valign="top">
                <asp:Panel runat="server" ID="Panel5">
                    <fieldset id="Fieldset4" runat="server" style="border-color: Orange">
                        <legend id="Legend4" runat="server" class="necessary">急難救助金暨醫療費申請表待處理案件列表</legend>
                        <asp:Label ID="lblMoneyApply" runat="server" Text=""></asp:Label>
                    </fieldset>
                </asp:Panel>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
