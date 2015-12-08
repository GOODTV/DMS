<%@ Page Language="C#" AutoEventWireup="true" CodeFile="MemberNameList.aspx.cs" Inherits="MemberMgr_MemberNameList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>讀者名冊</title>
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
        function Print(PrintType) {
            window.open('../MemberMgr/MemberNameList_Print.aspx', 'MemberNameList_Print', 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
        }
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
        讀者名冊</h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                身份別：
            </th>
            <td align="left" colspan="7">
                <asp:CheckBoxList ID="cblDonor_Type" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                狀態：
            </th>
            <td align="left" colspan="7">
                <asp:dropdownlist runat="server" ID="ddlMember_Status" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                通訊地址：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlCity" CssClass="font9" 
                     AutoPostBack="True" onselectedindexchanged="ddlCity_SelectedIndexChanged"></asp:dropdownlist>
                <asp:dropdownlist runat="server" ID="ddlArea" CssClass="font9"></asp:dropdownlist>
             </td>         
        </tr>
        <tr>
            <th align="right" colspan="1">
                讀者姓名：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
               編號： 
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxMember_NoS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="tbxMember_NoE" CssClass="font9" Width="70px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1" >
                生日月份：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlBirthMonth" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                文宣品：
            </th>
            <td align="left" colspan="3">
                <asp:CheckBoxList ID="cblPropaganda" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                其他：
            </th>
            <td align="left" colspan="1" >
                <asp:checkbox runat="server" ID="cbxIsAbroad" Text="海外地址" CssClass="font9"></asp:checkbox>
            </td>
        </tr>
        <tr>
            <td align="right" colspan="4" class="style1">
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" onclick="btnQuery_Click"/>
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="列表" OnClientClick="if (confirm('您是否確定要將查詢結果匯出？')==false) {return false;} Print('');"/>
                <asp:Button ID="btnToxls" CssClass="npoButton npoButton_Excel" runat="server"  Width="20mm"
                    Text="匯出" OnClientClick="return confirm('您是否確定要將查詢結果匯出？')" onclick="btnToxls_Click"/>
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="4">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>