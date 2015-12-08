<%@ Page Language="C#" AutoEventWireup="true" CodeFile="MemberQry.aspx.cs" Inherits="Member_MemberQry" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>讀者資料維護</title>
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
        function Print() {
            if (confirm('您是否確定要將查詢結果匯出？') == false) {
                return false;
            }
            else {
                window.open('../MemberMgr/MemberQry_Print.aspx', 'MemberQry_Print', 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
            }
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
        讀者資料維護 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1">
                讀者姓名：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                讀者編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxMember_No" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right"colspan="1" >
                聯絡電話：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxTel_Office" CssClass="font9"></asp:TextBox>
            </td>
            
        </tr>
        <tr>
            <th align="right" colspan="1" >
                身份別：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonor_Type" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                文宣品：
            </th>
            <td align="left" colspan="1">
                <%--紙本月刊本數<asp:Textbox runat="server" ID="tbxIsSendNewsNum" CssClass="font9" Width="50px" ></asp:Textbox>
                <asp:checkbox runat="server" ID="cbxIsDVD" Text="DVD" CssClass="font9"></asp:checkbox>--%>
                <asp:checkbox runat="server" ID="cbxIsSendEpaper" Text="電子文宣" CssClass="font9"></asp:checkbox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                地址：
            </th>
            <td align="left" colspan="1" >
                <asp:checkbox runat="server" ID="cbxIsAbroad" Text="海外地址" CssClass="font9"></asp:checkbox>
            </td>
            <td colspan="4" >
                <asp:dropdownlist runat="server" ID="ddlCity" CssClass="font9" 
                     AutoPostBack="True" onselectedindexchanged="ddlCity_SelectedIndexChanged"></asp:dropdownlist>
                <asp:dropdownlist runat="server" ID="ddlArea" CssClass="font9" 
                     AutoPostBack="True"></asp:dropdownlist>
                <asp:TextBox runat="server" ID="tbxAddress" CssClass="font9" Width="250px"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <td align="right" colspan="6" class="style1">
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" OnClick="btnQuery_Click"/>
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="列表" OnClientClick="Print();"/>
                <asp:Button ID="btnToxls" runat="server"  Width="20mm"
                    Text="匯出" OnClick="btnToxls_Click" CssClass="npoButton npoButton_Excel" OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');"/>
                <asp:Button ID="btnAdd" CssClass="npoButton npoButton_New" runat="server"  Width="20mm"
                    Text="新增" OnClick="btnAdd_Click" />
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="6" align="center">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
