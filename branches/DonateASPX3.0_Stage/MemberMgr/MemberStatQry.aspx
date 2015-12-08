<%@ Page Language="C#" AutoEventWireup="true" CodeFile="MemberStatQry.aspx.cs" Inherits="MemberMgr_MemberStatQry" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>讀者統計</title>
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
        function Print(PrintType) {
            $.ajax({
                type: 'post',
                url: "../common/ajax.aspx",
                data: 'Type=6' + '&Category=' + $("input[id^='<% =rblCategory.ClientID %>']:checked").val(),
                async: false, //同步
                success: function (result) {
                },
                error: function () { alert('ajax failed'); }
            })
            window.open('../MemberMgr/MemberStatQry_Print.aspx', 'MemberStatQry_Print', 'width=900,height=650,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
        }
        function toXML(PrintType) {
            $.ajax({
                type: 'post',
                url: "../common/ajax.aspx",
                data: 'Type=6' + '&Category=' + $("input[id^='<% =rblCategory.ClientID %>']:checked").val(),
                async: false, //同步
                success: function (result) {
                },
                error: function () { alert('ajax failed'); }
            })
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
        讀者統計
    </h1>
    <table width="60%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <td align="right">
            統計項目： 
            </td>
            <td align="left">
                <asp:RadioButtonList ID="rblCategory" runat="server" 
                    RepeatDirection="Horizontal">
                    <%--<asp:ListItem Value="1" Selected="True">類別</asp:ListItem>--%>
                    <asp:ListItem Value="2" Selected="True">性別</asp:ListItem>
                    <asp:ListItem Value="3">年齡</asp:ListItem>
                    <asp:ListItem Value="4">教育程度</asp:ListItem>
                    <asp:ListItem Value="5">職業別</asp:ListItem>
                    <asp:ListItem Value="6">婚姻狀況</asp:ListItem>
                    <asp:ListItem Value="7">宗教信仰</asp:ListItem>
                    <asp:ListItem Value="8">通訊縣市</asp:ListItem>
                    <asp:ListItem Value="9">狀態</asp:ListItem>
                </asp:RadioButtonList>
            </td>
        </tr>
        <tr>
            <td align="center"  colspan="2">
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" OnClientClick="if (confirm('您是否確定要將查詢結果匯出？')==false) {return false;} Print('');"/>
                <asp:Button ID="btnToxls" runat="server"  Width="20mm" OnClientClick="if (confirm('您是否確定要將查詢結果匯出？')==false) {return false;} toXML('');"
                    Text="匯出" OnClick="btnToxls_Click" CssClass="npoButton npoButton_Excel"/>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
