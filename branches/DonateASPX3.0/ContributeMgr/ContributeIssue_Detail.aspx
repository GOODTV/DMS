<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ContributeIssue_Detail.aspx.cs" Inherits="ContributeMgr_ContributeIssue_Detail" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>實物奉獻領用作業</title>
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
    <script type="text/javascript">
        function Export_Y() {
            $('#btnEdit').hide();
            $('#btnPrint').hide();
            $('#btn_Export').hide();
        }
        function Export_N() {
            $('#btn_ReExport').hide();
        }
        function Print(PrintType) {
        var HFD_Uid = document.getElementById('HFD_Uid');
            if (window.confirm('您是否確定要將列印單據？') == false) {
                return false;
            }
            else {
                window.open('../ContributeMgr/ContributeIssue_Print.aspx?Issue_Id=' + HFD_Uid.value, 'ContributeIssue_Print', 'width=800,height=650,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
                return true;
            }
        }
    </script>
    </head>
<body class="body">
    <form id="Form2" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        實物奉獻領用作業
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDept" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                領取人：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIssue_Processor" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                領用日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxIssue_Date" runat="server" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                領用用途：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxIssue_Purpose" runat="server" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                出貨單位：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIssue_Org" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td>
            <th colspan="1" align="right">
                備註：
            </th>
            <td colspan="2">
                <asp:Textbox runat="server" ID="tbxComment" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True" Width="200px"></asp:Textbox> 
            </td>
            <td align="center" colspan="1">
                <asp:checkbox runat="server" ID="cbxIssue_Pre" Text="手開領用收據" CssClass="font9" 
                    OnClick="cbxIssue_Pre_OnClick()" Enabled="False" ></asp:checkbox>
            </td>
            <th align="right" colspan="1">
                領用編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIssue_No" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True" ></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th colspan="1" align="right">
                領用內容：
            </th>
            <td  align="center" colspan="7">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        </table>
    </div>
    <div class="function">
        <asp:Button ID="btnEdit" class="npoButton npoButton_Modify" runat="server" 
            Text="修改領用資料" Width="110px" onclick="btnEdit_Click"/>
        <asp:Button ID="btnPrint" class="npoButton npoButton_Print" runat="server" 
            Text="單據補印" OnClientClick="Print('');" onclick="btnPrint_Click"/>
        <asp:Button ID="btn_Export" class="npoButton npoButton_Export" runat="server" 
            Text="單據作廢" onclick="btnExport_Click" OnClientClick= "return confirm('您是否確定要作廢單據？');"/>
        <asp:Button ID="btn_ReExport" class="npoButton npoButton_Export" runat="server" 
            Text="還原作廢收據" Width="120px" onclick="btn_ReExport_Click" OnClientClick= "return confirm('您是否確定要還原作廢收據？');" />
        <asp:Button ID="btnExit" class="npoButton npoButton_Cancel" runat="server"    
            Text="取消" Width="80px" onclick="btnExit_Click" />
    </div>
    </form>
</body>
</html>