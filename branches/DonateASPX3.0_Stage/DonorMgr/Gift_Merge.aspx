<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Gift_Merge.aspx.cs" Inherits="DonorMgr_Gift_Merge" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>公關贈品資料合併</title>
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

            $("#checkboxAll").click(function () {

                if ($("#checkboxAll").prop("checked")) {
                    $("input[id=checkbox]").prop("checked", true);
                }
                else {
                    $("input[id=checkbox]").prop("checked", false);
                }
            });
        });
    </script>
    <script type="text/javascript">
        function Query_OnClick() {
            var strRet = "";
            var tbxFrom_Donor_Id = document.getElementById('tbxFrom_Donor_Id');
            var tbxTo_Donor_Id = document.getElementById('tbxTo_Donor_Id');
            if (tbxFrom_Donor_Id.value == '') {
                strRet += "『轉出的捐款人編號』 ";
            }
            if (tbxTo_Donor_Id.value == '') {
                strRet += "『欲轉入的捐款人編號』 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet);
                return false;
            }
            if (isNaN(Number(tbxFrom_Donor_Id.value)) == true) {
                alert('欲轉出的捐款人編號  欄位必須為數字！');
                tbxFrom_Donor_Id.focus();
                return false;
            }
            if (isNaN(Number(tbxTo_Donor_Id.value)) == true) {
                alert('欲轉入的捐款人編號  欄位必須為數字！');
                tbxTo_Donor_Id.focus();
                return false;
            }
        }
        function Transfer_OnClick() {
            if (window.confirm('您是否確定要合併公關贈品資料？') == true) {
                return true
            }
            else {
                return false;
            }
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
        公關贈品資料合併 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1">
                捐款人編號：
            </th>
            <td align="left" colspan="2" >
                <asp:TextBox runat="server" ID="tbxFrom_Donor_Id" CssClass="font9"></asp:TextBox>
            &nbsp;轉出--&gt; <asp:TextBox runat="server" ID="tbxTo_Donor_Id" CssClass="font9"></asp:TextBox>
            &nbsp;轉入</td>
            <td align="left" colspan="2">
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" OnClientClick="return Query_OnClick();" OnClick="btnQuery_Click"/>
                <asp:Button ID="btntransfer" CssClass="npoButton npoButton_New" runat="server"  Width="30mm"
                    Text="資料合併" OnClientClick="return Transfer_OnClick();" onclick="btntransfer_Click"/>
            </td>
        </tr>
        <tr>
            <td width="50%" colspan="3" align="left">
                 <asp:Label ID="lblDonor_From" runat="server"></asp:Label>
            </td>
            <td width="50%" colspan="3" align="left">
                 <asp:Label ID="lblDonor_To" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td width="50%" colspan="3" align="left" valign="top">
                 <asp:Label ID="lblData_From" runat="server"></asp:Label>
            </td>
            <td width="50%" colspan="3" align="left" valign="top">
                 <asp:Label ID="lblData_To" runat="server"></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
