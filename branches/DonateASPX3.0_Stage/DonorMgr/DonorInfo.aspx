<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonorInfo.aspx.cs" Inherits="DonorMgr_DonatorInfo" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>捐款人基本資料維護</title>
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
                window.open('../DonorMgr/DonorInfo_Print.aspx', 'DonorInfo_Print', 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
            }
        }
        function CheckFieldMustFillBasic() {
            var tbxDonor_Id = document.getElementById('tbxDonor_Id');
            var tbxDonor_Id_Old = document.getElementById('tbxDonor_Id_Old');
            var tbxEMail = document.getElementById('tbxEMail');

            if (isNaN(Number(tbxDonor_Id.value)) == true) {
                alert('捐款人編號 欄位必須為數字！');
                tbxDonor_Id.focus();
                return false;
            }
            if (isNaN(Number(tbxDonor_Id_Old.value)) == true) {
                alert('舊編號 欄位必須為數字！');
                tbxDonor_Id_Old.focus();
                return false;
            }
            if (tbxEMail.value == "@") {
                alert('E-Mail 欄位不能只有查詢"@"符號！');
                tbxEMail.focus();
                return false;
            }
            return true;
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
        捐款人基本資料維護 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
             <th align="right" colspan="1">
                捐款人：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right"colspan="1" >
                捐款人編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Id" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1" >
                舊編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Id_Old" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                連絡電話：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxTel_Office" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1" >
                身份別：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonor_Type" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                文宣品：
            </th>
            <td align="left" colspan="3">
                紙本月刊本數<asp:Textbox runat="server" ID="tbxIsSendNewsNum" CssClass="font9" Width="50px" ></asp:Textbox>
                <asp:checkbox runat="server" ID="cbxIsDVD" Text="DVD" CssClass="font9"></asp:checkbox>
                <asp:checkbox runat="server" ID="cbxIsSendEpaper" Text="電子文宣" CssClass="font9"></asp:checkbox>
            </td> 
            <th align="right" colspan="1">
                收據抬頭：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxInvoice_Title" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                身份證/統編：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxIDNo" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                經手人：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlCreate_User" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                E-Mail：
            </th>
            <td align="left" colspan="9">
                <asp:TextBox runat="server" ID="tbxEMail" CssClass="font9" Width="300px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                地址：
            </th>
            <td align="left" colspan="9" >
                <asp:checkbox runat="server" ID="cbxIsAbroad" Text="海外地址" CssClass="font9" 
                    AutoPostBack="True" oncheckedchanged="cbxIsAbroad_CheckedChanged"></asp:checkbox>
            　　<asp:Panel runat="server" ID="PanelAbroad">
                    <asp:TextBox runat="server" ID="tbxIsAbroad_Address" CssClass="font9" Width="500px" maxlength="200"></asp:TextBox>
                </asp:Panel>
                <asp:Panel runat="server" ID="PanelLocal">
                    <asp:TextBox runat="server" ID="tbxZipCode" CssClass="font9" Width="60px"></asp:TextBox>
                    <asp:dropdownlist runat="server" ID="ddlCity" CssClass="font9" 
                         AutoPostBack="True" onselectedindexchanged="ddlCity_SelectedIndexChanged"></asp:dropdownlist>
                    <asp:dropdownlist runat="server" ID="ddlArea" CssClass="font9" 
                         AutoPostBack="True" onselectedindexchanged="ddlArea_SelectedIndexChanged"></asp:dropdownlist>
                    <asp:TextBox runat="server" ID="tbxStreet" CssClass="font9" Width="100px"></asp:TextBox>大道/路/街/部落
                    <asp:dropdownlist runat="server" ID="ddlSection" CssClass="font9"></asp:dropdownlist>段
                    <asp:TextBox runat="server" ID="tbxLane" CssClass="font9" Width="40px"></asp:TextBox>巷
                    <asp:TextBox runat="server" ID="tbxAlley" CssClass="font9" Width="40px"></asp:TextBox>
                    <asp:dropdownlist runat="server" ID="ddlAlley" CssClass="font9"></asp:dropdownlist>
                    <asp:TextBox runat="server" ID="tbxNo1" CssClass="font9" Width="55px"></asp:TextBox>號之
                    <asp:TextBox runat="server" ID="tbxNo2" CssClass="font9" Width="40px"></asp:TextBox>
                    <asp:TextBox runat="server" ID="tbxFloor1" CssClass="font9" Width="40px"></asp:TextBox>樓之
                    <asp:TextBox runat="server" ID="tbxFloor2" CssClass="font9" Width="40px"></asp:TextBox>&nbsp;
                    <asp:TextBox runat="server" ID="tbxRoom" CssClass="font9" Width="40px"></asp:TextBox>室
                </asp:Panel>
            
            </td> 
        </tr>
        <tr>
            <td colspan="1"></td>
            <td align="left" colspan="2"><font color="blue">1.捐款人編號、舊編號、連絡電話、台灣地址等欄位是精準查詢<br />2.其他欄位是模糊(雷同)查詢<br />3.此查詢base僅有捐款人資料。</font></td>
            <td align="right" colspan="8" class="style1">
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" OnClientClick= "return CheckFieldMustFillBasic();" OnClick="btnQuery_Click"/>
                <!--<asp:Button ID="btnFuzzyQuery" CssClass="npoButton npoButton_Search" runat="server" Width="35mm"
                    Text="雷同資料查詢" OnClientClick= "return CheckFieldMustFillBasic();" onclick="btnFuzzyQuery_Click"/>-->
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="列表" OnClientClick="Print();"/>
                <asp:Button ID="btnToxls" runat="server"  Width="20mm"
                    Text="匯出" OnClick="btnToxls_Click" CssClass="npoButton npoButton_Excel" OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');"/>
                <asp:Button ID="btnToPhone" runat="server"  Width="25mm"
                    Text="手機名單" OnClick="btnToPhone_Click" CssClass="npoButton npoButton_Excel" OnClientClick=" return confirm('您是否確定要將查詢結果手機名單？');"/>
                <asp:Button ID="brnToEMail" runat="server"  Width="27mm"
                    Text="EMail名單" OnClick="btnToEMail_Click" 
                    CssClass="npoButton npoButton_Excel" Height="34px" OnClientClick=" return confirm('您是否確定要將查詢結果？');"/>
                <asp:Button ID="btnAdd" CssClass="npoButton npoButton_New" runat="server"  Width="20mm"
                    Text="新增" OnClick="btnAdd_Click" />
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="10" align ="center">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>