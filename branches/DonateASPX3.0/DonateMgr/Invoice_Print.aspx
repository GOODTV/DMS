﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Invoice_Print.aspx.cs" Inherits="DonateMgr_Invoice_Print" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>收據列印</title>
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
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxDonate_DateS",   // id of the input field
                button: "imgDonate_DateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonate_DateE",   // id of the input field
                button: "imgDonate_DateE"     // 與觸發動作的物件ID相同
            });
        }
        function alertCon() {
            if ($("#ddlFormat").val() == "") {
                alert("名條格式 欄位不可為空白！");
                $("#ddlFormat").focus();
                return false;
            }
            else {
                return true;
            }
        }
        function Report_OnClick() {
            if (window.confirm('您是否確定要將查詢結果列印？') == true) {
                window.open('../DonateMgr/Invoice_Print_Data.aspx', 'Invoice_Print_Data', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
            }
        }
        function Address_OnClick(i) {
            var ddlFormat = document.getElementById('ddlFormat');
            if (ddlFormat.value == "") {
                alert('標籤格式 欄位不可為空')
                return false;
            }
            if (window.confirm('您是否確定要將查詢結果列印？') == true) {
                window.open('Invoice_Print/Invoice_Print_Style' + i + '.aspx', 'Invoice_Print_Style' + i, 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600', '');
            }
        }
        function Post_OnClick() {
            if (window.confirm('您是否確定要將查詢結果列印？') == true) {
                window.open('../DonateMgr/Invoice_Print_Post.aspx', 'Invoice_Print_Post', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
            }
        }
        function Help_OnClick() {
            window.open('../DonateMgr/Invoice_Help.htm', 'Invoice_Help', 'toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=300');

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
        收據列印</h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                經手人：
            </th>
            <td align="left" colspan="2" >
                <asp:dropdownlist runat="server" ID="ddlCreate_User" CssClass="font9" ></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
               捐款日期：
            </th>
            <td align="left" colspan="2" >
                <asp:TextBox runat="server" ID="tbxDonate_DateS" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgDonate_DateS" alt="" src="../images/date.gif" />
                ~<asp:TextBox runat="server" ID="tbxDonate_DateE" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgDonate_DateE" alt="" src="../images/date.gif" />
             </td>
             <th align="right" colspan="1">
                捐款人編號：
            </th>
            <td align="left" colspan="2" >
                <asp:TextBox runat="server" ID="tbxDonor_Id" CssClass="font9"></asp:TextBox>
                &nbsp;&nbsp;&nbsp;
                <asp:CheckBox ID="cbxInvoice" text = "收據重印" runat="server" />
                <asp:CheckBox ID="cbxAddress" text = "地址名條重印" runat="server" />
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="2" >
                <asp:dropdownlist runat="server" ID="ddlActName" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                收據編號： 
            </th>
            <td align="left" colspan="2" >
                <asp:TextBox runat="server" ID="tbxInvoice_NoS" CssClass="font9" Width="100px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="tbxInvoice_NoE" CssClass="font9" Width="100px"></asp:TextBox>
                <font color="#ff0000">(&nbsp;機構前置碼不需輸入&nbsp;)</font></td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                收據開立：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblInvoice_Type" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1" >
                 捐款方式：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblDonate_Payment" runat="server"  RepeatDirection="Horizontal"></asp:CheckBoxList>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1" >
                 捐款用途：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblDonate_Purpose" runat="server"  RepeatDirection="Horizontal"></asp:CheckBoxList>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1" >
                 地址：
            </th>
            <td align="left" colspan="5">
                <asp:RadioButtonList ID="rblLocal" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Value="1" Selected="True">國內</asp:ListItem>
                    <asp:ListItem Value="2">國外 </asp:ListItem>
                </asp:RadioButtonList>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                其他：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblExp_Pre" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                名條內容：
            </th>
            <td align="left" colspan="5">
                 <asp:RadioButtonList ID="rblContent" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Value="1" Selected="True">收據地址</asp:ListItem>
                    <asp:ListItem Value="2">通訊地址 </asp:ListItem>
                </asp:RadioButtonList>
            </td> 
        </tr>
        <tr>
            
            <th align="right" colspan="1">
                名條格式：
            </th>
            <td align="left" colspan="5">
                <asp:dropdownlist runat="server" ID="ddlFormat" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                排序方式：
            </th>
            <td align="left" colspan="4">
                 <asp:RadioButtonList ID="rblSort" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Value="0" Selected="True">收據編號</asp:ListItem>
                    <asp:ListItem Value="1">捐款人編號</asp:ListItem>
                    <asp:ListItem Value="2">郵遞區號</asp:ListItem>
                </asp:RadioButtonList>
            </td> 
            <td><asp:CheckBox ID="cbxPrint_NoName" text = "不含身份別主知名" runat="server" Checked="true" /></td>
        </tr>
        <tr>
            <td align="right" colspan="8" class="style1">
                <asp:Button ID="btnReport" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="收據" onclick="btnReport_Click"/>
                <asp:Button ID="btnAddress" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="名條" onclick="btnAddress_Click"/>
                <asp:Button ID="btnPost" CssClass="npoButton npoButton_Print" runat="server"  Width="25mm"
                    Text="大宗掛號" onclick="btnPost_Click"/>
                <asp:Button ID="btnHelp" CssClass="npoButton npoButton_Single" runat="server"  Width="20mm"
                    Text="說明" OnClientClick="Help_OnClick();"/>
            </td>
        </tr>
    </table>
    </div>
    </form>
    <p>
        &nbsp;</p>
</body>
</html>
