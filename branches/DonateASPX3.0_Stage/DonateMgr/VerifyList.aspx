<%@ Page Language="C#" AutoEventWireup="true" CodeFile="VerifyList.aspx.cs" Inherits="DonateMgr_VerifyList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>奉獻徵信錄</title>
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
        function Print() {
            if (window.confirm('您是否確定要將查詢結果匯出？') == true) {
                var Form_name = document.getElementById('Form1');
                for (var i = 0; i < Form_name.rblDonate_Purpose_Type.length; i++) {
                    if (Form_name.rblDonate_Purpose_Type[i].checked) {
                        var style = Form_name.rblDonate_Purpose_Type[i].value;
                    }
                }
                window.open('../DonateMgr/VerifyList_Print_style' + style + '.aspx', 'DonorQry_Print_style' + style, 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
                return true;
            }
            else {
                return false;
            }
        }
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxDonateDateS",   // id of the input field
                ifFormat: "%Y/%m/%d",       // format of the input field
                showsTime: false,
                timeFormat: "24",
                button: "imgDonateDateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE",   // id of the input field
                ifFormat: "%Y/%m/%d",       // format of the input field
                showsTime: false,
                timeFormat: "24",
                button: "imgDonateDateE"     // 與觸發動作的物件ID相同
            });
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
        奉獻徵信錄</h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                捐款機構：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>       
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款人姓名：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
            </td>
        </tr>
        <tr>
             <th align="right" colspan="1">
               捐款日期：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonateDateS" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgDonateDateS" alt="" src="../images/date.gif" />
                ~<asp:TextBox runat="server" ID="tbxDonateDateE" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgDonateDateE" alt="" src="../images/date.gif" />
             </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlAct_Id" CssClass="font9"></asp:dropdownlist>
            </td>
         </tr>
        <tr>
            <th align="right" colspan="1">
                條列方式：
            </th>
            <td align="left" colspan="1" >
                <asp:RadioButtonList ID="rblDonate_Purpose_Type" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Value="1" Selected="True">日期條列</asp:ListItem>
                    <asp:ListItem Value="2">金額條列</asp:ListItem>
                    <asp:ListItem Value="3">捐款用途條列</asp:ListItem>
                </asp:RadioButtonList>
            </td>
        </tr>
        <tr>
            <td align="right" colspan="2" class="style1">
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"/>
                <asp:Button ID="btnToxls" CssClass="npoButton npoButton_Excel" runat="server"  Width="20mm"
                    Text="匯出" OnClientClick="return confirm('您是否確定要將查詢結果匯出？');"  onclick="btnToxls_Click"/>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>