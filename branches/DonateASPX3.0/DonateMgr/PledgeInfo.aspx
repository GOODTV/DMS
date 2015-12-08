<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PledgeInfo.aspx.cs" Inherits="DonateMgr_PledgeQry" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>固定轉帳授權書</title>
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
                inputField: "txtDonateDate",   // id of the input field
                button: "imgDonateDate"     // 與觸發動作的物件ID相同
            }); 
        }
        function Print(PrintType) {
            window.open('../DonateMgr/PledgeInfo_Print.aspx', 'PledgeInfo_Print', 'width=900,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
        }
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var cnt = 0;
            var sName = '';
            var tbxAccount_No = document.getElementById('tbxAccount_No');
            var tbxDonor_Id = document.getElementById('tbxDonor_Id'); 
            cnt = 0;
            sName = tbxAccount_No.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 16) {
                alert('卡號 欄位長度超過限制！');
                return false;
            }
            if (isNaN(Number(tbxDonor_Id.value)) == true) {
                alert('捐款人編號 欄位必須為數字！');
                tbxDonor_Id.focus();
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
    <form id="Form1" runat="server" defaultbutton="btnQuery">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        固定轉帳授權書 
    </h1>
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
                授權編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPledge_Id" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                捐款人編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Id" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right"colspan="1" >
                捐款授權到期日：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlYear_Donate_ToDate" CssClass="font9"></asp:dropdownlist>
                年
                <asp:dropdownlist runat="server" ID="ddlMonth_Donate_ToDate" CssClass="font9"></asp:dropdownlist>
                月
            </td>
            <th align="right" colspan="1">
                持卡人：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxCard_Owner" CssClass="font9"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                授權方式：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Payment" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                指定用途：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Purpose" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                卡號/帳號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxAccount_No" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                信用卡有效月年：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlMonth_Valid_Date" CssClass="font9"></asp:dropdownlist>
                月
                <asp:dropdownlist runat="server" ID="ddlYear_Valid_Date" CssClass="font9"></asp:dropdownlist>
                年
            </td> 
            <td align="right" colspan="1">
                <asp:Button ID="btnAuthorize" CssClass="npoButton npoButton_New" 
                    runat="server"  Width="35mm"
                    Text="授權到期通知" onclick="btnAuthorize_Click"/>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                轉帳日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="txtDonateDate" runat="server" 
                    onchange="CheckDateFormat(this, '轉帳日期');" Width="90px"></asp:TextBox>
                <img id="imgDonateDate" alt="" src="../images/date.gif" /> </td>
            <th align="right" colspan="1">
                轉帳週期：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Period" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                狀態：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlStatus" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                經手人：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlCreate_User" CssClass="font9"></asp:dropdownlist>
            </td> 
            <td align="right" colspan="1">
                <asp:Button ID="btnCreditCard" CssClass="npoButton npoButton_New" 
                    runat="server"  Width="35mm"
                    Text="信用卡到期通知" onclick="btnCreditCard_Click"/>
            </td>
        </tr>
        <tr>
            <td align="left" colspan="2" class="style1">
                扣款金額合計：
                <asp:Label ID="lblAmt" runat="server" Text=""></asp:Label>
                元</td>
            <td align="left" colspan="4" class="style1">
                &nbsp;
            </td>
            <td align="right" colspan="4" class="style1">
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" onclick="btnQuery_Click" OnClientClick= "return CheckFieldMustFillBasic();"/>
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="列表" 
                    OnClientClick="if (confirm('您是否確定要將查詢結果匯出？')==false) {return false;} Print('');"/>
                <asp:Button ID="btnToxls" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" 
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" onclick="btnToxls_Click"/>
                <asp:Button ID="btnAdd" CssClass="npoButton npoButton_New" runat="server"  Width="25mm"
                    Text="新增資料" onclick="btnAdd_Click"/>
            </td>
        </tr>
        <tr>
            <td align="center" width="100%" colspan="10">
                 <asp:Label ID="lblGridList" runat="server" ></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
