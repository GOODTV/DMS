<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Donate_New_InvoiceNo.aspx.cs" Inherits="DonateMgr_Donate_New_InvoiceNo" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>收據維護紀錄【重新取號】</title>
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
            var tbxInvoice_No = document.getElementById('tbxInvoice_No');
            tbxInvoice_No.readOnly = true;
            InitMenu();
        });
    </script>
    <script type="text/javascript">
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxDonate_Date",   // id of the input field
                button: "imgDonate_Date"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxRequest_Date",   // id of the input field
                button: "imgRequest_Date"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxAccoun_Date",   // id of the input field
                button: "imgAccoun_Date"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxInvoceSend_Date",   // id of the input field
                button: "imgInvoceSend_Date"     // 與觸發動作的物件ID相同
            });
        }
        function SumAccou() {
            var tbxDonate_Amt = document.getElementById('tbxDonate_Amt');
            var tbxDonate_Fee = document.getElementById('tbxDonate_Fee');
            var tbxDonate_Accou = document.getElementById('tbxDonate_Accou');
            if (isNaN(Number(tbxDonate_Amt.value)) || isNaN(Number(tbxDonate_Fee.value))) {
                tbxDonate_Accou.value = '0';
            } else {
                tbxDonate_Accou.value = Number(tbxDonate_Amt.value) - Number(tbxDonate_Fee.value);
            }
        }
        function Same_Owner_OnClick() {
            var Same_Owner = document.getElementById('cbxSame_Owner');
            var Card_Owner = document.getElementById('tbxCard_Owner');
            var Donor_Name = document.getElementById('tbxDonor_Name');
            var Owner_IDNo = document.getElementById('tbxOwner_IDNo');
            var DonorIDNo = document.getElementById('tbxIDNo');
            var Relation = document.getElementById('tbxRelation');

            if (Same_Owner.checked) {
                Card_Owner.value = Donor_Name.value;
                Owner_IDNo.value = DonorIDNo.value;
                Relation.value = '本人';
            }
        }
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var cnt = 0;
            var sName = '';
            var tbxDonate_Date = document.getElementById('tbxDonate_Date');
            var ddlDonate_Payment = document.getElementById('ddlDonate_Payment');
            var ddlDonate_Purpose = document.getElementById('ddlDonate_Purpose');
            var tbxDonate_Amt = document.getElementById('tbxDonate_Amt');
            var tbxDonate_Fee = document.getElementById('tbxDonate_Fee');
            var tbxDonate_Accou = document.getElementById('tbxDonate_Accou');
            var tbxDonate_Forign = document.getElementById('tbxDonate_Forign');
            var tbxInvoice_Title = document.getElementById('tbxInvoice_Title');
            var tbxInvoice_No = document.getElementById('tbxInvoice_No');
            var tbxRequest_Date = document.getElementById('tbxRequest_Date');
            var ddlAccoun_Bank = document.getElementById('ddlAccoun_Bank');
            var tbxAccoun_Date = document.getElementById('tbxAccoun_Date');
            var ddlDonate_Type = document.getElementById('ddlDonate_Type');
            var tbxDonation_NumberNo = document.getElementById('tbxDonation_NumberNo');
            var tbxDonation_SubPoenaNo = document.getElementById('tbxDonation_SubPoenaNo');
            var ddlAccounting_Title = document.getElementById('ddlAccounting_Title');
            var tbxInvoceSend_Date = document.getElementById('tbxInvoceSend_Date');
            var ddlAct_Id = document.getElementById('ddlAct_Id');
            var tbxComment = document.getElementById('tbxComment');
            var tbxInvoice_PrintComment = document.getElementById('tbxInvoice_PrintComment');

            if (tbxDonate_Date.value == "") {
                strRet += "捐款日期 ";
            }
            if (ddlDonate_Payment.value == "") {
                strRet += "捐款方式 ";
            }
            if (ddlDonate_Purpose.value == "") {
                strRet += "捐款用途 ";
            }
            if (tbxDonate_Amt.value == "") {
                strRet += "捐款金額 ";
            }
            if (tbxInvoice_No.value == "") {
                strRet += "收據編號 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet);
                return false;
            }
            cnt = 0;
            sName = ddlDonate_Payment.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('捐款方式 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = ddlDonate_Purpose.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('捐款用途 欄位長度超過限制！');
                return false;
            }
            if (isNaN(Number(tbxDonate_Amt.value)) == true) {
                alert('捐款金額 欄位必須為數字！');
                tbxDonate_Amt.focus();
                return false;
            }
            if (isNaN(Number(tbxDonate_Fee.value)) == true) {
                alert('手續費 欄位必須為數字！');
                tbxDonate_Fee.focus();
                return false;
            }
            if (tbxDonate_Fee.value == '') {
                tbxDonate_Fee.value = '0';
            }
            if (isNaN(Number(tbxDonate_Accou.value)) == true) {
                alert('實收金額  欄位必須為數字！');
                tbxDonate_Accou.focus();
                return false;
            }
            cnt = 0;
            sName = tbxDonate_Forign.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('外幣 欄位長度超過限制！');
                return false;
            }
            if (Number(tbxDonate_Accou.value) > Number(tbxDonate_Amt.value)) {
                alert('實收金額不可大於捐款金額');
                tbxDonate_Accou.focus();
                return false;
            }
            cnt = 0;
            sName = tbxInvoice_Title.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 100) {
                alert('收據抬頭 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxInvoice_No.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('收據編號 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            var sName = ddlAccoun_Bank.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('入帳銀行 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            var sName = ddlDonate_Type.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('捐款類別 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = ddlAccounting_Title.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 30) {
                alert('會計科目 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxDonation_NumberNo.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('傳票號碼 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxDonation_SubPoenaNo.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('劃撥 / 匯款單號 欄位長度超過限制！');
                return false;
            }
            var tbxInvoice_No = $('#HFD_Invoice_No').val()
            var str = "您是否確定要重新取號？\n";
            str += "原捐款紀錄『收據編號：" + tbxInvoice_No + "』將會被作廢";
            if (window.confirm(str) == true) {
                return true;
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
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <asp:HiddenField runat="server" ID="HFD_Donor_Id" />
    <asp:HiddenField runat="server" ID="HFD_Invoice_No" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        收據維護紀錄【重新取號】 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1">
                捐款人：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right"colspan="1" >
                類別：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxCategory" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                收據地址：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxAddress" CssClass="font9"  Width="250px" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                身分別：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxDonor_Type" CssClass="font9"
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                身分證/統編：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIDNo" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        
        <tr>
            <td colspan="8">

            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Date" runat="server" onchange="CheckDateFormat(this, '捐款日期');" ></asp:TextBox>
                <img id="imgDonate_Date" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                捐款方式：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Payment" CssClass="font9" ></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                捐款用途：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Purpose" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                收據開立：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxInvoice_Type" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款金額：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Amt" runat="server" OnPropertyChange="SumAccou()" ></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                手續費：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Fee" Text="0" runat="server" OnPropertyChange="SumAccou()" ></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                實收金額：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox ID="tbxDonate_Accou" Text="0" runat="server" ></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                外幣幣別：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Forign" runat="server" ></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                外幣金額：
            </th>
            <td align="left" colspan="5">
                <asp:TextBox ID="tbxDonate_ForignAmt" runat="server"></asp:TextBox>
            </td>
        </tr>
        <asp:Panel ID="PanelCheck" runat="server"  Visible ="false">
        <tr>
            <th align="right"colspan="1" >
                支票號碼：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCheck_No" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                到期日：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxCheck_ExpireDate" runat="server" onchange="CheckDateFormat(this, '到期日');"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <td colspan="4">
                
            </td>
        </tr>
        </asp:Panel>
        <asp:Panel ID="PanelCreditCard" runat="server" Visible ="false">
        <tr>
            <th align="right" colspan="1">
                銀行別：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCard_Bank" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                信用卡別：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCard_Type" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                信用卡號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxAccount_No1" CssClass="font9" Width="28px"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>-
                <asp:TextBox runat="server" ID="tbxAccount_No2" CssClass="font9" Width="28px"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>-
                <asp:TextBox runat="server" ID="tbxAccount_No3" CssClass="font9" Width="28px"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>-
                <asp:TextBox runat="server" ID="tbxAccount_No4" CssClass="font9" Width="28px"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                有效月年：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxValid_Date" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                持卡人：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCard_Owner" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True" ></asp:TextBox>
                <asp:checkbox runat="server" ID="cbxSame_Owner" Text="同捐款人" CssClass="font9" Enabled="False"></asp:checkbox>
            </td>
            <th align="right" colspan="1">
                持卡人身分證：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxOwner_IDNo" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                與捐款人關係：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxRelation" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                授權碼：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxAuthorize" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        </asp:Panel>
        <asp:Panel ID="PanelAccount" runat="server"  Visible ="false">
        <tr>
             <th align="right" colspan="1">
                存簿戶名：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_Name" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
                <asp:checkbox runat="server" ID="cbxSame_Owner2" Text="同捐款人" CssClass="font9"  Enabled="False"></asp:checkbox>
            </td>
            <th align="right" colspan="1">
                持有人身分證：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_IDNo" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                存簿局號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_SavingsNo" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                存簿帳號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_AccountNo" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        </asp:Panel>
        <tr>
            <td colspan="8">

            </td>
        </tr>
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDept" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                收據抬頭：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxInvoice_Title" CssClass="font9" Width="210px" ></asp:TextBox>
            </td>
            <td align="center" colspan="1">
                <asp:checkbox runat="server" ID="cbxInvoice_Pre" Text="手開收據" CssClass="font9" Enabled="False" 
                ></asp:checkbox>
                
            </td>
            <th align="right" colspan="1">
                收據編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxInvoice_No" CssClass="font9" 
                    BackColor="#FFE1AF"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                請款日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxRequest_Date" runat="server" onchange="CheckDateFormat(this, '請款日期');" ></asp:TextBox>
                <img id="imgRequest_Date" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                入帳銀行：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlAccoun_Bank" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                沖帳日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxAccoun_Date" runat="server" onchange="CheckDateFormat(this, '沖帳日期');" ></asp:TextBox>
                <img id="imgAccoun_Date" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                捐款類別：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Type" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                傳票號碼：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonation_NumberNo" CssClass="font9" ></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                劃撥 / 匯款單號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonation_SubPoenaNo" CssClass="font9" ></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                會計科目：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlAccounting_Title" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                收據寄送：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxInvoceSend_Date" runat="server" onchange="CheckDateFormat(this, '收據寄送');" ></asp:TextBox>
                <img id="imgInvoceSend_Date" alt="" src="../images/date.gif" />
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                專案活動：
            </th>
            <td align="left" colspan="3">
                <asp:dropdownlist runat="server" ID="ddlAct_Id" CssClass="font9"></asp:dropdownlist>
            </td>
         </tr>
         <tr>
            <th colspan="1" align="right">
                捐款備註：
            </th>
            <td colspan="3">
                <asp:Textbox runat="server" ID="tbxComment" CssClass="font9" 
                    TextMode="MultiLine" Width="450px" Height="100px"></asp:Textbox> 
            </td>
            <th colspan="1" align="right">
                收據備註：<br />
               (列印用) 
            </th>
            <td colspan="3">
                <asp:Textbox runat="server" ID="tbxInvoice_PrintComment" CssClass="font9" 
                    TextMode="MultiLine" Width="450px" Height="100px"></asp:Textbox> 
            </td>
         </tr>
    </table>
    </div>
    <div class="function">
         <asp:Button ID="btn_Re_No" class="npoButton npoButton_Modify" runat="server" 
            Text="重新取號" onclick="btn_Re_No_Click"  OnClientClick= "return CheckFieldMustFillBasic();" />
        <asp:Button ID="btnExit" class="npoButton npoButton_Cancel" runat="server"    
            Text="取消" Width="80px" onclick="btnExit_Click1"  />
    </div>
    </form>
</body>
</html>


