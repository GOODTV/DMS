<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Contribute_Add.aspx.cs" Inherits="ContributeMgr_Contribute_Add" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>物品捐贈輸入</title>
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
    </script>
    <script type="text/javascript">
        window.onload = initCalendar;
        function ReadOnly() {
            var tbxInvoice_No = document.getElementById('tbxInvoice_No');
            tbxInvoice_No.readOnly = true;
            tbxInvoice_No.style.backgroundColor = '#FFE1AF';
            tbxInvoice_No.value = "自動編號";
            for (i = 1; i <= 5; i++) {
                var Goods_Name = document.getElementById('tbxGoods_Name_' + i);
                var Goods_Unit = document.getElementById('tbxGoods_Unit_' + i);
                var Contribute_IsStock = document.getElementById('cbxContribute_IsStock_' + i); 
                Goods_Name.readOnly = true;
                Goods_Name.style.backgroundColor = '#FFE1AF';
                Goods_Unit.readOnly = true;
                Goods_Unit.style.backgroundColor = '#FFE1AF';
                Contribute_IsStock.checked = true;
            }
        }
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxContribute_Date",   // id of the input field
                button: "imgContribute_Date"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxAccoun_Date",   // id of the input field
                button: "imgAccoun_Date"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxGoods_DueDate_1",   // id of the input field
                button: "imgGoods_DueDate_1"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxGoods_DueDate_2",   // id of the input field
                button: "imgGoods_DueDate_2"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxGoods_DueDate_3",   // id of the input field
                button: "imgGoods_DueDate_3"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxGoods_DueDate_4",   // id of the input field
                button: "imgGoods_DueDate_4"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxGoods_DueDate_5",   // id of the input field
                button: "imgGoods_DueDate_5"     // 與觸發動作的物件ID相同
            });
        }
        function cbxInvoice_Pre_OnClick() {
            var cbxInvoice_Pre = document.getElementById('cbxInvoice_Pre');
            var tbxInvoice_No = document.getElementById('tbxInvoice_No');
            if (cbxInvoice_Pre.checked == true) {
                tbxInvoice_No.style.backgroundColor = '#ffffff';
                tbxInvoice_No.readOnly = false;
                tbxInvoice_No.value = '';
                tbxInvoice_No.focus();
            } else {
                tbxInvoice_No.style.backgroundColor = '#FFE1AF';
                tbxInvoice_No.readOnly = true;
                tbxInvoice_No.value = '自動編號';
            }
        }
        function Contribute_Cancel_OnClick(i) {
            var Goods_Id = document.getElementById('HFD_Goods_Id_' + i);
            var Goods_Name = document.getElementById('tbxGoods_Name_' + i);
            var Goods_Qty = document.getElementById('tbxGoods_Qty_' + i);
            var Goods_Unit = document.getElementById('tbxGoods_Unit_' + i);
            var Goods_Amt = document.getElementById('tbxGoods_Amt_' + i);
            var Goods_DueDate =  document.getElementById('tbxGoods_DueDate_' + i);
            var Goods_Comment = document.getElementById('tbxGoods_Comment_' + i);
            if ( Goods_Name.value != '') {
                if (confirm('您是否確定要清除『 ' +  Goods_Name.value + ' 』？')) {
                    Goods_Id.value = '';
                    Goods_Name.value = '';
                    Goods_Qty.value = '';
                    Goods_Unit.value = '';
                    Goods_Amt.value = '0';
                    Goods_DueDate.value = '';
                    Goods_Comment.value = '';
                }
            } else {
                if (confirm('您是否確定要清除？')) {
                    Goods_Id.value = '';
                    Goods_Name.value = '';
                    Goods_Qty.value = '';
                    Goods_Unit.value = '';
                    Goods_Amt.value = '0';
                    Goods_DueDate.value = '';
                    Goods_Comment.value = '';
                }
            }
        }
        function cbxContribute_IsStock_OnClick(i) {
            var Goods_Id = document.getElementById('HFD_Goods_Id_' + i);
            var Goods_Name = document.getElementById('tbxGoods_Name_' + i);
            var Goods_Qty = document.getElementById('tbxGoods_Qty_' + i);
            var Goods_Unit = document.getElementById('tbxGoods_Unit_' + i);
            var Goods_Amt = document.getElementById('tbxGoods_Amt_' + i);
            var Goods_DueDate = document.getElementById('tbxGoods_DueDate_' + i);
            var Goods_Comment = document.getElementById('tbxGoods_Comment_' + i);
            var cbxContribute_IsStock = document.getElementById('cbxContribute_IsStock_' + i);
            if (cbxContribute_IsStock.checked == false) {
                Goods_Id.value = '';
                Goods_Name.value = '';
                Goods_Qty.value = '';
                Goods_Unit.value = '';
                Goods_Amt.value = '0';
                Goods_DueDate.value = '';
                Goods_Comment.value = '';
                Goods_Name.readOnly = false;
                Goods_Unit.readOnly = false;
                Goods_Name.style.backgroundColor = '#ffffff';
                Goods_Unit.style.backgroundColor = '#ffffff';
            }
            else {
                Goods_Id.value = '';
                Goods_Name.value = '';
                Goods_Qty.value = '';
                Goods_Unit.value = '';
                Goods_Amt.value = '0';
                Goods_DueDate.value = '';
                Goods_Comment.value = '';
                Goods_Name.readOnly = true;
                Goods_Unit.readOnly = true;
                Goods_Name.style.backgroundColor = '#FFE1AF';
                Goods_Unit.style.backgroundColor = '#FFE1AF';
            }
        }
        function WindowsOpen(i) {
            window.open('GoodsShow.aspx?Goods_Id=HFD_Goods_Id_' + i + '&Goods_Name=tbxGoods_Name_' + i + '&Goods_Unit=tbxGoods_Unit_' + i, 'NewWindows',
                        'status=no,scrollbars=yes,top=100,left=120,width=500,height=450);');
        }
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var tbxContribute_Date = document.getElementById('tbxContribute_Date');
            var ddlContribute_Payment = document.getElementById('ddlContribute_Payment');
            var ddlContribute_Purpose = document.getElementById('ddlContribute_Purpose');
            var ddlInvoice_Type = document.getElementById('ddlInvoice_Type');
            var ddlDept = document.getElementById('ddlDept');
            var tbxInvoice_Title = document.getElementById('tbxInvoice_Title');
            var tbxInvoice_No = document.getElementById('tbxInvoice_No');
            var ddlAccounting_Title = document.getElementById('ddlAccounting_Title'); 
            if (tbxContribute_Date.value == "") {
                strRet += "捐贈日期 ";
            }
            if (ddlContribute_Payment.value == "") {
                strRet += "捐贈方式 ";
            }
            if (ddlContribute_Purpose.value == "") {
                strRet += "捐贈用途 ";
            }
            if (ddlInvoice_Type.value == "") {
                strRet += "收據開立 ";
            }
            if (ddlDept.value == "") {
                strRet += "機構名稱 ";
            }
            if (tbxInvoice_No.value == "") {
                strRet += "收據編號 ";
            }
            if (tbxInvoice_Title.value == "") {
                strRet += "收據抬頭 ";
            }
            if ($("#tbxGoods_Name_1").val() == "" && $("#tbxGoods_Name_2").val() == "" && $("#tbxGoods_Name_3").val() == "" && $("#tbxGoods_Name_4").val() == "" && $("#tbxGoods_Name_5").val() == "") {
                strRet += "物品名稱 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet)
                return false;
            }
            
            for (i = 1; i <= 5; i++) {
                var Contribute_IsStock = document.getElementById('cbxContribute_IsStock_' + i);
                var Goods_Id = document.getElementById('HFD_Goods_Id_' + i);
                var Goods_Name = document.getElementById('tbxGoods_Name_' + i);
                var Goods_Qty = document.getElementById('tbxGoods_Qty_' + i);
                var Goods_Unit = document.getElementById('tbxGoods_Unit_' + i);
                var Goods_Amt = document.getElementById('tbxGoods_Amt_' + i); 
                var Goods_Comment = document.getElementById('tbxGoods_Comment_' + i);
                var Contribute_IsStock = document.getElementById('cbxContribute_IsStock_' + i);
                if (Goods_Name.value != '') {
                    if (Contribute_IsStock.checked && Goods_Id.value == '') {
                        alert('您輸入的『' + Goods_Name.value + '』該商品代號不存在！');
                        return false;
                    }
                }
                cnt = 0;
                sName = Goods_Name.value
                for (var j = 0; j < sName.length; j++) {
                    if (escape(sName.charAt(j)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 50) {
                    alert('『' + Goods_Name.value + '』物品名稱  欄位長度超過限制！');
                    return false;
                }
                if (Goods_Name.value != '' && Goods_Qty.value == '') {
                    alert('『' + Goods_Name.value + '』數量  欄位不可為空白！');
                    Goods_Qty.focus();
                    return false;
                }
                if (isNaN(Number(Goods_Qty.value)) == true) {
                    alert('『' + Goods_Name.value + '』數量  欄位必須為數字！');
                    Goods_Qty.focus();
                    return false;
                }
                cnt = 0;
                sName = Goods_Qty.value
                for (var j = 0; j < sName.length; j++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 7) {
                    alert('『' + Goods_Name.value + '』數量 欄位長度超過限制！');
                    return false;
                }
                if (Goods_Name.value != '' && Goods_Unit.value == '') {
                    alert('『' + Goods_Name.value + '』單位 欄位不可為空白！');
                    Goods_Unit.focus();
                    return false;
                }
                cnt = 0;
                sName = Goods_Unit.value
                for (var j = 0; j < sName.length; j++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 10) {
                    alert('『' + Goods_Name.value + '』單位 欄位長度超過限制！');
                    return false;
                }
                if (isNaN(Number(Goods_Amt.value)) == true) {
                    alert('『' + Goods_Name.value + '』折合現金 欄位必須為數字！');
                    Goods_Amt.focus();
                    return false;
                }
                cnt = 0;
                sName = Goods_Amt.value
                for (var j = 0; j < sName.length; j++) {
                    if (escape(sName.charAt(j)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 7) {
                    alert('『' + Goods_Name.value + '』折合現金  欄位長度超過限制！');
                    return false;
                }

                if (Goods_Comment.value != '') {
                    cnt = 0;
                    sName = Goods_Comment.value
                    for (var j = 0; j < sName.length; j++) {
                        if (escape(sName.charAt(j)).length >= 4) cnt += 2;
                        else cnt++;
                    }
                    if (cnt > 100) {
                        alert('『' + Goods_Name_.value + '』備註 欄位長度超過限制！');
                        return false;
                    }
                }
            }

            for (j = i + 1; j < 5; j++) {
                if (document.getElementById('Goods_Id_' + i).value == document.getElementById('Goods_Id_' + j).value && document.getElementById('Goods_Name_' + i).value == document.getElementById('Goods_Name_' + j).value && document.getElementById('Goods_Unit_' + i).value == document.getElementById('Goods_Unit_' + j).value) {
                    alert('『' + document.getElementById('Goods_Name_' + j).value + '』物品名稱重覆出現！');
                    return false;
                }
            }

            cnt = 0;
             sName = ddlContribute_Payment.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('捐贈方式 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = ddlContribute_Purpose.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('捐贈用途 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = ddlInvoice_Type.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('收據開立 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxInvoice_Title.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 80) {
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
            sName = ddlAccounting_Title.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 30) {
                alert('會計科目 欄位長度超過限制！');
                return false;
            }


            return true;
        } 
    </script>
    <style type="text/css">
        .style2
        {
            width: 110px;
        }
    </style>
    </head>
<body class="body">
    <form id="Form2" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <asp:HiddenField runat="server" ID="HFD_Invoice_No" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        物品捐贈輸入
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1">
                捐贈人：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <%--<th align="right"colspan="1" >
                類別：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxCategory" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>--%>
            <th align="right" colspan="1">
                身分別：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxDonor_Type" CssClass="font9"
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                收據地址：
            </th>
            <td align="left" colspan="4">
                <asp:TextBox runat="server" ID="tbxAddress" CssClass="font9"  Width="250px" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                收據開立：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxInvoice_Type" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                身分證/統編：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIDNo" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                最近捐贈日：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxLast_DonateDateC" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        <tr>
           <th colspan="1" align="right">
                捐款人備註：
            </th>
            <td colspan="3">
                <asp:Textbox runat="server" ID="tbxRemark" TextMode="MultiLine" 
                     Width="450px" Height="50px" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True" ></asp:Textbox> 
            </td>
        </tr>
        <tr>
            <td colspan="9">

            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐贈日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxContribute_Date" runat="server" onchange="CheckDateFormat(this, '捐贈日期');"></asp:TextBox>
                <img id="imgContribute_Date" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                捐款方式：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlContribute_Payment" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                捐款用途：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlContribute_Purpose" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1" class="style2">
                收據開立：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlInvoice_Type" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr> 
        <tr>
            <td colspan="8">

            </td>
        </tr>
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9" 
                    AutoPostBack="True"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                收據抬頭：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxInvoice_Title" CssClass="font9" Width="210px"></asp:TextBox>
            </td>
            <td align="center" colspan="1">
                <asp:checkbox runat="server" ID="cbxInvoice_Pre" Text="手開收據" CssClass="font9" OnClick="cbxInvoice_Pre_OnClick()" ></asp:checkbox>
            </td>
            <th align="right" colspan="1" class="style2">
                收據編號：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxInvoice_No" CssClass="font9"
                   ></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                沖帳日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxAccoun_Date" runat="server" onchange="CheckDateFormat(this, '沖帳日期');"></asp:TextBox>
                <img id="imgAccoun_Date" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                會計科目：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlAccounting_Title" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="4">
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
            <td colspan="4">
                <asp:Textbox runat="server" ID="tbxInvoice_PrintComment" CssClass="font9" 
                    TextMode="MultiLine" Width="450px" Height="100px"></asp:Textbox> 
            </td>
         </tr>
         <tr>
            <th colspan="1" align="center">
                寫入庫存
            </th>
            <th colspan="1" align="center">
                物品名稱
            </th>
            <th colspan="1" align="center">
                物品數量
            </th>
            <th colspan="1" align="center">
                物品單位
            </th>
            <th colspan="1" align="center">
                折合現金
            </th>
            <th colspan="1" align="center">
                物品保存期限
            </th>
            <th colspan="2" align="center" width="15%">
                備註
            </th>
            <th colspan="1" width="6%" align="center">
                清除
            </th>
         </tr>
         <tr>
            <td colspan="1" align="center">
                <asp:HiddenField runat="server" ID="HFD_Goods_Id_1"/>
                <asp:CheckBox ID="cbxContribute_IsStock_1" runat="server" OnClick="cbxContribute_IsStock_OnClick('1')" />
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Name_1" runat="server" ></asp:TextBox>&nbsp;<a href onclick="WindowsOpen(1)" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Qty_1" runat="server"></asp:TextBox>
            </td>
            <td align="center">
                <asp:TextBox ID="tbxGoods_Unit_1" runat="server"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Amt_1" runat="server" Text="0"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_DueDate_1" runat="server" Width="90px"  onchange="CheckDateFormat(this, '物品保存期限 ');"></asp:TextBox>
                <img id="imgGoods_DueDate_1" alt="" src="../images/date.gif" />
            </td>
            <td colspan="2" align="center">
                <asp:TextBox ID="tbxGoods_Comment_1" runat="server" Width="160px" MaxLength="100" ></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <img border="0" src="../images/toolbar_cancel.gif" width="20" onClick="Contribute_Cancel_OnClick('1')" style="cursor:hand">
            </td>
         </tr>
         <tr>
            <td colspan="1" align="center">
                <asp:HiddenField runat="server" ID="HFD_Goods_Id_2" />
                <asp:CheckBox ID="cbxContribute_IsStock_2" runat="server" OnClick="cbxContribute_IsStock_OnClick('2')"/>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Name_2" runat="server"></asp:TextBox>&nbsp;<a href onclick="WindowsOpen(2)" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Qty_2" runat="server"></asp:TextBox>
            </td>
            <td align="center">
                <asp:TextBox ID="tbxGoods_Unit_2" runat="server"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Amt_2" runat="server" Text="0"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_DueDate_2" runat="server" Width="90px" onchange="CheckDateFormat(this, '物品保存期限 ');"></asp:TextBox>
                <img id="imgGoods_DueDate_2" alt="" src="../images/date.gif" />
            </td>
            <td colspan="2" align="center">
                <asp:TextBox ID="tbxGoods_Comment_2" runat="server" Width="160px" MaxLength="100" ></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <img border="0" src="../images/toolbar_cancel.gif" width="20" onClick="Contribute_Cancel_OnClick('2')" style="cursor:hand">
            </td>
         </tr>
         <tr>
            <td colspan="1" align="center">
                <asp:HiddenField runat="server" ID="HFD_Goods_Id_3" />
                <asp:CheckBox ID="cbxContribute_IsStock_3" runat="server" OnClick="cbxContribute_IsStock_OnClick('3')"/>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Name_3" runat="server"></asp:TextBox>&nbsp;<a href onclick="WindowsOpen(3)" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Qty_3" runat="server"></asp:TextBox>
            </td>
            <td align="center">
                <asp:TextBox ID="tbxGoods_Unit_3" runat="server"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Amt_3" runat="server" Text="0"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_DueDate_3" runat="server" Width="90px" onchange="CheckDateFormat(this, '物品保存期限 ');"></asp:TextBox>
                <img id="imgGoods_DueDate_3" alt="" src="../images/date.gif" />
            </td>
            <td colspan="2" align="center">
                <asp:TextBox ID="tbxGoods_Comment_3" runat="server" Width="160px" MaxLength="100" ></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <img border="0" src="../images/toolbar_cancel.gif" width="20" onClick="Contribute_Cancel_OnClick('3')" style="cursor:hand">
            </td>
         </tr>
         <tr>
            <td colspan="1" align="center">
                <asp:HiddenField runat="server" ID="HFD_Goods_Id_4" />
                <asp:CheckBox ID="cbxContribute_IsStock_4" runat="server" OnClick="cbxContribute_IsStock_OnClick('4')"/>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Name_4" runat="server"></asp:TextBox>&nbsp;<a href onclick="WindowsOpen(4)" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Qty_4" runat="server"></asp:TextBox>
            </td>
            <td align="center">
                <asp:TextBox ID="tbxGoods_Unit_4" runat="server"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Amt_4" runat="server" Text="0"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_DueDate_4" runat="server" Width="90px" onchange="CheckDateFormat(this, '物品保存期限 ');"></asp:TextBox>
                <img id="imgGoods_DueDate_4" alt="" src="../images/date.gif" />
            </td>
            <td colspan="2" align="center">
                <asp:TextBox ID="tbxGoods_Comment_4" runat="server" Width="160px" MaxLength="100" ></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <img border="0" src="../images/toolbar_cancel.gif" width="20" onClick="Contribute_Cancel_OnClick('4')" style="cursor:hand">
            </td>
         </tr>
         <tr>
            <td colspan="1" align="center">
                <asp:HiddenField runat="server" ID="HFD_Goods_Id_5" />
                <asp:CheckBox ID="cbxContribute_IsStock_5" runat="server" OnClick="cbxContribute_IsStock_OnClick('5')"/>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Name_5" runat="server"></asp:TextBox>&nbsp;<a href onclick="WindowsOpen(5)" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Qty_5" runat="server"></asp:TextBox>
            </td>
            <td align="center">
                <asp:TextBox ID="tbxGoods_Unit_5" runat="server"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_Amt_5" runat="server" Text="0"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGoods_DueDate_5" runat="server" Width="90px" onchange="CheckDateFormat(this, '物品保存期限 ');"></asp:TextBox>
                <img id="imgGoods_DueDate_5" alt="" src="../images/date.gif" />
            </td>
            <td colspan="2" align="center">
                <asp:TextBox ID="tbxGoods_Comment_5" runat="server" Width="160px" MaxLength="100" ></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <img border="0" src="../images/toolbar_cancel.gif" width="20" onClick="Contribute_Cancel_OnClick('5')" style="cursor:hand">
            </td>
         </tr>
    </table>
    </div>
    <div class="function">
        <asp:Button ID="btnAdd" class="npoButton npoButton_New" runat="server" 
            Text="存檔" onclick="btnAdd_Click" OnClientClick= "return CheckFieldMustFillBasic(); "/>
        <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
            Text="取消" onclick="btnExit_Click" />
    </div>
    </form>
</body>
</html>

