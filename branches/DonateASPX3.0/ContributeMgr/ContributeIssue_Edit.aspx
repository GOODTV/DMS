<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ContributeIssue_Edit.aspx.cs" Inherits="ContributeMgr_ContributeIssue_Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>實物奉獻領用作業【修改】</title>
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
            var tbxIssue_No = document.getElementById('tbxIssue_No');
            tbxIssue_No.readOnly = true;
            tbxIssue_No.style.backgroundColor = '#FFE1AF';
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
                inputField: "tbxIssue_Date",   // id of the input field
                button: "imgIssue_Date"     // 與觸發動作的物件ID相同
            });
        }
        function cbxIssue_Pre_OnClick() {
            var cbxIssue_Pre = document.getElementById('cbxIssue_Pre');
            var tbxIssue_No = document.getElementById('tbxIssue_No');
            var HFD_Issue_No = document.getElementById('HFD_Issue_No');
            if (cbxIssue_Pre.checked == true) {
                tbxIssue_No.style.backgroundColor = '#ffffff';
                tbxIssue_No.readOnly = false;
                tbxIssue_No.value = '';
                tbxIssue_No.focus();
            } else {
                tbxIssue_No.style.backgroundColor = '#FFE1AF';
                tbxIssue_No.readOnly = true;
                tbxIssue_No.value = HFD_Issue_No.value;
            }
        }
        function Contribute_Cancel_OnClick(i) {
            var Goods_Id = document.getElementById('HFD_Goods_Id_' + i);
            var Goods_Name = document.getElementById('tbxGoods_Name_' + i);
            var Goods_Qty = document.getElementById('tbxGoods_Qty_' + i);
            var Goods_Unit = document.getElementById('tbxGoods_Unit_' + i);
            var Goods_Comment = document.getElementById('tbxGoods_Comment_' + i);
            if (Goods_Name.value != '') {
                if (confirm('您是否確定要清除『 ' + Goods_Name.value + ' 』？')) {
                    Goods_Id.value = '';
                    Goods_Name.value = '';
                    Goods_Qty.value = '';
                    Goods_Unit.value = '';
                    Goods_Comment.value = '';
                }
            } else {
                if (confirm('您是否確定要清除？')) {
                    Goods_Id.value = '';
                    Goods_Name.value = '';
                    Goods_Qty.value = '';
                    Goods_Unit.value = '';
                    Goods_Comment.value = '';
                }
            }
        }
        function cbxContribute_IsStock_OnClick(i) {
            var Goods_Id = document.getElementById('HFD_Goods_Id_' + i);
            var Goods_Name = document.getElementById('tbxGoods_Name_' + i);
            var Goods_Qty = document.getElementById('tbxGoods_Qty_' + i);
            var Goods_Unit = document.getElementById('tbxGoods_Unit_' + i);
            var Goods_Comment = document.getElementById('tbxGoods_Comment_' + i);
            var cbxContribute_IsStock = document.getElementById('cbxContribute_IsStock_' + i);
            if (cbxContribute_IsStock.checked == false) {
                Goods_Id.value = '';
                Goods_Name.value = '';
                Goods_Qty.value = '';
                Goods_Unit.value = '';
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
            var cnt = 0;
            var sName = '';
            var tbxIssue_Processor = document.getElementById('tbxIssue_Processor');
            var tbxIssue_Date = document.getElementById('tbxIssue_Date');
            var ddlIssue_Purpose = document.getElementById('ddlIssue_Purpose');
            var tbxIssue_Org = document.getElementById('tbxIssue_Org');
            var tbxIssue_Comment = document.getElementById('tbxIssue_Comment');
            var cbxIssue_Pre = document.getElementById('cbxIssue_Pre');
            var tbxIssue_No = document.getElementById('tbxIssue_No');
            if (tbxIssue_Processor.value == '') {
                alert('領取人 欄位不可為空白！');
                tbxIssue_Processor.focus();
                return false;
            }
            cnt = 0;
            sName = tbxIssue_Processor.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('領取人 欄位長度超過限制！');
                return false;
            }
            if (tbxIssue_Date.value == '') {
                alert('領用日期 欄位不可為空白！');
                tbxIssue_Date.focus();
                return false;
            }
            if (ddlIssue_Purpose.value == '') {
                alert('領用用途 欄位不可為空白！');
                ddlIssue_Purpose.focus();
                return false;
            }
            cnt = 0;
            sName = ddlIssue_Purpose.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('領用用途 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxIssue_Org.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('出貨單位 欄位長度超過限制！');
                return false;
            }

            cnt = 0;
            var sName = ddlIssue_Purpose.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 100) {
                alert('備註 欄位長度超過限制！');
                return false;
            }
            if (cbxIssue_Pre.checked) {
                if (tbxIssue_No.value == '') {
                    alert('收據編號 欄位不可為空白！');
                    tbxIssue_No.focus();
                    return false;
                }
                cnt = 0;
                sName = tbxIssue_No.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 20) {
                    alert('收據編號 欄位長度超過限制！');
                    return false;
                }
            }

            for (i = 1; i <= 5; i++) {
                var Contribute_IsStock = document.getElementById('cbxContribute_IsStock_' + i);
                var Goods_Id = document.getElementById('HFD_Goods_Id_' + i);
                var Goods_Name = document.getElementById('tbxGoods_Name_' + i);
                var Goods_Qty = document.getElementById('tbxGoods_Qty_' + i);
                var Goods_Unit = document.getElementById('tbxGoods_Unit_' + i);
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
            return true;
        }
        function GetFormatDate(InputValue) {
            if (InputValue < 10) {
                InputValue = '0' + InputValue;
            }

            return InputValue;
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
    <asp:HiddenField runat="server" ID="HFD_Issue_No" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        實物奉獻領用作業【修改】
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9" 
                    AutoPostBack="True"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                領取人：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIssue_Processor" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                領用日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxIssue_Date" runat="server" onchange="CheckDateFormat(this, '領用日期');"></asp:TextBox>
                <img id="imgIssue_Date" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                領用用途：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlIssue_Purpose" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                出貨單位：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIssue_Org" CssClass="font9"></asp:TextBox>
            </td>
            <th colspan="1" align="right">
                備註：
            </th>
            <td colspan="2">
                <asp:Textbox runat="server" ID="tbxComment" CssClass="font9" Width="200px" ></asp:Textbox> 
            </td>
            <td align="center" colspan="1">
                <asp:checkbox runat="server" ID="cbxIssue_Pre" Text="手開收據" CssClass="font9" OnClick="cbxIssue_Pre_OnClick()" ></asp:checkbox>
            </td>
            <th align="right" colspan="1">
                領用編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIssue_No" CssClass="font9"
                   ></asp:TextBox>
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
        <tr>
            <th colspan="1" align="center">
                庫存品
            </th>
            <th colspan="1" align="center">
                物品名稱
            </th>
            <th colspan="1" align="center">
                數量
            </th>
            <th colspan="1" align="center">
                單位
            </th>
            <th colspan="3" align="center">
                備註
            </th>
            <th colspan="1" align="center">
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
            <td colspan="3" align="center">
                <asp:TextBox ID="tbxGoods_Comment_1" runat="server" Width="250px" MaxLength="100" ></asp:TextBox>
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
            <td colspan="3" align="center">
                <asp:TextBox ID="tbxGoods_Comment_2" runat="server" Width="250px" MaxLength="100" ></asp:TextBox>
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
            <td colspan="3" align="center">
                <asp:TextBox ID="tbxGoods_Comment_3" runat="server" Width="250px" MaxLength="100" ></asp:TextBox>
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
            <td colspan="3" align="center">
                <asp:TextBox ID="tbxGoods_Comment_4" runat="server" Width="250px" MaxLength="100" ></asp:TextBox>
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
            <td colspan="3" align="center">
                <asp:TextBox ID="tbxGoods_Comment_5" runat="server" Width="250px" MaxLength="100" ></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <img border="0" src="../images/toolbar_cancel.gif" width="20" onClick="Contribute_Cancel_OnClick('5')" style="cursor:hand">
            </td>
         </tr>
    </table>
    </div>
    <div class="function">
        <br />
        <asp:Button ID="btnEdit2" class="npoButton npoButton_Modify" runat="server" 
            Text="修改" Width="89px" onclick="btnEdit_Click" OnClientClick= "return CheckFieldMustFillBasic(); "/>
        <asp:Button ID="btnDel" class="npoButton npoButton_Del" runat="server" 
            Text="刪除" Width="89px" onclick="btnDel_Click" OnClientClick= "return confirm('您是否確定要刪除？'); "/>
        <asp:Button ID="btnExit2" class="npoButton npoButton_Cancel" runat="server" 
            Text="取消" Width="80px" onclick="btnExit_Click" />
    </div>
    </form>
</body>
</html>


