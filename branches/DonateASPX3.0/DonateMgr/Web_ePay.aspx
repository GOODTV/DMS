<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Web_ePay.aspx.cs" Inherits="DonateMgr_Web_ePay" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
    <title>>線上付款方式確認</title>
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

            $("#checkboxAll").click(function () {

                if ($("#checkboxAll").prop("checked")) {
                    $("input[id=checkbox]").prop("checked", true);
                }
                else {
                    $("input[id=checkbox]").prop("checked", false);
                }

                CheckboxCount();

            });


            $("input[id=checkbox]").click(function () {

                CheckboxCount();

            });

            $("select[id^='ddl_']").change(function (e) {

                if (confirm('確定要修改嗎？')) {
                    var orderid = this.id.replace('ddl_', '');
                    var paytype = this.value;
                    //alert('orderid=' + orderid + ',paytype=' + paytype);

                    var result = '';
                    $.ajax({
                        type: 'post',
                        url: "../common/ajax.aspx",
                        async: false,
                        data: "Type=20&orderid=" + orderid + "&paytype=" + paytype,
                        success: function (data) {
                            if (data == 'Y') {
                                result = data;
                            }
                        },
                        error: function () {
                            //alert('ajax failed');
                        }

                    });

                    if (result == 'Y') {
                        $('#btnQuery').click();
                    }
                    else {
                        this.value = '99';
                        alert('修改失敗！');
                    }
                }
                else
                    this.value = '99';

            });


        });


        function CheckboxCount() {

            var Order_cnt = 0;
            var Order_amt = 0;
            $("input[id=checkbox]:checked").each(function (i) { //判斷checkbox勾選了幾個
                Order_cnt++;
                Order_amt += parseInt($(this).val());
            });
            //alert(Order_amt);

            $('#lblCheckCnt').html("<font color=blue>已點選 " + thousandComma(Order_cnt) + " 筆 / 捐款金額合計：" + thousandComma(Order_amt) + " 元</font>");

        }
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "txtDonateDateS",   // id of the input field
                button: "imgDonateDateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtDonateDateE",   // id of the input field
                button: "imgDonateDateE"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtDonateDate",   // id of the input field
                button: "imgDonateDate"     // 與觸發動作的物件ID相同
            });
        }

        function thousandComma(number) {
            var num = number.toString();
            var pattern = /(-?\d+)(\d{3})/;

            while (pattern.test(num)) {
                num = num.replace(pattern, "$1,$2");

            }
            return num;

        }

        function Print(PrintType) {
            window.open('../Web_ePay_Print.aspx', 'Web_ePay_Print', 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
        }


        function Confirm_Data() {

            var Total_Row = "";
            var Order_cnt = 0;
            var paytype99_cnt = 0;

            $("input[id=checkbox]:checked").each(function () { //判斷checkbox勾選了幾個
                Order_cnt++;
                if ($(this).parent().parent().find('select[id^=ddl_]').length > 0) {
                    paytype99_cnt++;
                }
                
            });

            if (paytype99_cnt > 0) {
                if (!confirm("要轉收據的資料內有非信用卡的付款方式，\n建議先修改成與銀行的付款方式一樣。\n請按『取消』再做修改？\n或請按『確定』繼續轉收據？")) {
                    return false;
                }

            }

            if (Order_cnt > 0) {
                if (!confirm("已選擇 " + Order_cnt + " 筆，\n您是否確定要執行？")) {
                    return false;
                }

            }
            else
            {
                alert('未勾選資料！');
                return false;
            }

            if ($("#txtDonateDate").val() == "") {
                alert("轉收據的捐款日期欄位不可為空白！");
                return false;
            }

        }
 

    </script>
</head>
<body class="body">
    <form id="Form1" runat="server">
        <div id="menucontrol">
            <a href="#">
                <img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
            <a href="#">
                <img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        </div>
        <div id="container">
            <h1 style="padding-bottom: 0px;">
                <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
                線上付款方式確認 
            </h1>
            <table width="100%" border="0" align="left" cellpadding="0" cellspacing="1" class="table_v">
                <tr>
                    <!--<th align="right" colspan="1">捐款狀態：
                    </th>
                    <td align="left" colspan="1">
                        <asp:DropDownList ID="Status" runat="server" class="font9">
                            <asp:ListItem Value="">請選擇</asp:ListItem>
                            <asp:ListItem Value="0">捐款成功</asp:ListItem>
                            <asp:ListItem Value="1">等待捐款中</asp:ListItem>
                        </asp:DropDownList>
                    </td>-->
                    <th align="right" colspan="1">線上捐款日期：
                    </th>
                    <td align="left" colspan="2">
                        <asp:TextBox ID="txtDonateDateS" runat="server"
                            onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                        <img id="imgDonateDateS" alt="" src="../images/date.gif" />
                        ~
                <asp:TextBox ID="txtDonateDateE" runat="server"
                    onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                        <img id="imgDonateDateE" alt="" src="../images/date.gif" />
                    </td>
                    <th align="right" colspan="1">轉收據的捐款日期：
                    </th>
                    <td align="left" colspan="1">
                        <asp:TextBox ID="txtDonateDate" runat="server"
                            onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                        <img id="imgDonateDate" alt="" src="../images/date.gif" />
                        </td>
                    <td align="right" colspan="3" class="style1">
                        <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server" Width="80px"
                            Text="查詢" OnClick="btnQuery_Click" />
                        <asp:Button ID="btnInput" runat="server"  Width="150px"
                            Text="轉捐款記錄(收據)" CssClass=" npoButton npoButton_Export" 
                           OnClientClick="return Confirm_Data()" onclick="btnInput_Click" Enabled="False" />
                                     <!--  OnClientClick="return CheckFieldMustFillBasic(); "
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm" Text="列表"
                 OnClientClick="if (confirm('您是否確定要將查詢結果匯出？')==false) {return false;} Print('');"/>
                <asp:Button ID="btnToxls" runat="server"  Width="20mm"
                    Text="匯出" OnClick="btnToxls_Click" CssClass="npoButton npoButton_Excel" OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');"/>
                --></td>
                </tr>
                <tr>
                    <td align="left" colspan="2">
                        <asp:Label id="lblQueryCnt" name="lblQueryCnt" runat="server"></asp:Label></td>
                    <td align="left" colspan="6">
                        <asp:Label id="lblCheckCnt" name="lblCheckCnt" runat="server">&nbsp</asp:Label></td>
                </tr>
                <tr>
                    <td align="center" width="100%" colspan="8">
                        <asp:Label ID="lblGridList" runat="server"></asp:Label>
                    </td>
                </tr>
            </table>

        </div>
    </form>
</body>
</html>
