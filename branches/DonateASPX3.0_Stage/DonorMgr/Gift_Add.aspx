<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Gift_Add.aspx.cs" Inherits="DonorMgr_Gift_Add" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>公關贈品輸入【新增】</title>
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

            // 2014/8/27 增加全選的功能
            $("#checkboxAll").click(function () {

                if ($("#checkboxAll").prop("checked")) {
                    $("input[id=checkbox]").prop("checked", true);
                }
                else {
                    $("input[id=checkbox]").prop("checked", false);
                }
            });
        });
        
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxGift_Date",   // id of the input field
                button: "imgContribute_Date"     // 與觸發動作的物件ID相同
            });
        }
        function Contribute_Cancel_OnClick(i) {
            var Gifts_Id = document.getElementById('HFD_Gifts_Id_' + i);
            var Gifts_Name = document.getElementById('tbxGifts_Name_' + i);
            var Gifts_Qty = document.getElementById('tbxGifts_Qty_' + i);
            var Gifts_Comment = document.getElementById('tbxGifts_Comment_' + i);
            if (Gifts_Name.value != '') {
                if (confirm('您是否確定要清除『 ' + Gifts_Name.value + ' 』？')) {
                    Gifts_Id.value = '';
                    Gifts_Name.value = '';
                    Gifts_Qty.value = '';
                    Gifts_Comment.value = '';
                }
            } else {
                if (confirm('您是否確定要清除？')) {
                    Gifts_Id.value = '';
                    Gifts_Name.value = '';
                    Gifts_Qty.value = '';
                    Gifts_Comment.value = '';
                }
            }
        }
        function WindowsOpen(i) {
            window.open('GiftShow.aspx?Gifts_Id=HFD_Gifts_Id_' + i + '&Gifts_Name=tbxGifts_Name_' + i , 'NewWindows',
                        'status=no,scrollbars=yes,top=100,left=120,width=500,height=450;');
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
    <asp:HiddenField runat="server" ID="HFD_Gift_Id" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        公關贈品輸入【新增】
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
        </tr>
        <tr>
            <td colspan="4">

            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxGift_Date" runat="server" onchange="CheckDateFormat(this, '日期');"></asp:TextBox>
                <img id="imgContribute_Date" alt="" src="../images/date.gif" />
            </td>
        </tr> 
        <tr>
            <td colspan="4">

            </td>
        </tr>
        <tr>
            <th colspan="1" align="right">
                備註：
            </th>
            <td colspan="3">
                <asp:Textbox runat="server" ID="tbxComment" CssClass="font9" 
                    TextMode="MultiLine" Width="450px" Height="100px"></asp:Textbox> 
            </td>
         </tr>
    </table>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
        <tr>
            <td align="center" width="100%" colspan="5">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    <div class="function">
        <asp:Button ID="btnAdd" class="npoButton npoButton_New" runat="server" 
            Text="存檔" onclick="btnAdd_Click" />
        <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
            Text="取消" onclick="btnExit_Click" />
    </div>
    </form>
</body>
</html>


