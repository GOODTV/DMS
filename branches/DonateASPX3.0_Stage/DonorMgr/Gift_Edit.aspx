<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Gift_Edit.aspx.cs" Inherits="DonorMgr_Gift_Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>公關贈品輸入【修改】</title>
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
            for (i = 1; i <= 5; i++) {
                var Gifts_Name = document.getElementById('tbxGifts_Name_' + i);
                Gifts_Name.readOnly = true;
                Gifts_Name.style.backgroundColor = '#FFE1AF';
            }
        });
    </script>
    <script type="text/javascript">
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
            window.open('GiftShow.aspx?Gifts_Id=HFD_Gifts_Id_' + i + '&Gifts_Name=tbxGifts_Name_' + i, 'NewWindows',
                        'status=no,scrollbars=yes,top=100,left=120,width=500,height=450;');
        }
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var tbxGift_Date = document.getElementById('tbxGift_Date');
            if (tbxGift_Date.value == "") {
                strRet += "日期 ";
            }
            //201404025 修改 by Ian_Kao
//            if ($("#tbxGifts_Name_1").val() == "" && $("#tbxGifts_Name_2").val() == "" && $("#tbxGifts_Name_3").val() == "" && $("#tbxGifts_Name_4").val() == "" && $("#tbxGifts_Name_5").val() == "") {
//                strRet += "物品名稱 ";
//            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet)
                return false;
            }
            for (i = 1; i <= 5; i++) {
                var Gifts_Id = document.getElementById('HFD_Gifts_Id_' + i);
                var Gifts_Name = document.getElementById('tbxGifts_Name_' + i);
                var Gifts_Qty = document.getElementById('tbxGifts_Qty_' + i);
                var Gifts_Comment = document.getElementById('tbxGifts_Comment_' + i);
                if (Gifts_Name.value != '') {
                    if (Gifts_Id.value == '') {
                        alert('您輸入的『' + Gifts_Name.value + '』該商品代號不存在！');
                        return false;
                    }
                }
//                if (Gifts_Name.value != '' && Gifts_Qty.value == '') {
//                    alert('『' + Gifts_Name.value + '』數量  欄位不可為空白！');
//                    Gifts_Qty.focus();
//                    return false;
//                }
                if (isNaN(Number(Gifts_Qty.value)) == true) {
                    alert('『' + Gifts_Name.value + '』數量  欄位必須為數字！');
                    Gifts_Qty.focus();
                    return false;
                }
                cnt = 0;
                sName = Gifts_Qty.value
                for (var j = 0; j < sName.length; j++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 7) {
                    alert('『' + Gifts_Name.value + '』數量 欄位長度超過限制！');
                    return false;
                }
                if (Gifts_Comment.value != '') {
                    cnt = 0;
                    sName = Gifts_Comment.value
                    for (var j = 0; j < sName.length; j++) {
                        if (escape(sName.charAt(j)).length >= 4) cnt += 2;
                        else cnt++;
                    }
                    if (cnt > 100) {
                        alert('『' + Gifts_Name_.value + '』備註 欄位長度超過限制！');
                        return false;
                    }
                }
                for (j = i + 1; j < 5; j++) {
                    if (document.getElementById('HFD_Gifts_Id_' + i).value == document.getElementById('HFD_Gifts_Id_' + j).value && document.getElementById('tbxGifts_Name_' + i).value == document.getElementById('tbxGifts_Name_' + j).value && document.getElementById('HFD_Gifts_Id_' + i).value != '') {
                        alert('『' + document.getElementById('tbxGifts_Name_' + j).value + '』物品名稱重覆出現！');
                        return false;
                    }
                }
                return true;
            }
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
    <asp:HiddenField runat="server" ID="HFD_Donor_Id" />
    <asp:HiddenField runat="server" ID="HFD_Invoice_No" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        公關贈品輸入【修改】
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
           <%-- <th align="right"colspan="1" >
                類別：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxCategory" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>--%>
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
        <%--<tr>
            <th align="right" colspan="1">
                收據地址：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxAddress" CssClass="font9"  Width="250px" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>--%>
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
         <tr>
            <th colspan="1" align="right">
                贈送內容：
            </th>
            <td  align="center" colspan="3">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
         <tr>
            <th colspan="1" align="center">
                名稱
            </th>
            <th colspan="1" align="center">
                數量
            </th>
            <th colspan="1" align="center" width="15%">
                備註
            </th>
            <th colspan="1" width="6%" align="center">
                清除
            </th>
         </tr>
         <tr>
            <td colspan="1" align="center">
                <asp:HiddenField runat="server" ID="HFD_Gifts_Id_1"/>
                <asp:TextBox ID="tbxGifts_Name_1" runat="server"></asp:TextBox>&nbsp;<a href onclick="WindowsOpen(1)" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGifts_Qty_1" runat="server"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGifts_Comment_1" runat="server" Width="160px" MaxLength="100" ></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <img border="0" src="../images/toolbar_cancel.gif" width="20" onClick="Contribute_Cancel_OnClick('1')" style="cursor:hand">
            </td>
         </tr>
         <tr>
            <td colspan="1" align="center">
                <asp:HiddenField runat="server" ID="HFD_Gifts_Id_2" />
                <asp:TextBox ID="tbxGifts_Name_2" runat="server"></asp:TextBox>&nbsp;<a href onclick="WindowsOpen(2)" style="cursor:hand" ><img border="0" src="../images/toolbar_search.gif" width="17"></a>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGifts_Qty_2" runat="server"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGifts_Comment_2" runat="server" Width="160px" MaxLength="100" ></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <img border="0" src="../images/toolbar_cancel.gif" width="20" onClick="Contribute_Cancel_OnClick('2')" style="cursor:hand">
            </td>
         </tr>
         <tr>
            <td colspan="1" align="center">
                <asp:HiddenField runat="server" ID="HFD_Gifts_Id_3" />
                <asp:TextBox ID="tbxGifts_Name_3" runat="server"></asp:TextBox>&nbsp;<a href onclick="WindowsOpen(3)" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGifts_Qty_3" runat="server"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGifts_Comment_3" runat="server" Width="160px" MaxLength="100" ></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <img border="0" src="../images/toolbar_cancel.gif" width="20" onClick="Contribute_Cancel_OnClick('3')" style="cursor:hand">
            </td>
         </tr>
         <tr>
            <td colspan="1" align="center">
                <asp:HiddenField runat="server" ID="HFD_Gifts_Id_4" />
                <asp:TextBox ID="tbxGifts_Name_4" runat="server"></asp:TextBox>&nbsp;<a href onclick="WindowsOpen(4)" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGifts_Qty_4" runat="server"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGifts_Comment_4" runat="server" Width="160px" MaxLength="100" ></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <img border="0" src="../images/toolbar_cancel.gif" width="20" onClick="Contribute_Cancel_OnClick('4')" style="cursor:hand">
            </td>
         </tr>
         <tr>
            <td colspan="1" align="center">
                <asp:HiddenField runat="server" ID="HFD_Gifts_Id_5" />
                <asp:TextBox ID="tbxGifts_Name_5" runat="server"></asp:TextBox>&nbsp;<a href onclick="WindowsOpen(5)" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGifts_Qty_5" runat="server"></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <asp:TextBox ID="tbxGifts_Comment_5" runat="server" Width="160px" MaxLength="100" ></asp:TextBox>
            </td>
            <td colspan="1" align="center">
                <img border="0" src="../images/toolbar_cancel.gif" width="20" onClick="Contribute_Cancel_OnClick('5')" style="cursor:hand">
            </td>
         </tr>
    </table>
    </div>
    <div class="function">
        <asp:Button ID="btnEdit" class="npoButton npoButton_New" runat="server" 
            Text="存檔" onclick="btnEdit_Click" OnClientClick= "return CheckFieldMustFillBasic(); "/>
        <asp:Button ID="btnDel" class="npoButton npoButton_Del" runat="server" 
            Text="刪除" OnClientClick="return confirm('您確定要刪除嗎？')" onclick="btnDel_Click" />
        <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
            Text="取消" onclick="btnExit_Click" />
    </div>
    </form>
</body>
</html>


