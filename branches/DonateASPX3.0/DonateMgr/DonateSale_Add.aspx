<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateSale_Add.aspx.cs" Inherits="DonateMgr_DonateSaleQry_Add" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>收據文宣內容【新增】 </title>
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
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxSale_BeginDate",   // id of the input field
                button: "imgSale_BeginDate"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxSale_EndDate",   // id of the input field
                button: "imgSale_EndDate"     // 與觸發動作的物件ID相同
            });
        }
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var cnt = 0;
            var sName = "";
            var ddlDept = document.getElementById('ddlDept');
            var tbxSale_Subject = document.getElementById('tbxSale_Subject');
            var tbxSale_BeginDate = document.getElementById('tbxSale_BeginDate');
            var tbxSale_EndDate = document.getElementById('tbxSale_EndDate');
            var tbxSale_Content = document.getElementById('tbxSale_Content');
            if (ddlDept.value == "") {
                strRet += "機構 ";
            }
            if (tbxSale_Subject.value == "") {
                strRet += "文宣標題 ";
            }
            if (tbxSale_BeginDate.value == "") {
                strRet += "列印起日 ";
            }
            if (tbxSale_Content.value == "") {
                strRet += "文宣內容 ";
            }
            if (tbxSale_Subject.value == "" || tbxSale_BeginDate.value == "" || tbxSale_Content.value == "") {
                strRet += "欄位不可為空白！";
            }
            if (strRet != "") {
                alert(strRet)
                return false;
            }
            cnt = 0;
            sName = tbxSale_Subject.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 50) {
                alert('文宣標題  欄位長度超過限制！');
                return false;
            }
            var begindate = new Date(tbxSale_BeginDate.value);
            var enddate = new Date(tbxSale_EndDate.value);
            var diffdate = (Date.parse(begindate.toString()) - Date.parse(enddate.toString())) / (1000 * 60 * 60 * 24);
            if (parseInt(diffdate) > 0) {
                alert('刊登起日不可大於迄日！');
                tbxSale_BeginDate.focus();
                return false;
            }
            if (tbxSale_EndDate.value == "") {
                tbxSale_EndDate.value = "2099/12/31 ";
            }
            return true;
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
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" /> 
        收據文宣內容【新增】 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                文宣標題：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxSale_Subject" runat="server"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                列印期間：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxSale_BeginDate" runat="server" onchange="CheckDateFormat(this, ' 列印日期');"></asp:TextBox>
                <img id="imgSale_BeginDate" alt="" src="../images/date.gif" />~
                <asp:TextBox ID="tbxSale_EndDate" runat="server" onchange="CheckDateFormat(this, ' 列印日期');"></asp:TextBox>
                <img id="imgSale_EndDate" alt="" src="../images/date.gif" />
            </td>
        </tr>
        <tr>
            <th colspan="1" align="right">
                文宣內容：
            </th>
            <td colspan="7">
                <asp:Textbox runat="server" ID="tbxSale_Content" CssClass="font9" 
                    TextMode="MultiLine" Width="700px" Height="100px"></asp:Textbox> 
            </td>
         </tr>
    </table>
    </div>
    <div class="function">
        <asp:Button ID="btnAdd" class="npoButton npoButton_New" runat="server" OnClientClick= "return CheckFieldMustFillBasic(); "
            Text="存檔" onclick="btnAdd_Click"/>
        <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
            Text="取消" onclick="btnExit_Click"/>
    </div>
    </form>
</body>
</html>
