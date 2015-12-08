<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Act_Edit.aspx.cs" Inherits="DonateMgr_Act_Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title> 專案活動【修改】</title>
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
                inputField: "tbxAct_BeginDate",   // id of the input field
                button: "imgAct_BeginDate"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxAct_EndDate",   // id of the input field
                button: "imgAct_EndDate"     // 與觸發動作的物件ID相同
            });
        }
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var cnt = 0;
            var sName = "";
            var tbxAct_Name = document.getElementById('tbxAct_Name');
            var tbxAct_ShortName = document.getElementById('tbxAct_ShortName');
            var tbxAct_BeginDate = document.getElementById('tbxAct_BeginDate');
            var tbxAct_EndDate = document.getElementById('tbxAct_EndDate');
            var tbxAct_OrgName = document.getElementById('tbxAct_OrgName');
            var tbxAct_OrgName2 = document.getElementById('tbxAct_OrgName2');
            var tbxAct_Subject = document.getElementById('tbxAct_Subject');
            var tbxAct_Licence = document.getElementById('tbxAct_Licence');
            if (tbxAct_Name.value == "") {
                strRet += "活動名稱 ";
            }
            if (tbxAct_ShortName.value == "") {
                strRet += "活動簡稱 ";
            }
            if (tbxAct_BeginDate.value == "") {
                strRet += "活動起日 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet)
                return false;
            }
            cnt = 0;
            sName = tbxAct_Name.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 60) {
                alert('活動名稱 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxAct_ShortName.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 60) {
                alert('活動簡稱 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxAct_OrgName.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 60) {
                alert('主辦單位 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            var sName = tbxAct_OrgName2.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 60) {
                alert('協辦單位 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            var sName = tbxAct_Subject.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 60) {
                alert('活動主題 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            var sName = tbxAct_Licence.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 80) {
                alert('勸募許可文號 欄位長度超過限制！');
                return false;
            }

            else {
                if (tbxAct_EndDate.value == "") {
                    tbxAct_EndDate.value = "2099/12/31 ";
                }
                return true;
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
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" /> 
        專案活動【修改】 
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
                活動名稱：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxAct_Name" runat="server"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                活動簡稱：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxAct_ShortName" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                主辦單位：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxAct_OrgName" runat="server"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                協辦單位：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxAct_OrgName2" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                活動主題：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxAct_Subject" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                活動期間：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxAct_BeginDate" runat="server" onchange="CheckDateFormat(this, '開始活動期間');"></asp:TextBox>
                <img id="imgAct_BeginDate" alt="" src="../images/date.gif" />~
                <asp:TextBox ID="tbxAct_EndDate" runat="server" onchange="CheckDateFormat(this, '結束活動期間');"></asp:TextBox>
                <img id="imgAct_EndDate" alt="" src="../images/date.gif" />
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                勸募許可文號：
            </th>
            <td align="left" colspan="5">
                <asp:TextBox runat="server" ID="tbxAct_Licence" CssClass="font9"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th colspan="1" align="right">
                備註：<br />
            </th>
            <td colspan="5">
                <asp:Textbox runat="server" ID="tbxRemark" CssClass="font9" 
                    TextMode="MultiLine" Width="450px" Height="100px"></asp:Textbox> 
            </td>
         </tr>
    </table>
    </div>
    <div class="function">
        <asp:Button ID="btnEdit" class="npoButton npoButton_Modify" runat="server" 
            Text="修改" OnClientClick= "return CheckFieldMustFillBasic();"
            onclick="btnEdit_Click" />
        <asp:Button ID="btnDel" class="npoButton npoButton_Del" runat="server" 
            Text="刪除" OnClientClick="return confirm('您是否確定要刪除？')"
             onclick="btnDel_Click" />
        <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
            Text="取消" onclick="btnExit_Click" />
    </div>
    </form>
</body>
</html>

