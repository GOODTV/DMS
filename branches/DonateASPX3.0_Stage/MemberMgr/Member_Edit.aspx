<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Member_Edit.aspx.cs" Inherits="MemberMgr_Member_Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>讀者資料維護</title>
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
                inputField: "tbxBirthday",   // id of the input field
                button: "imgBirthday"     // 與觸發動作的物件ID相同
            });
        }
        function CBL_SingleChoice(sender) {
            var container = sender.parentNode;
            if (container.tagName.toUpperCase() == "TD") { // table 布局，否則為span布局
                container = container.parentNode.parentNode; // 層次: <table><tr><td><input />
            }
            var chkList = container.getElementsByTagName("input");
            var senderState = sender.checked;
            for (var i = 0; i < chkList.length; i++) {
                chkList[i].checked = false;
            }
            sender.checked = senderState;
        }
        function ReadOnly() {
            var tbxIntroducer_Name = document.getElementById('tbxIntroducer_Name');
            tbxIntroducer_Name.readOnly = true;
            tbxIntroducer_Name.style.backgroundColor = '#FFE1AF';
        }
        function WindowsOpen(i) {
            window.open('MemberShow.aspx?Donor_Id=HFD_Donor_Id' + '&Donor_Name=tbxIntroducer_Name', 'NewWindows',
                        'status=no,scrollbars=yes,top=100,left=120,width=600,height=450);');
        }
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var cnt = 0;
            var sName = "";

            var tbxDonor_Name = document.getElementById('tbxDonor_Name');
            var ddlSex = document.getElementById('ddlSex');
            var ddlTitle = document.getElementById('ddlTitle');
            var cblDonor_Type = document.getElementById('cblDonor_Type');
            var ddlMember_Status = document.getElementById('ddlMember_Status');
            var tbxIDNo = document.getElementById('tbxIDNo');
            //var tbxReligionName = document.getElementById('tbxReligionName');
            var tbxCellular_Phone = document.getElementById('tbxCellular_Phone');
            var tbxTel_Office = document.getElementById('tbxTel_Office');
            var tbxTel_Office_Loc = document.getElementById('tbxTel_Office_Loc');
            var tbxTel_Office_Ext = document.getElementById('tbxTel_Office_Ext');
            //var tbxTel_Home = document.getElementById('tbxTel_Home');
            var tbxFax = document.getElementById('tbxFax');
            var tbxFax_Loc = document.getElementById('tbxFax_Loc');
            var tbxEMail = document.getElementById('tbxEMail');
            var tbxContactor = document.getElementById('tbxContactor');
            //var tbxOrgName = document.getElementById('tbxOrgName');
            //var tbxJobTitle = document.getElementById('tbxJobTitle');
            var cbxIsLocal = document.getElementById('cbxIsLocal');
            var ddlCity = document.getElementById('ddlCity');
            var ddlArea = document.getElementById('ddlArea');
            var tbxStreet = document.getElementById('tbxStreet');
            var tbxLane = document.getElementById('tbxLane');
            var tbxAlley = document.getElementById('tbxAlley');
            var tbxNo1 = document.getElementById('tbxNo1');
            var tbxNo2 = document.getElementById('tbxNo2');
            var tbxFloor1 = document.getElementById('tbxFloor1');
            var tbxFloor2 = document.getElementById('tbxFloor2');
            var tbxRoom = document.getElementById('tbxRoom');
            var cbxIsAbroad = document.getElementById('cbxIsAbroad');
            var tbxOverseasCountry = document.getElementById('tbxOverseasCountry');
            var tbxOverseasAddress = document.getElementById('tbxOverseasAddress');
            var tbxInvoice_Street = document.getElementById('tbxInvoice_Street');
            var tbxInvoice_Lane = document.getElementById('tbxInvoice_Lane');
            var tbxInvoice_Alley0 = document.getElementById('tbxInvoice_Alley0');
            var tbxInvoice_No1 = document.getElementById('tbxInvoice_No1');
            var tbxInvoice_No2 = document.getElementById('tbxInvoice_No2');
            var tbxInvoice_Floor1 = document.getElementById('tbxInvoice_Floor1');
            var tbxInvoice_Floor2 = document.getElementById('tbxInvoice_Floor2');
            var tbxInvoice_Room = document.getElementById('tbxInvoice_Room');
            var tbxInvoice_OverseasCountry = document.getElementById('tbxInvoice_OverseasCountry');
            var tbxInvoice_OverseasAddress = document.getElementById('tbxInvoice_OverseasAddress');
            var tbxInvoice_Title = document.getElementById('tbxInvoice_Title');
            var tbxIsSendNewsNum = document.getElementById('tbxIsSendNewsNum');
            if (tbxDonor_Name.value == '') {
                strRet += "姓名 ";
            }
            /*if (cblDonor_Type.value == '') {
                strRet += "身份別 ";
            }*/
            if (cbxIsLocal.checked) {
                if (ddlCity.value == '縣 市') {
                    strRet += "通訊地址-台灣本島(縣市) ";
                }
                if (ddlArea.value == '鄉鎮市區') {
                    strRet += "'通訊地址-台灣本島(鄉鎮市區) ";
                }
                if (tbxStreet.value == '') {
                    strRet += "'通訊地址-台灣本島(大道/路/街/部落) ";
                }
            }
            if (cbxIsAbroad.checked) {
                if (tbxOverseasCountry.value == '') {
                    strRet += "通訊地址-海外地址(國家/省城市/區) ";
                }
                if (tbxOverseasAddress.value == '') {
                    strRet += "通訊地址-海外地址(地址) ";
                }
            }
            if (cbxIsLocal.checked == false && cbxIsAbroad.checked == false) {
                strRet += "請勾選通訊地址 - 台灣本島或是海外地址 ";
            }
            if (strRet != "") {
                strRet += "欄位不得為空白！"
                alert(strRet);
                return false;
            }


            cnt = 0;
            sName = tbxDonor_Name.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 100) {
                alert('捐款人 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxIDNo.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('身分證/統編 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            /*sName = tbxReligionName.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('所屬教會 欄位長度超過限制！');
                return false;
            }
            cnt = 0;*/
            sName = tbxCellular_Phone.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('手機 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxTel_Office.value;
            for (i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('電話 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            /*sName = tbxTel_Home.value;
            for (i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('電話(夜) 欄位長度超過限制！');
                return false;
            }
            cnt = 0;*/
            sName = tbxFax.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('傳真 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxEMail.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 80) {
                alert('E-Mail 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxContactor.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('聯絡人 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            /*sName = tbxJobTitle.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('職稱 欄位長度超過限制！');
                return false;
            }
            cnt = 0;*/
            sName = tbxStreet.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('大道/路/街/部落 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxLane.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('巷 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxAlley.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('弄/衖 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxNo1.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('號 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxNo2.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('號之 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxFloor1.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('樓 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxFloor2.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('樓之 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxRoom.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('室 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxOverseasCountry.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 50) {
                alert('通訊地址國家/省城市/區 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxOverseasAddress.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 100) {
                alert('通訊地址 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxIsSendNewsNum.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 3) {
                alert('紙本月刊本數 欄位長度不得超過3位數！');
                return false;
            }

            if (isNaN(Number(tbxCellular_Phone.value)) == true) {
                alert('手機 欄位必須為數字！');
                tbxCellular_Phone.focus();
                return false;
            }
            if (isNaN(Number(tbxTel_Office_Loc.value)) == true || isNaN(Number(tbxTel_Office.value)) == true || isNaN(Number(tbxTel_Office_Ext.value)) == true) {
                alert('電話 欄位必須為數字！');
                tbxTel_Office.focus();
                return false;
            }
            if (isNaN(Number(tbxFax.value)) == true) {
                alert('傳真 欄位必須為數字！');
                tbxFax.focus();
                return false;
            }
            if (isNaN(Number(tbxIsSendNewsNum.value)) == true) {
                alert('紙本月刊本數 欄位必須為數字！');
                tbxIsSendNewsNum.focus();
                return false;
            }
        }
    </script>
    <style type="text/css">
        .style1
        {
            width: 35px;
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
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
        <tr>
            <td width="100%" colspan="10">
                &nbsp;&nbsp;
                <asp:Literal runat="server" ID="MemberMenu"></asp:Literal>
            </td>
        </tr>
    </table>
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        讀者資料維護
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1">姓名：
            </th>
            <td align="left" colspan="1" class="style1">
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9" Width="200px"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                性別：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlSex" CssClass="font9" 
                    AutoPostBack="True" onselectedindexchanged="ddlSex_SelectedIndexChanged"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1" >
                稱謂：
            </th>
            <td align="left" colspan="3">
                <asp:dropdownlist runat="server" ID="ddlTitle" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                身分別：
            </th>
            <td align="left" colspan="7">
                <%--<asp:CheckBoxList ID="cblDonor_Type" runat="server" RepeatDirection="Horizontal">
                </asp:CheckBoxList>--%>
                <%--20140514 新增 可以多選的CheckBoxList--%>
                <asp:Label ID="lblCheckBoxList" runat="server"></asp:Label>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                狀態：
            </th>
            <td align="left" colspan="7">
                <asp:dropdownlist runat="server" ID="ddlMember_Status" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                身分證/統編：
            </th>
            <td align="left" colspan="1" class="style1">
                <asp:TextBox runat="server" ID="tbxIDNo" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                出生日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxBirthday" onchange="CheckDateFormat(this, '出生日期');" runat="server" ></asp:TextBox>
                <img id="imgBirthday" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                手機：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxCellular_Phone" CssClass="font9"></asp:TextBox>
            </td>
            <!--th align="right" colspan="1">
                教育程度：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlEducation" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                職業別：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlOccupation" CssClass="font9"></asp:dropdownlist>
            </td-->
        </tr>
        <!--tr>
            <th align="right" colspan="1">
                婚姻狀況：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlMarriage" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                宗教信仰：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlReligion" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                所屬教會：
            </th>
            <td align="left" colspan="3" class="style1">
                <asp:TextBox runat="server" ID="tbxReligionName" CssClass="font9" Width="200px"></asp:TextBox>
            </td>
        </tr-->
        <tr>
            
            <th align="right" colspan="1">
                電話：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxTel_Office_Loc" CssClass="font9" 
                    MaxLength="5" Width="35px"></asp:TextBox>&nbsp;-
                <asp:TextBox runat="server" ID="tbxTel_Office" CssClass="font9"></asp:TextBox>&nbsp;-
                <asp:TextBox runat="server" ID="tbxTel_Office_Ext" CssClass="font9" 
                    MaxLength="5" Width="35px"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                傳真：
            </th>
            <td align="left" colspan="3"">
                <asp:TextBox runat="server" ID="tbxFax_Loc" CssClass="font9" 
                    MaxLength="5" Width="35px"></asp:TextBox>&nbsp;-
                <asp:TextBox runat="server" ID="tbxFax" CssClass="font9"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                E-Mail：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxEMail" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                聯絡人：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxContactor" CssClass="font9"></asp:TextBox>
            </td>
            <!--th align="right" colspan="1">
                服務單位：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxOrgName" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                職稱：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxJobTitle" CssClass="font9"></asp:TextBox>
            </td-->
        </tr>
        <tr>
            <th align="right" colspan="1">通訊地址：
            </th>
             <td align="left" colspan="7" >
                <asp:CheckBox ID="cbxIsLocal" runat="server" text="台灣本島" 
                     AutoPostBack="True" oncheckedchanged="cbxIsLocal_CheckedChanged"></asp:CheckBox>
                <asp:TextBox runat="server" ID="tbxZipCode" CssClass="font9" Width="60px"></asp:TextBox>
                <asp:dropdownlist runat="server" ID="ddlCity" CssClass="font9" 
                     AutoPostBack="True" onselectedindexchanged="ddlCity_SelectedIndexChanged"></asp:dropdownlist>
                <asp:dropdownlist runat="server" ID="ddlArea" CssClass="font9" 
                     AutoPostBack="True" onselectedindexchanged="ddlArea_SelectedIndexChanged"></asp:dropdownlist>
                <asp:TextBox runat="server" ID="tbxStreet" CssClass="font9" Width="100px"></asp:TextBox>大道/路/街/部落
                <asp:dropdownlist runat="server" ID="ddlSection" CssClass="font9"></asp:dropdownlist>段
                <asp:TextBox runat="server" ID="tbxLane" CssClass="font9" Width="40px"></asp:TextBox>巷<asp:TextBox runat="server" ID="tbxAlley" CssClass="font9" Width="40px"></asp:TextBox>
                <asp:dropdownlist runat="server" ID="ddlAlley" CssClass="font9"></asp:dropdownlist>
                <asp:TextBox runat="server" ID="tbxNo1" CssClass="font9" Width="40px"></asp:TextBox>-
                <asp:TextBox runat="server" ID="tbxNo2" CssClass="font9" Width="40px"></asp:TextBox>號
                <asp:TextBox runat="server" ID="tbxFloor1" CssClass="font9" Width="40px"></asp:TextBox>-
                <asp:TextBox runat="server" ID="tbxFloor2" CssClass="font9" Width="40px"></asp:TextBox>樓
                <asp:TextBox runat="server" ID="tbxRoom" CssClass="font9" Width="40px"></asp:TextBox>室
                Attn：<asp:TextBox runat="server" ID="tbxAttn" CssClass="font9" Width="300px"></asp:TextBox>
                <br />
                <asp:CheckBox ID="cbxIsAbroad" runat="server" text="海外地址" 
                     AutoPostBack="True" oncheckedchanged="cbxIsAbroad_CheckedChanged"></asp:CheckBox>
                <asp:TextBox runat="server" ID="tbxOverseasAddress" CssClass="font9" Width="250px"></asp:TextBox>國家/省城市/區 
                <asp:TextBox runat="server" ID="tbxOverseasCountry" CssClass="font9" Width="350px"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                介紹人：
            </th>
            <td align="left" colspan="7">
                <asp:TextBox runat="server" ID="tbxIntroducer_Name" CssClass="font9" ></asp:TextBox>
                <a href onclick="WindowsOpen()" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17">
                <asp:HiddenField runat="server" ID="HFD_Donor_Id"/>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                文宣品：
            </th>
            <td align="left" colspan="3">
                紙本月刊本數<asp:Textbox runat="server" ID="tbxIsSendNewsNum" CssClass="font9" 
                    Width="50px" Text="1"></asp:Textbox>
                <%--<asp:checkbox runat="server" ID="cbxIsDVD" Text="DVD" CssClass="font9" Checked="True"></asp:checkbox>--%>
                <asp:checkbox runat="server" ID="cbxIsSendEpaper" Text="電子文宣" CssClass="font9" Checked="True"></asp:checkbox>
            </td>
            <%--<td colspan="2">
                <asp:checkbox runat="server" ID="cbxIsContact" Text="不主動聯絡" CssClass="font9"></asp:checkbox>
            </td>--%>
            <td colspan="4">
                <asp:checkbox runat="server" ID="cbxIsErrAddress" Text="不主動聯絡-錯址" CssClass="font9"></asp:checkbox>
            </td> 
        </tr>
        <tr>
            <th colspan="1" align="right">
                備註：
            </th>
            <td colspan="7">
                <asp:Textbox runat="server" ID="tbxRemark" CssClass="font9" 
                    TextMode="MultiLine" Width="600px" Height="100px"></asp:Textbox> 
            </td>
         </tr>
         <tr>
             <th align="right" colspan="1">
                資料建檔日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCreate_Date" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                資料建檔人員：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCreate_User" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
                &nbsp;</td>
            <th align="right" colspan="1">
                最後異動日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxLastUpdate_Date" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                最後異動人員：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxLastUpdate_User" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
    </table>
    </div>
    <div class="function">
        <asp:Button ID="btnEdit" class="npoButton npoButton_Modify" runat="server" 
            Text="修改" OnClientClick="return CheckFieldMustFillBasic()" onclick="btnEdit_Click" />
        <asp:Button ID="btnDel" class="npoButton npoButton_Del" runat="server" 
            Text="刪除" OnClientClick="return confirm('您確定要刪除嗎？')" onclick="btnDel_Click" />
        <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
            Text="取消" onclick="btnExit_Click" />
    </div>
    </form>
</body>
</html>

