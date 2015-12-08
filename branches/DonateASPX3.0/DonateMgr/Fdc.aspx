<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Fdc.aspx.cs" Inherits="DonateMgr_Fdc" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>匯出扣除額單據</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            InitMenu();
        });
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var cnt = 0;
            var sName = "";
            var ddlDept = document.getElementById('ddlDept');
            var ddlDonate_Date_Year = document.getElementById('ddlDonate_Date_Year');
            var ddlAct_Id = document.getElementById('ddlAct_Id');
            var tbxUniform_No = document.getElementById('tbxUniform_No');
            var tbxLicence = document.getElementById('tbxLicence');
            if (ddlDept.value == "") {
                strRet += "機構 ";
            }
            if (ddlDonate_Date_Year.value == "") {
                strRet += "捐款年度 ";
            }
            if (tbxUniform_No.value == "") {
                strRet += "機構統編 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet)
                return false;
            }
            if (isNaN(Number(tbxUniform_No.value)) == true) {
                alert('機構統編  欄位必須為數字！');
                return false;
            }
            cnt = 0;
            sName = tbxUniform_No.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 8) {
                alert('機構統編 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxLicence.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 60) {
                alert('許可文號 欄位長度超過限制！');
                return false;
            }
            else {
                if (window.confirm('您是否確定要將查詢結果匯出？') == true) {
                    return true;
                }
                else {
                    return false;
                } 
            }
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
    <form id="Form1" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        匯出扣除額單據 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9" 
                    AutoPostBack="True" onselectedindexchanged="ddlDept_SelectedIndexChanged"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                捐款年度：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Date_Year" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlAct_Id" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                機構統編：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxUniform_No" CssClass="font9" size="16" maxlength="8"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                許可文號：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxLicence" CssClass="font9" size="30" 
                    maxlength="60" Width="250px"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                ※注意事項：
            </th>
            <td align="left" colspan="5">
                1.本單據適用對象僅限於收據身分證(統一編號)資料已填寫者。<br />
                2.【許可文號】泛指專案核准文號、註管機關登記、立案文號(如有向法院辦理登記應加註法院登記簿<br />之&nbsp;&nbsp;&nbsp;冊、頁、號)
                或財團法人設立許可文號。<br />
                3.匯出後請將檔案另存成國稅局標準檔名<b>【捐款年度民國年.31.統一編號.txt】</b>
            </td> 
        </tr>
        <tr>
            <td align="left" colspan="6" class="style1">
                <asp:Button ID="btnTxt" runat="server"  Width="20mm"
                    Text="匯出"  CssClass="npoButton npoButton_Txt"
                    OnClientClick=" return CheckFieldMustFillBasic();" onclick="btnTxt_Click"/>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>