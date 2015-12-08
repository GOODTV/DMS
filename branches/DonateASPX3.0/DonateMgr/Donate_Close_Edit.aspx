<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Donate_Close_Edit.aspx.cs" Inherits="DonateMgr_Donate_Close_Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title> 系統關帳【修改】</title>
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
        var oldOnload = window.onload || function () {
            showtime();
        };
        window.onload = function () {
            oldOnload();
            initCalendar();
        }
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxDonate_Close",   // id of the input field
                button: "imgDonate_Close"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxContribute_Close",   // id of the input field
                button: "imgContribute_Close"     // 與觸發動作的物件ID相同
            });
        }
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var tbxDonate_Close = document.getElementById('tbxDonate_Close');
            var tbxContribute_Close = document.getElementById('tbxContribute_Close');
            if (tbxDonate_Close.value == "") {
                strRet += "捐款關帳日 ";
            }
            if (tbxContribute_Close.value == "") {
                strRet += "捐物資關帳日 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet)
                return false;
            }
            else {
                return true;
            }
        }
        function showtime() {
            var now = new Date();
            var year = now.getFullYear();
            var month = now.getMonth() + 1;
            var day = now.getDate();
            time = year + '/' + month + '/' + day;
            var lblTime = document.getElementById('lblTime');
            var lblTime2 = document.getElementById('lblTime2');
            lblTime.innerHTML = time;
            lblTime2.innerHTML = time;
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
        系統關帳【修改 
        】</h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDept" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款關帳日：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Close" onchange="CheckDateFormat(this, '捐款關帳日');" runat="server" ></asp:TextBox>
                <img id="imgDonate_Close" alt="" src="../images/date.gif" />
                <font color="#000080">&nbsp;&nbsp;開放權限修改人員(&nbsp;<font color="#ff0000">限<asp:Label 
                    ID="lblTime" runat="server" ForeColor="Red"></asp:Label></font> )：</font>
                <asp:dropdownlist runat="server" ID="ddlDonate_Open_User" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐物資關帳日：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxContribute_Close" onchange="CheckDateFormat(this, '捐物資關帳日');" runat="server"  ></asp:TextBox>
                <img id="imgContribute_Close" alt="" src="../images/date.gif" />
                <font color="#000080">&nbsp;&nbsp;開放權限修改人員(&nbsp;<font color="#ff0000">限<asp:Label 
                    ID="lblTime2" runat="server" ForeColor="Red"></asp:Label></font> )：</font>
                <asp:dropdownlist runat="server" ID="ddlContribute_Open_User" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
    </table>
    </div>
    <div class="function">
        <asp:Button ID="btnEdit" class="npoButton npoButton_Modify" runat="server" 
            Text="修改" OnClientClick= "return CheckFieldMustFillBasic();"
            onclick="btnEdit_Click" />
        <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
            Text="離開" onclick="btnExit_Click" />
    </div>
    </form>
</body>
</html>

