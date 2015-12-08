<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Import_Ibon.aspx.cs" Inherits="Import_Import_Ibon" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>Ibon捐款匯入</title>
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
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxDonate_Date",   // id of the input field
                button: "imgDonate_Date"     // 與觸發動作的物件ID相同
            });
        }
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var cnt = 0;
            var sName = "";
            var ddlDept = document.getElementById('ddlDept');
            var tbxSheet_Name = document.getElementById('tbxSheet_Name');
            var FileUpload = document.getElementById('FileUpload');
            if (ddlDept.value == "") {
                strRet += "機構 ";
            }
            if (tbxSheet_Name.value == "") {
                strRet += "工作表名稱 ";
            }
            if (FileUpload.value == "") {
                strRet += "檔案位置 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet)
                return false;
            }
            else {
                if (window.confirm('您是否確定要匯入捐款資料？\n\n※請注意匯入過程中請勿關閉視窗') == true) {
                    return true;
                }
                else {
                    return false;
                }
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
    <asp:HiddenField runat="server" ID="HFD_Donor_Id" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        Ibon捐款匯入 
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
                捐款日期：
            </th>
            <td align="left" colspan="1">
                <asp:RadioButton ID="rbDonate_Date_File" runat="server" Text="依檔案內的『 拉單狀態日期 』" 
                    Checked="True" GroupName="1" />
                <asp:RadioButton ID="rbDonate_Date_New" runat="server" Text="另訂捐款日期為：" 
                    GroupName="1"/>
                <asp:TextBox ID="tbxDonate_Date" runat="server" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgDonate_Date" alt="" src="../images/date.gif" />
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                工作表名稱：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxSheet_Name" CssClass="font9" size="11" maxlength="60" value="Sheet1"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                上傳檔案：
            </th>
            <td align="left" colspan="1">
                <asp:FileUpload ID="FileUpload" runat="server" />
                <asp:Button ID="btnInput" runat="server"  Width="45mm"
                    Text="匯入『Ibon捐款』資料"  CssClass="npoButton npoButton_Export"
                    OnClientClick=" return CheckFieldMustFillBasic();" onclick="btnInput_Click"/>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                ※注意事項：
            </th>
            <td align="left" colspan="1">
                1.您選擇的『&nbsp;<font color="#FF0000">機構名稱</font>&nbsp;』及『&nbsp;<font color="#FF0000">捐款日期</font>&nbsp;』會影響&nbsp;<font color="#FF0000">收據編號</font>&nbsp;取號規則，上傳前請您務必確認。<br />
                2.工作表名稱必須與檔案的工作表名稱一致否則將導致作業失敗。<br />
                3.檔案必須為&nbsp;Excel&nbsp;格式，如您的檔案來源自網站下載請務必&nbsp;<font color="#FF0000">另存新檔</font>&nbsp;，檔案大小請勿超過10M。<br />
                4.相關欄位說明請下載&nbsp;<a href='example/Ibon參考範例.xls'>參考範例</a>&nbsp;。
            </td> 
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
