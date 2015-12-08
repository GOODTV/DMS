<%@ Page Language="C#" AutoEventWireup="true" CodeFile="GiftImport.aspx.cs" Inherits="DonorMgr_GiftImport" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>公關贈品批次匯入</title>
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
            var FileUpload = document.getElementById('FileUpload');
            if (ddlDept.value == "") {
                strRet += "機構 ";
            }
            if (FileUpload.value == "") {
                strRet += "檔案位置 ";
            }
            if (strRet != "") {
                strRet += "不可為空白！";
                alert(strRet)
                return false;
            }
            else {
                if (window.confirm('您是否確定要匯入公關贈品資料？\n\n※請注意匯入過程中請勿關閉視窗') == true) {
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
        公關贈品批次匯入 
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
                上傳檔案：
            </th>
            <td align="left" colspan="1">
                <asp:FileUpload ID="FileUpload" runat="server" />
                <asp:Button ID="btnInput" runat="server"  Width="30mm"
                    Text="匯入公關贈品"  CssClass="npoButton npoButton_Export"
                    OnClientClick=" return CheckFieldMustFillBasic();" onclick="btnInput_Click"/>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                ※注意事項：
            </th>
            <td align="left" colspan="1" style="font-size:16px;line-height:24px;">
                1.檔案必須為&nbsp;Excel&nbsp;格式，如您的檔案來源自網站下載請務必&nbsp;<font color="#FF0000">另存新檔</font>&nbsp;，檔案大小請勿超過10M。<br />
                2.相關欄位說明<font color="red"><b>請務必下載</b></font>&nbsp;<a style="font-size:16px; color:blue; font-weight:bold;" href='../Import/example/公關贈品匯入參考範例.xls'>參考範例</a>&nbsp;。
            </td> 
        </tr>
        <tr>
            <td width="100%" colspan="8" align="center">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>