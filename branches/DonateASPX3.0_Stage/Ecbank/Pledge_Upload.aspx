<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Pledge_Upload.aspx.cs" Inherits="Ecbank_Pledge_Upload" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>轉帳授權書上傳</title>
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
        function Confirm() {
            if (confirm("尚未上傳檔案或是無此檔案，您要下載範例檔嗎？") == true)
                location.href = '../UpLoad/定額捐款授權4合一表.doc';
            else
                return false;
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
        轉帳授權書上傳 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1">
                轉帳授權書上傳：
            </th>
            <td align="left" colspan="1">
                <asp:FileUpload ID="FileUpload" runat="server" />
                <asp:Button ID="btnUpload" runat="server"  Width="25mm" Text="重新上傳"  
                CssClass="npoButton npoButton_Upload" onclick="btnUpload_Click"/>
                <asp:Button ID="btnDownload" runat="server"  Width="35mm" Text="轉帳授權書下載"  
                CssClass="npoButton npoButton_Download" onclick="btnDownload_Click"/>
            </td> 
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
