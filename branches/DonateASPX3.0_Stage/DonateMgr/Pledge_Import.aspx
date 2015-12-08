<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Pledge_Import.aspx.cs" Inherits="DonateMgr_Pledge_Import" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>匯入</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        function ReturnOpener() {
            var url = window.location.toString();
            var str = "";
            var HFD_Pledge_Import = "";
            if (url.indexOf("?") != -1) {
                HFD_Pledge_Import = url.split("=")[1];
            }
            opener.document.getElementById(HFD_Pledge_Import).value = '1';
            window.close();
        }
    </script>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <div id="container">
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
        <tr>
            <th align="right" colspan="1">
                信用卡批次授權(一般)回覆檔匯入：
            </th>
            <td align="left" colspan="1" >
                <asp:FileUpload ID="FileUpload" runat="server" />
            </td>
            <td align="right" colspan="1" >
                <asp:Button ID="btnImport" CssClass="npoButton npoButton_New" runat="server"  Width="20mm"
                    Text="匯入" OnClientClick="return confirm('您是否確定要匯入TXT授權資料？')" OnClick="btnInput_Click" />
            </td>
        </tr> 
    </table>
    </div>
    </form>
</body>
</html>
