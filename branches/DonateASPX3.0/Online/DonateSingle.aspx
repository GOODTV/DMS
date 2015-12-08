<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateSingle.aspx.cs"
    Inherits="Online_DonateSingle" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>線上奉獻</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/jquery.address.js"></script>
    <script type="text/javascript" src="../include/jquery.field.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {

        });                     //end of ready()

    </script>
</head>
<body bgcolor="#474747">
    <form id="Form1" runat="server">
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_ItemList" runat="server" />
    <%--    <h1>
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        <asp:Literal ID="litTitle" runat="server">Good TV 線上奉獻 - 確認奉獻方式</asp:Literal>
    </h1>--%>
    <table border="0" align="center">
        <tr>
            <td rowspan="2" background="images/奉獻後台＿BK_03.jpg" width="361" height="57">
            </td>
            <td width="309" height="57" rowspan="2" align="center" background="images/奉獻後台＿BK_04.jpg">
                <table width="80%" border="0" cellspacing="2" cellpadding="2">
                    <tr>
                        <td align="left">
                             <font color="#0000FF"> - 單　筆　奉　獻</font>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table align="center" border="0" cellpadding="0" cellspacing="1" class="table_v"
        width="60%">
        <tr>
            <td >
                <asp:Label ID="lblTitle" runat="server" Style="color: #8B0000;">平安שלום שלום！單筆奉獻處理</asp:Label>
            </td>
        </tr>
    </table>
    <div class="function">
        <asp:Button ID="btnBackDefault" class="npoButton npoButton_Submit" runat="server" Text="確定"
            OnClick="btnBackDefault_Click" Width="100px" />
    </div>
    </form>
</body>
</html>
