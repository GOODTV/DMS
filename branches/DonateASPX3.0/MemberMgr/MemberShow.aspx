<%@ Page Language="C#" AutoEventWireup="true" CodeFile="MemberShow.aspx.cs" Inherits="MemberMgr_MemberShow" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>會員查詢</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        function ReturnOpener(val) {
            var url = window.location.toString();
            var str = "";
            var HFD_Donor_Id = "";
            var tbxIntroducer_Name = "";
            if (url.indexOf("?") != -1) {
                var ary = url.split("?")[1].split("&");
                for (var i in ary) {
                    str = ary[i].split("=")[0];
                    if (str == "Donor_Id") {
                        HFD_Donor_Id = decodeURI(ary[i].split("=")[1]);
                    }
                    if (str == "Donor_Name") {
                        tbxIntroducer_Name = decodeURI(ary[i].split("=")[1]);
                    }
                }
            }
            var Donor_Id = val.split("|")[0];
            var Introducer_Name = val.split("|")[1];

            opener.document.getElementById(HFD_Donor_Id).value = Donor_Id;
            opener.document.getElementById(tbxIntroducer_Name).value = Introducer_Name;
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
                會員姓名：
            </th>
            <td align="left" colspan="1" >
                 <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9" ></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                會員編號：
            </th>
            <td align="left" colspan="1" >
                 <asp:TextBox runat="server" ID="tbxDonor_Id" CssClass="font9" ></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                身分別：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlDonor_Type" CssClass="font9" ></asp:dropdownlist>
            </td>
            <td align="right" colspan="1" >
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" OnClick="btnQuery_Click"/>
            </td>
        </tr> 
        <tr>
            <td  align="center" width="100%" colspan="7">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
