<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Donate_Other_Report6_Print_Excel.aspx.cs" Inherits="Report_Custom_Report_Donate_Other_Report6_Print_Excel" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>捐款人統計分析查詢報表</title>
    <link href="/include/main.css" rel="stylesheet" type="text/css" />
    <link href="/include/table.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <script type="text/javascript">
        function SelfClose() {
            window.opener = null;
            window.open('', '_self', '');
            window.close();
        }
    </script>
    <form id="Form1" runat="server">
    <asp:HiddenField runat="server" ID="HFD_Type_other6" />
    <asp:HiddenField runat="server" ID="HFD_DonateDateS_other6" />
    <asp:HiddenField runat="server" ID="HFD_DonateDateE_other6" />
    <asp:HiddenField runat="server" ID="HFD_Is_Abroad_other6" />
    <asp:HiddenField runat="server" ID="HFD_Is_ErrAddress_other6" />
    <asp:HiddenField runat="server" ID="HFD_Sex_other6" />
        <div>
           <asp:Literal ID="GridList" runat="server"></asp:Literal>
        </div>
   </form>
</body>
</html>
