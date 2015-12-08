<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Donate_Other_Report8_Print_Excel.aspx.cs" Inherits="Report_Custom_Report_Donate_Other_Report8_Print_Excel" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html  xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>個別奉獻動機及收視管道統計分析</title>
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
    <asp:HiddenField runat="server" ID="HFD_DonorName_other8" />
    <asp:HiddenField runat="server" ID="HFD_DonateMotive1_other8" />
    <asp:HiddenField runat="server" ID="HFD_DonateMotive2_other8" />
    <asp:HiddenField runat="server" ID="HFD_DonateMotive3_other8" />
    <asp:HiddenField runat="server" ID="HFD_DonateMotive4_other8" />
    <asp:HiddenField runat="server" ID="HFD_DonateMotive5_other8" />
    <asp:HiddenField runat="server" ID="HFD_DonateMotive6_other8" />
    <asp:HiddenField runat="server" ID="HFD_WatchMode1_other8" />
    <asp:HiddenField runat="server" ID="HFD_WatchMode2_other8" />
    <asp:HiddenField runat="server" ID="HFD_WatchMode3_other8" />
    <asp:HiddenField runat="server" ID="HFD_WatchMode4_other8" />
    <asp:HiddenField runat="server" ID="HFD_WatchMode5_other8" />
    <asp:HiddenField runat="server" ID="HFD_WatchMode6_other8" />
    <asp:HiddenField runat="server" ID="HFD_WatchMode7_other8" />
    <asp:HiddenField runat="server" ID="HFD_WatchMode8_other8" />
    <asp:HiddenField runat="server" ID="HFD_WatchMode9_other8" />
        <div>
           <asp:Literal ID="GridList" runat="server"></asp:Literal>
        </div>
   </form>
</body>
</html>
