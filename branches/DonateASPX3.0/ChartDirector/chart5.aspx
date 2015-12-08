<%@ Page Language="C#" AutoEventWireup="true" CodeFile="chart5.aspx.cs" Inherits="ChartDirector_chart5" %>
<%@ Import Namespace="ChartDirector" %>
<%@ Register TagPrefix="chart" Namespace="ChartDirector" Assembly="netchartdir" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
</head>
<body>
    <form id="Form1" method="post" runat="server">
        <div>
            捐款年度：
            <asp:DropDownList ID="yearSelect" runat="server">
                    </asp:DropDownList>年
            <asp:Button id="OKPB" runat="server" Text="查詢"></asp:Button>
            <br /><br />
            <chart:WebChartViewer id="WebChartViewer1" runat="server" />
        </div>
    </form>
</body>
</html>
