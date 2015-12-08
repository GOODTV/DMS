<%@ Page Language="C#" AutoEventWireup="true" CodeFile="chart4.aspx.cs" Inherits="ChartDirector_chart4" %>
<%@ Import Namespace="ChartDirector" %>
<%@ Register TagPrefix="chart" Namespace="ChartDirector" Assembly="netchartdir" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <chart:WebChartViewer id="WebChartViewer1" runat="server" />
    </div>
    </form>
</body>
</html>
