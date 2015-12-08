<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateMobile.aspx.cs" Inherits="Online_DonateMobile" EnableEventValidation="false" validateRequest="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
    <title>行動版/OTT奉獻</title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:HiddenField ID="strRqXML" runat="server" value=""/>
    <div align="center">
    <iframe id="my_iframe" name="my_iframe" width="500" height="400" scrolling="no" src="https://sslpayment.uwccb.com.tw/EPOSService/Payment/Mobile/OrderInitial.aspx" 
        frameborder="0" style="position: relative; top: -55px"></iframe>
    </div>
    </form>
</body>
</html>
