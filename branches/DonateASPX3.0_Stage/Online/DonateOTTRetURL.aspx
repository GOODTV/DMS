<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateOTTRetURL.aspx.cs" Inherits="Online_DonateOTTRetURL"  validateRequest="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
    <title>信用卡奉獻交易確認</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <style type="text/css">
    <!--
    body {
	    background-color: #cdf2c5;
	    font-family: Georgia, "Times New Roman", Times, serif;
	    font-size: 16px;
	    font-style: normal;
	    line-height: 20px;
	    font-weight: bold;
	    font-variant: normal;
	    margin-left: 0px;
	    margin-top: 0px;
	    margin-right: 0px;
	    margin-bottom: 0px;
    }
    body,td,th {
	    font-family: Verdana, Geneva, sans-serif;
	    font-size: 16px;
	    line-height: normal;
	    font-weight: normal;
    }
    .title {
	    font-family: Verdana, Geneva, sans-serif;
	    font-size: 18px;
	    font-weight: bold;
	    text-align: left;
    }
    -->
    </style>
</head>
<body>
    <form id="Form1" runat="server" >
    <asp:HiddenField ID="storeid" runat="server"  value=""/>
    <asp:HiddenField ID="password" runat="server" value=""/>
    <asp:HiddenField ID="orderid" runat="server" value=""/>
    <asp:HiddenField ID="account" runat="server" value=""/>
    <asp:HiddenField ID="remark" runat="server" value=""/>
    <asp:HiddenField ID="param" runat="server" value=""/>
    <asp:HiddenField ID="customer" runat="server" value=""/>
    <asp:HiddenField ID="cellphone" runat="server" value=""/>
    <asp:HiddenField ID="charset" runat="server" value=""/>
    <asp:HiddenField ID="invoiceflag" runat="server" value=""/>
    <asp:HiddenField ID="paytype" runat="server" value=""/>
    <asp:HiddenField ID="purpose" runat="server" value=""/>

    <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr style="background-color:#FFF44E; background-repeat:repeat" align="center" valign="top">
            <td align="center" valign="middle">
                <img alt="aa" src="images/bar_logo2.jpg" width="129" height="32" /><strong><font color="#003399">信用卡奉獻交易確認</font></strong>
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr valign="top">
            <td height="405" colspan="3" align="center">
                <br />
                <asp:Label ID="lblTitle" runat="server" Style="color: #8B0000; font-family:'Microsoft JhengHei';"><font color="red" size="4"><b>感謝您的奉獻！</b></font>
                    <br/><br/> 請確認您本次刷卡授權記錄的結果。 
					<br/><br/>願神大大賜福與您！</asp:Label>
                <br />
                <br />
                <font color="#8B0000" face="Microsoft JhengHei">您有任何捐款問題請和我們聯絡：
                <br />
                總機 Tel：02-8024-3911
                <br />
                傳真 Fax：02-8024-3933
                <br />
                地址：235 新北市中和區中正路911號6樓 捐款服務組</font>
            </td>
        </tr>
        <tr>
            <td height="10" colspan="2">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center">
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
