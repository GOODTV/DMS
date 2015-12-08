<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateReturnUrl.aspx.cs"
    Inherits="Online_DonateReturnUrl" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>信用卡奉獻交易確認</title>
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
<body style="background-color: #EEE">
    <form id="Form1" runat="server" >
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_ItemList" runat="server" />
<%--
          <input type="hidden" name="storeid" value="1350010845">
          <input type="hidden" name="password" value="iepaytest">
          <input type="hidden" name="orderid" value="20130613MKEP">
          <input type="hidden" name="account" value="222">
          <input type="hidden" name="remark" value="">
          <input type="hidden" name="param" value="">
          <input type="hidden" name="customer" value="222">	
          <input type="hidden" name="cellphone" value="0987123456">
          <input type="hidden" name="charset" value="big5">
          <input type="hidden" name="invoiceflag" value="">
          <input type="hidden" name="paytype" value="1:信用卡付款">	
--%>
<%--
    <asp:HiddenField ID="storeid" runat="server"  value="1350010845"/>
    <asp:HiddenField ID="password" runat="server" value="iepaytest"/>
    <asp:HiddenField ID="orderid" runat="server" value="20130613MKEP"/>
    <asp:HiddenField ID="account" runat="server" value="222"/>
    <asp:HiddenField ID="remark" runat="server" value=""/>
    <asp:HiddenField ID="param" runat="server" value=""/>
    <asp:HiddenField ID="customer" runat="server" value="222"/>
    <asp:HiddenField ID="cellphone" runat="server" value="0987123456"/>
    <asp:HiddenField ID="charset" runat="server" value="big5"/>
    <asp:HiddenField ID="invoiceflag" runat="server" value=""/>
    <asp:HiddenField ID="paytype" runat="server" value="1:信用卡付款"/>
    <asp:HiddenField ID="purpose" runat="server" value=""/>
--%>
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

    <table width="860" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td width="436">
                <img alt="aa" src="images/bar_logo.jpg" width="390" height="40" />
            </td>
            <td width="424" style="background-image: url('images/bar_item.jpg')">
                <table width="95%" border="0" align="center" cellpadding="5" cellspacing="5">
                    <tr>
                        <td align="left" valign="bottom">
                            <strong><font color="#003399">信用卡奉獻交易確認</font></strong>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td height="20" colspan="3">
            </td>
        </tr>
        <tr valign="top">
            <td height="405" colspan="3" style="background-image: url('images/bk_fish.jpg')"
                align="center">
                <br />
                            <asp:Label ID="lblTitle" runat="server" Style="color: #8B0000;"> 感謝您的奉獻！ 
							<br/><br/> 以下是您本次刷卡授權記錄的結果，您可以列印保留 <br/><br/>
							
							願神大大賜福與您!</asp:Label>
                <br />
                <br />
<%--                <table align="center" border="0" cellpadding="0" cellspacing="1" class="table_v" width="70%">
                    <tr>
                        <td>
                            <asp:Label ID="lblDonorName" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Label ID="lblPhoneNum" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Label ID="lblEmail" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Label ID="lblDonateContent" runat="server"></asp:Label>
                        </td>
                    </tr>
					<asp:Label ID="lblCart" runat="server"></asp:Label>
					
					
                </table>--%>
			                      
                      
                            <font color='red'>您有任何捐款問題請和我們聯絡：
                            <br />總機：02-8024-3911<br />傳真：02-8024-3933<br />
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
                <asp:Button ID="btnBackPay" class="Online Online_Button5" runat="server" Text="重新授權付款"
                    OnClick="btnBackPay_Click" Width="174px" Visible="false" />
                <asp:Button ID="btnConfirm" class="Online Online_Button5" runat="server" Text="回到奉獻首頁"
                    OnClick="btnConfirm_Click" Width="174px" />
            </td>
        </tr>
    </table>

   <%-- <div class="function">
        <asp:Button ID="btnBackPay" class="npoButton npoButton_Submit" runat="server" Text="重新授權付款"
            OnClick="btnBackPay_Click" Width="140px" Visible="false" />
        <asp:Button ID="btnConfirm" class="npoButton npoButton_Submit" runat="server" Text="回到奉獻首頁"
            OnClick="btnConfirm_Click" Width="140px" />
    </div>--%>


    </form>
</body>
</html>
