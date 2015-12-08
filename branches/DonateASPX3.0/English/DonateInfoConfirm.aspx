<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateInfoConfirm.aspx.cs"
    Inherits="Online_DonateInfoConfirm" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Confirm Donation</title>
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
            $('#lblDonateContent td:nth-child(1)').css('text-align', 'center');
            $('#lblDonateContent td:nth-child(2)').css('text-align', 'center');
            $('#lblDonateContent td:nth-child(3)').css('text-align', 'right');
        });                     //end of ready()

    </script>
    <style type="text/css">
    td,th {
    font-size: 14px;
    color: #000;
    }
    </style>
</head>
<body style="background-color: #EEE" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">    
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
    <asp:HiddenField ID="storeid" runat="server"  value="1184837278"/>
    <asp:HiddenField ID="password" runat="server" value="81582054"/>
    <asp:HiddenField ID="orderid" runat="server" value=""/>
    <asp:HiddenField ID="account" runat="server" value=""/>
    <asp:HiddenField ID="remark" runat="server" value=""/>
    <asp:HiddenField ID="param" runat="server" value=""/>
    <asp:HiddenField ID="customer" runat="server" value="GOODTV"/>
    <asp:HiddenField ID="cellphone" runat="server" value=""/>
    <asp:HiddenField ID="charset" runat="server" value=""/>
    <asp:HiddenField ID="invoiceflag" runat="server" value=""/>
    <asp:HiddenField ID="paytype" runat="server" value=""/>
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
                            <strong><font color="#003399">Confirm Donation</font></strong>
                        </td>
                        <td align="right" valign="bottom">
                            <strong><font color="#003399">Language：</font><a href="../Online/DonateInfoConfirm.aspx" target="_self"><font size="3">中文</font></a></strong>
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
            <td height="450" colspan="3" style="background-image: url('images/bk_fish2.jpg')"
                align="center">
                <br />
                           <%-- <asp:Label ID="lblTitle" runat="server" Style="color: #8B0000;">平安！請確認奉獻明細</asp:Label>
                <br />--%>
                <table align="center" border="0" cellpadding="0" cellspacing="1">
                <tr>
                    <td>
                        <asp:Label ID="lblStep1" runat="server" Text=" Your Donation Items >> " ForeColor="MediumBlue" />
                        <asp:Label ID="lblStep2" runat="server" Text=" Donor Information >> " ForeColor="MediumBlue" />
                        <asp:Label ID="lblStep3" runat="server" Text=" Confirm Donation >> " ForeColor="Red" />
                        <asp:Label ID="lblStep4" runat="server" Text=" Pledge Information >> " ForeColor="MediumBlue" />
                        <asp:Label ID="lblStep5" runat="server" Text=" Confirm Payment " ForeColor="MediumBlue" />
                    </td>
                </tr>
                </table>
                <br />               
                
                <table align="center" border="0" cellpadding="0" cellspacing="1" class="table_v2" width="70%">
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
                            <asp:Label ID="lblDonateContent" runat="server"  Style="font-weight: bold;"></asp:Label>
                        </td>
                    </tr>
                    <asp:Label ID="lblCart" runat="server"></asp:Label>
                </table><br>
                <asp:Panel ID="msgOnce" runat="server" Visible="false">
                    <span align='center' style='color: #444444; font-size: 10; font-weight: bold;'>
                            For one time donation: We accept VISA, Master, JCB, and UnionPay. 
                    <asp:Image ID="Image4" runat="server" ImageUrl="images/Visa_Master.png" /><asp:Image ID="Image5" runat="server" ImageUrl="images/JCB.png" /><asp:Image ID="Image6" runat="server" ImageUrl="images/UnionPay.png" />
                    </span></asp:Panel><br>
                <asp:Panel ID="msgPeriod" runat="server" Visible="false">
                <span align='center' style='color: #444444; font-size: 10; font-weight: bold;'>
                        1.Recurring plan: your donation may recur on the monthly, quarterly and yearly basis of your choice.<br><br>
                        2.If you would like to terminate your recurring plan, please call (02)8024-3911 .Thank you for your support.
                    </span></asp:Panel>
                <asp:Panel ID="msgPanel" runat="server" Visible="false">
                <table align="center" width="80%" border="0" cellpadding="0" cellspacing="1">   
                    <tr>
                        <td><font color="red" size="3"><b>
                            <br />We have acknowledged your recurring donation. Thank you! You will be notified when the initial payment is successfully processed.<br><br>
                            Next we are going to proceed the one time donation in your cart.</b></font>
                        </td>
                    </tr>                   
                </table></asp:Panel>
            </td>
        </tr>
        <tr>
            <td height="10" colspan="2">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center">                
                    <asp:Button ID="btnPrev" class="Online Online_Button5" runat="server" Text="Back to Donor information"
                        OnClick="btnPrev_Click" Width="174px" height="40px" />
                    <asp:Button ID="btnConfirm" class="Online Online_Button5" runat="server" Text="Next"
                        OnClick="btnConfirm_Click" Width="174px" height="40px" />
            </td>
            
        </tr>
<%--        <tr>
            <td colspan="2">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                        <td width="133">
                            <asp:ImageButton ID="btnContinue" runat="server" OnClick="btnContinue_Click" ImageUrl="images/btn_繼續奉獻_O.jpg"
                                onmouseover="this.src='images/btn_繼續奉獻_down.jpg'" onmouseout="this.src='images/btn_繼續奉獻_O.jpg'" />
                        </td>
                        <td width="10">
                            &nbsp;
                        </td>
                        <td width="154">
                            <asp:ImageButton ID="btnCheckOut" runat="server" OnClick="btnCheckOut_Click" ImageUrl="images/btn_完成奉獻_O.jpg"
                                onmouseover="this.src='images/btn_完成奉獻_down.jpg'" onmouseout="this.src='images/btn_完成奉獻_O.jpg'" />
                        </td>
                    </tr>
                </table>
            </td>
        </tr>--%>
    </table>

    <%--<div class="function">
        
    </div>--%>


    </form>
</body>
</html>
