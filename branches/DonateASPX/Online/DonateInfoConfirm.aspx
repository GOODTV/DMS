<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateInfoConfirm.aspx.cs"
    Inherits="Online_DonateInfoConfirm" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>線上明細確認</title>
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
--%>
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
    <asp:HiddenField ID="purpose" runat="server" />

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
                            <strong><font color="#003399">確認您的奉獻明細</font></strong>
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
                           <%-- <asp:Label ID="lblTitle" runat="server" Style="color: #8B0000;">平安！請確認奉獻明細</asp:Label>
                <br />--%>
                <table align="center" border="0" cellpadding="0" cellspacing="1">
                <tr>
                    <td>
                        <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="Chocolate" />
                        <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="Chocolate" />
                        <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="Brown" />
                        <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="Chocolate" />
                        <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="Chocolate" />
                    </td>
                </tr>
                </table>
                <br />               
                
                <table align="center" border="0" cellpadding="0" cellspacing="1" class="table_v" width="70%">
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
                </table>

                <asp:Panel ID="msgPanel" runat="server" Visible="false">
                <table align="center" border="0" cellpadding="0" cellspacing="1">   
                    <tr>
                        <td><font color="red" size="2">
                            <br />我們已收到您的定期定額授權資料，
                            <br />當第一次授權扣款成功後，將會再通知您。
                            <p />接下來，將進行「單筆奉獻」扣款作業。</font>
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
                    <asp:Button ID="btnPrev" class="npoButton npoButton_PrevStep" runat="server" Text="上一步"
                        OnClick="btnPrev_Click" height="40px" />
                    <asp:Button ID="btnConfirm" class="npoButton npoButton_Submit" runat="server" Text="確認, 開始付款"
                        OnClick="btnConfirm_Click" Width="140px" height="40px" />
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
