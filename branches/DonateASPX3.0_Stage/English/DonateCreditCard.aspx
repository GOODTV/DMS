<%@ Page Language="C#" EnableEventValidation="false" AutoEventWireup="true" CodeFile="DonateCreditCard.aspx.cs"
    Inherits="Online_DonateCreditCard" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Pledge Information</title>
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
    <script type="text/javascript" src="../include/jquery.maxlength.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            $('#tbxAccount_No1,#tbxAccount_No2,#tbxAccount_No3,#tbxAccount_No4,#txtAuthorize').not('').bind("keyup", function (e) {
                var val = $(this).val();
                if (isNaN(val)) {
                    val = val.replace(/[^0-9]/g, '');
                    $(this).val(val);
                }
            });
        });                     //end of ready()

        function Foolproof() {
            if ($('#txtCardBank').val() == '') {
                alert('Bank is required.');
                $('#txtCardBank').focus();
                return false;
            }
            if ($('#txtCardBank').val().length > 20) {
                alert('The length of the field "Bank" can not exceed 20 characters!');
                $('#txtCardBank').focus();
                return false;
            }
            if ($('#ddlCardType').val() == '') {
                alert('Please choose card type!');
                $('#ddlCardType').focus();
                return false;
            }
            if (($('#tbxAccount_No1').val().trim() + $('#tbxAccount_No2').val().trim() + $('#tbxAccount_No3').val().trim() + $('#tbxAccount_No4').val().trim()).length < 16) {
                alert('The length of the field "Card No." should be 16 characters.');
                $('#tbxAccount_No1').focus();
                return false;
            }
            if ($('#txtAuthorize').val().length < 3) {
                alert('The length of the field "CSC No." should be 3 characters.');
                $('#txtAuthorize').focus();
                return false;
            }
            if ($('#ddlValidMonth').val() == '') {
                alert('Please choose the Month of Valid date.');
                $('#ddlValidMonth').focus();
                return false;
            }
            if ($('#ddlValidYear').val() == '') {
                alert('Please choose the Year of Valid date.');
                $('#ddlValidYear').focus();
                return false;
            }
            if ($('#txtCardOwner').val() == '') {
                alert('Card holder is required.');
                $('#txtCardOwner').focus();
                return false;
            }
            if ($('#txtCardOwner').val().length > 20) {
                alert('The length of the field "Card holder" can not exceed 20 characters!');
                $('#txtCardOwner').focus();
                return false;
            }


            // 有效期限驗證是否已過期 by Hilty 2013/12/18
            var today = new Date();
            var tMonth = today.getMonth() + 1;
            var tYear = today.getFullYear().toString().substring(2);

            if ($('#ddlValidYear').val() <= tYear && $('#ddlValidMonth').val() < tMonth) 
            {
                alert('Please check if the expiration date is valid!');
                return false;
            }

        }
        function Account_No_Keyup(AccountNo) {
            var Account_No1 = document.getElementById('tbxAccount_No1');
            var Account_No2 = document.getElementById('tbxAccount_No2');
            var Account_No3 = document.getElementById('tbxAccount_No3');
            var Account_No4 = document.getElementById('tbxAccount_No4');
            if (AccountNo == 1 && Account_No1.value.length == 4) {
                Account_No2.focus();
            } else if (AccountNo == 2 && Account_No2.value.length == 4) {
                Account_No3.focus();
            } else if (AccountNo == 3 && Account_No3.value.length == 4) {
                Account_No4.focus();
            }
        }
        //20140523 新增by Ian 增加信用卡 卡別和卡號的檢核機制
        function CheckCreditCardNumJ() {
            var Card_Type = document.getElementById('ddlCardType').value;
            var Account_No1 = document.getElementById('tbxAccount_No1').value;
            var Account_No2 = document.getElementById('tbxAccount_No2').value;
            var Account_No3 = document.getElementById('tbxAccount_No3').value;
            var Account_No4 = document.getElementById('tbxAccount_No4').value;
            var num = Account_No1 + Account_No2 + Account_No3 + Account_No4;
            var len = num.length;
            var sum = 0;
            var mul = 1;
            for (i = (len - 1); i >= 0; --i) {
                numAtI = parseInt(num.charAt(i));
                tempSum = numAtI * mul;
                if (tempSum > 9) {
                    first = tempSum % 10;
                    tempSum = 1 + first;
                }
                sum += tempSum;
                if (mul == 1) mul++;
                else mul--;
            }
            if (sum % 10) {
                if (len == 15) {
                    txtcard = CheckFor15(num, 0);
                    if (txtcard == 'invalid') {
                        alert('The card number is invalid!');
                        return false;
                    }
                }
                else {
                    alert('The card number is invalid!');
                    return false;
                }
            }
            switch (len) {
                case 13:
                    txtcard = CheckFor13(num);
                    break;
                case 15:
                    txtcard = CheckFor15(num, 1);
                    break;
                case 16:
                    txtcard = CheckFor16(num);
                    break;
                default:
                    txtcard = 'invalid';
            }
            if (txtcard == 'invalid') {
                alert('The card number is invalid!');
                return false;
            }
            else {
                if (Card_Type == txtcard) {
                    alert('The card type and the card number are correct!');
                    return true;
                }
                else {
                    alert('The card number is correct but the card type should be ' + txtcard);
                    return false;
                }
            }
        }
        //20140523 新增by Ian Checks for Visa
        function CheckFor13(number) {
            if (parseInt(number.charAt(0)) == 4)
                return 'VISA';
            return 'invalid';
        }
        //20140523 新增by Ian Check for American Express, enRoute, JCB
        function CheckFor15(number, chec) {
            if ((number.charAt(0) == 3) && ((number.charAt(1) == 4) || (number.charAt(1) == 7)) && chec)
                return 'AE';
            var FirstFour = parseInt(number.charAt(0)) * 1000;
            FirstFour += parseInt(number.charAt(1)) * 100;
            FirstFour += parseInt(number.charAt(2)) * 10;
            FirstFour += parseInt(number.charAt(3));
            if ((FirstFour == 2014) || (FirstFour == 2149))
                return 'enRoute';
            if (((FirstFour == 2131) || (FirstFour == 1800)) && chec)
                return 'JCB';
            return 'invalid';
        }
        //20140523 新增by Ian Check for Visa, MasterCard, Discover or JCB
        function CheckFor16(number) {
            var a = parseInt(number.charAt(0));
            var b = parseInt(number.charAt(1));
            switch (a) {
                case 5:
                    if ((b > 0) && (b < 6))
                        return 'MASTER';
                    else
                        return 'invalid';
                    break;
                case 4:
                    return 'VISA';
                    break;
                case 6:
                    if ((b == 0) && (parseInt(number.charAt(2)) == 1) && (parseInt(number.charAt(3)) == 1))
                        return 'Discover';
                    else
                        return 'invalid';
                    break;
                case 3:
                    return 'JCB';
                    break;
                default:
                    return 'invalid';
                    break;
            }
        }
    </script>
</head>
<body style="background-color: #EEE" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form id="Form1" runat="server">
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_ItemList" runat="server" />
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
                            <strong><font color="#003399">Pledge Information</font></strong>
                        </td>
                        <td align="right" valign="bottom">
                            <strong><font color="#003399">Language：</font><a href="../Online/DonateCreditCard.aspx" target="_self"><font size="3">中文</font></a></strong>
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
                <asp:Label ID="lblStep1" runat="server" Text=" Your Donation Items >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep2" runat="server" Text=" Donor Information >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep3" runat="server" Text=" Confirm Donation >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep4" runat="server" Text=" Pledge Information >> " ForeColor="Red" Font-Bold="true" />
                <asp:Label ID="lblStep5" runat="server" Text=" Confirm Payment " ForeColor="MediumBlue" />
                <p />
                            <font color="#8B0000">Dear</font> <asp:Label ID="lblTitle" runat="server" Style="color: #8B0000;"> , please enter your credit card to proceed.</asp:Label>
                <br />
                
                <table align="center" border="0" cellpadding="0" cellspacing="1" class="table_v" width="60%">
                    <tr>
                        <th align="right" style="width: 30mm">
                            <span class="necessary">Card type： </span>
                        </th>
                        <td align="left">
                            <%-- <asp:DropDownList ID="ddlCardType" runat="server">
                            </asp:DropDownList>--%>
                            <asp:Label ID="lblCardType" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <th align="right" style="width: 30mm">
                            <span class="necessary">Card holder：</span>
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtCardOwner" runat="server" Width="40mm" MaxLength="20"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <th align="right" style="width: 30mm">
                            <span class="necessary">Bank：</span>
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtCardBank" runat="server" Width="40mm" MaxLength="20"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <th align="right" style="width: 30mm">
                            <span class="necessary">Card No.： </span>
                        </th>
                        <td align="left">
                            <asp:TextBox runat="server" ID="tbxAccount_No1" CssClass="font9" Width="40px"  OnKeyup="Account_No_Keyup('1');" MaxLength="4"></asp:TextBox>-
                            <asp:TextBox runat="server" ID="tbxAccount_No2" CssClass="font9" Width="40px"  OnKeyup="Account_No_Keyup('2');" MaxLength="4"></asp:TextBox>-
                            <asp:TextBox runat="server" ID="tbxAccount_No3" CssClass="font9" Width="40px"  OnKeyup="Account_No_Keyup('3');" MaxLength="4"></asp:TextBox>-
                            <asp:TextBox runat="server" ID="tbxAccount_No4" CssClass="font9" Width="40px" onblur="CheckCreditCardNumJ();" MaxLength="4"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <th align="right" style="width: 30mm">
                            <span class="necessary">Valid date： </span>
                        </th>
                        <td align="left">
                            <asp:DropDownList ID="ddlValidMonth" runat="server">
                            </asp:DropDownList>
                            Month/
                            <asp:DropDownList ID="ddlValidYear" runat="server">
                            </asp:DropDownList>
                            Year
                        </td>
                    </tr>
                    <tr>
                        <th align="right" style="width: 30mm">
                            <span class="necessary">CVV/CSC No.： </span>
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtAuthorize" runat="server" Width="10mm" MaxLength="3"></asp:TextBox><asp:Image ID="Image_Credit3no" ImageUrl="~/images/credit3no.gif" Width="110px" runat="server" />
                        </td>
                    </tr>
                </table><br />
                <table align="center" border="0" cellpadding="0" cellspacing="1" width="70%">
                    <tr><span align='center' style='color: red; font-size: medium; font-weight:bold;'>
                        All fields are required.<br><br>
                        Please be sure about the validity of card number and expiration date on your card.
                        </span>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td height="10" colspan="2">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center">
                <asp:Button ID="Button1" class="Online Online_Button5" runat="server" Text="Previous"
                    OnClick="btnPrev_Click" Width="174px" height="40px"/>
                <asp:Button ID="Button2" class="Online Online_Button5" runat="server" Text="Next"
                    OnClientClick="return Foolproof();" OnClick="btnConfirm_Click" Width="174px" height="40px" />
            </td>
        </tr>
    </table>
    <%--<div class="function">
        <asp:Button ID="btnPrev" class="npoButton npoButton_PrevStep" runat="server" Text="上一步"
            OnClick="btnPrev_Click" height="40px"/>
        <asp:Button ID="btnConfirm" class="npoButton npoButton_Submit" runat="server" Text="完成填寫"
            OnClientClick="return Foolproof();" OnClick="btnConfirm_Click" Width="120px" height="40px" />
    </div>--%>
    </form>
</body>
</html>
