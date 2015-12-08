<%@ Page Language="C#" EnableEventValidation="false" CodeFile="DonateCreditCard.aspx.cs" Inherits="Online_DonateCreditCard" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>線上定期定額奉獻</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/jquery.address.js"></script>
    <script type="text/javascript" src="../include/jquery.field.js"></script>
    <script type="text/javascript" src="../include/jquery.maxlength.js"></script>
    <script type="text/javascript">

        $(document).ready(function () {

            $('#tbxAccount_No1,#tbxAccount_No2,#tbxAccount_No3,#tbxAccount_No4,#txtAuthorize').not('').bind("keyup", function (e) {
                var val = $(this).val();
                //if (isNaN(val)) {
                    val = val.replace(/[^0-9]/g, '');
                    $(this).val(val);
                //}
            });

            //卡片背面後3碼
            $('#txtAuthorize').bind("blur", function (e) {
                if ($('#txtAuthorize').val() == '') {
                    $('#lblAuthorizePoint').text('必填');
                    $('#lblAuthorizePoint').addClass('blink');
                }
                else {
                    if ($('#txtAuthorize').val().length == 3) {
                        $('#lblAuthorizePoint').text('');
                        $('#lblAuthorizePoint').removeClass('blink');
                    }
                    else {
                        $('#lblAuthorizePoint').text('長度需為3碼');
                        $('#lblAuthorizePoint').addClass('blink');
                    }
                }
            });

            //有效期限
            $('#ddlValidYear,#ddlValidMonth').bind("change", function (e) {
                if ($('#ddlValidMonth').val() != '' && $('#ddlValidYear').val() != '') {

                    // 有效期限驗證是否已過期
                    var today = new Date();
                    var tMonth = today.getMonth() + 1;
                    var tYear = today.getFullYear().toString().substring(2);

                    if ($('#ddlValidYear').val() <= tYear && $('#ddlValidMonth').val() < tMonth) {
                        $('#lblValidMonthYearPoint').text('過期');
                        $('#lblValidMonthYearPoint').addClass('blink');
                    }
                    else {
                        $('#lblValidMonthYearPoint').removeClass('blink');
                        $('#lblValidMonthYearPoint').text('');
                    }

                }
                else {
                    $('#lblValidMonthYearPoint').text('必填');
                    $('#lblValidMonthYearPoint').addClass('blink');
                }
            });

            //信用卡卡號
            $('#tbxAccount_No1,#tbxAccount_No2,#tbxAccount_No3,#tbxAccount_No4').bind("change", function (e) {
                $('#lblCardNoPoint').removeClass('blink');
                $('#lblCardNoPoint').text('');
            });

            $('#tbxAccount_No4').bind("blur", function (e) {
                var strCardNo = $('#tbxAccount_No1').val() + $('#tbxAccount_No2').val() + $('#tbxAccount_No3').val() + $('#tbxAccount_No4').val();
                if (strCardNo == '') {
                    $('#lblCardNoPoint').text('必填');
                    $('#lblCardNoPoint').addClass('blink');
                }
                else {
                    $('#lblCardNoPoint').removeClass('blink');
                    $('#lblCardNoPoint').text('');

                }
            });

            //發卡銀行
            $('#txtCardBank').bind("blur", function (e) {
                if ($('#txtCardBank').val() == '') {
                    $('#lblCardBankPoint').text('必填');
                    $('#lblCardBankPoint').addClass('blink');
                }
                else {
                    $('#lblCardBankPoint').removeClass('blink');
                    $('#lblCardBankPoint').text('');
                }
            });

            //持卡人姓名
            $('#txtCardOwner').bind("blur", function (e) {
                if ($('#txtCardOwner').val() == '') {
                    $('#lblCardOwnerPoint').text('必填');
                    $('#lblCardOwnerPoint').addClass('blink');
                }
                else {
                    $('#lblCardOwnerPoint').removeClass('blink');
                    $('#lblCardOwnerPoint').text('');
                }
            });

            //信用卡種類
            $('#rdoCardType1,#rdoCardType2,#rdoCardType3').bind("click", function (e) {

                $('#lblCardTypePoint').removeClass('blink');
                $('#lblCardTypePoint').text('');
                if (($('#tbxAccount_No1').val() + $('#tbxAccount_No2').val() + $('#tbxAccount_No3').val() + $('#tbxAccount_No4').val()).length > 12) {
                    $('#tbxAccount_No4').blur();
                }

            });

        });                     //end of ready()

        //回上一頁
        function btnPrev_click() {
            //history.back();
            document.forms[0].action = "DonateInfoConfirm.aspx";
            document.forms[0].submit();
            return false;
        }

        function Foolproof() {

            var error_cut = 0;

            //卡片背面後3碼
            if ($('#txtAuthorize').val() == '') {
                error_cut++;
                $('#lblAuthorizePoint').text('必填');
                $('#lblAuthorizePoint').addClass('blink');
                $("#txtAuthorize").focus();
            }
            else {
                if ($('#txtAuthorize').val().length == 3) {
                    $('#lblAuthorizePoint').text('');
                    $('#lblAuthorizePoint').removeClass('blink');
                }
                else {
                    error_cut++;
                    $('#lblAuthorizePoint').text('長度需為3碼');
                    $('#lblAuthorizePoint').addClass('blink');
                    $("#txtAuthorize").focus();
                }
            }

            //有效期限
            if ($('#ddlValidMonth').val() != '' && $('#ddlValidYear').val() != '') {

                // 有效期限驗證是否已過期
                var today = new Date();
                var tMonth = today.getMonth() + 1;
                var tYear = today.getFullYear().toString().substring(2);

                if ($('#ddlValidYear').val() <= tYear && $('#ddlValidMonth').val() < tMonth) {
                    error_cut++;
                    $('#lblValidMonthYearPoint').text('過期');
                    $('#lblValidMonthYearPoint').addClass('blink');
                    $("#ddlValidMonth").focus();
                }
                else {
                    $('#lblValidMonthYearPoint').removeClass('blink');
                    $('#lblValidMonthYearPoint').text('');
                }

            }
            else {
                error_cut++;
                $('#lblValidMonthYearPoint').text('必填');
                $('#lblValidMonthYearPoint').addClass('blink');
                $("#ddlValidMonth").focus();
            }

            //信用卡卡號
            var strCardNo = $('#tbxAccount_No1').val() + $('#tbxAccount_No2').val() + $('#tbxAccount_No3').val() + $('#tbxAccount_No4').val();
            if (strCardNo == '') {
                error_cut++;
                $('#lblCardNoPoint').text('必填');
                $('#lblCardNoPoint').addClass('blink');
                $('#tbxAccount_No1').focus();
            }
            else {
                if (strCardNo.length < 13 || strCardNo.length > 16) {
                    error_cut++;
                    $('#lblCardNoPoint').text('長度不對');
                    $('#lblCardNoPoint').addClass('blink');
                    $('#tbxAccount_No1').focus();
                }
                else {
                    $('#lblCardNoPoint').removeClass('blink');
                    $('#lblCardNoPoint').text('');
                }
            }

            //發卡銀行
            if ($('#txtCardBank').val() == '') {
                error_cut++;
                $('#lblCardBankPoint').text('必填');
                $('#lblCardBankPoint').addClass('blink');
                $('#txtCardBank').focus();
            }
            else {
                $('#lblCardBankPoint').removeClass('blink');
                $('#lblCardBankPoint').text('');
            }

            //持卡人姓名
            if ($('#txtCardOwner').val() == '') {
                error_cut++;
                $('#lblCardOwnerPoint').text('必填');
                $('#lblCardOwnerPoint').addClass('blink');
                $('#txtCardOwner').focus();
            }
            else {
                $('#lblCardOwnerPoint').removeClass('blink');
                $('#lblCardOwnerPoint').text('');
            }

            //信用卡種類
            if ($('#rdoCardType1').prop("checked") == false && $('#rdoCardType2').prop("checked") == false
                && $('#rdoCardType3').prop("checked") == false) {
                error_cut++;
                $('#lblCardTypePoint').text('必選');
                $('#lblCardTypePoint').addClass('blink');
                $('#rdoCardType1').focus();
            }
            else {

                $('#lblCardTypePoint').removeClass('blink');
                $('#lblCardTypePoint').text('');

                //信用卡卡號檢核
                if (error_cut == 0) {

                    if (CheckCreditCardNumJ()) {
                        $('#lblCardNoPoint').removeClass('blink');
                        $('#lblCardNoPoint').text('');
                    }
                    else {
                        error_cut++;
                        $('#lblCardNoPoint').text('驗證錯誤');
                        $('#lblCardNoPoint').addClass('blink');
                        $('#tbxAccount_No1').focus();
                    }
                }

            }

            if (error_cut > 0) {
                return false;
            }
            else {
                return true;
            }

        }

        //信用卡號欄位按鍵處理
        function Account_No_Keyup(AccountNo) {

            var Account_No1 = document.getElementById('tbxAccount_No1');
            var Account_No2 = document.getElementById('tbxAccount_No2');
            var Account_No3 = document.getElementById('tbxAccount_No3');
            var Account_No4 = document.getElementById('tbxAccount_No4');

            if (event.keyCode == 8) {
                if (AccountNo == 2 && Account_No2.value.length == 0) {
                    Account_No1.focus();
                } else if (AccountNo == 3 && Account_No3.value.length == 0) {
                    Account_No2.focus();
                } else if (AccountNo == 4 && Account_No4.value.length == 0) {
                    Account_No3.focus();
                }
                return false;
            }

            if (AccountNo == 1 && Account_No1.value.length == 4) {
                Account_No2.focus();
            } else if (AccountNo == 2 && Account_No2.value.length == 4) {
                Account_No3.focus();
            } else if (AccountNo == 3 && Account_No3.value.length == 4) {
                Account_No4.focus();
            }
        }

        //信用卡 卡別和卡號的檢核機制
        function CheckCreditCardNumJ() {

            var Card_Type = $("#lblCardType input:checked").val();
            var num = $('#tbxAccount_No1').val() + $('#tbxAccount_No2').val() + $('#tbxAccount_No3').val() + $('#tbxAccount_No4').val();
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
                        //alert('此卡號無效!!');
                        return false;
                    }
                }
                else {
                    //alert('此卡號無效!!');
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
                //alert('此卡號無效!!');
                return false;
            }
            else {
                if (Card_Type == txtcard) {
                    //alert('此卡別及卡號正確!!');
                    return true;
                }
                else {
                    //alert('此卡號正確!!但卡片應該是：' + txtcard);
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

        //Check for American Express, enRoute, JCB
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

        //Check for Visa, MasterCard, Discover or JCB
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
    <style type="text/css">

        body {
            font-size: 18px;
        }

        .table_v tr {
            color: black;
            height: 35px;
            font-size: 18px;
        }

        input {
            font-size: 18px;
        }

    </style>
</head>
<body onkeydown="if(event.keyCode==13) return false;" style="margin: 0px; background-color: #EEE">
    <form id="Form1" runat="server">
    <asp:HiddenField ID="HFD_DonorID" runat="server" />
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_Amount" runat="server" />
    <asp:HiddenField ID="HFD_PayType" runat="server" />
    <asp:HiddenField ID="HFD_DonorName" runat="server" />
    <asp:HiddenField ID="HFD_LoginOK" runat="server" />

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
                            <strong><font color="#003399">填寫定期定額授權資料</font></strong>
                        </td>
                        <td align="right" valign="bottom">
                            <!--<strong><font color="#003399">Language：</font><a href="../English/DonateCreditCard.aspx" target="_self"><font size="3">English</font></a></strong>-->
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
            <td colspan="3" align="center" style="white-space: nowrap;">                                                    
                <br />
                <span>我目前的位置：</span>
                <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="Red" Font-Bold="True" />
                <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="MediumBlue" />
                <br /><br />
                    <asp:Label ID="lblTitle" runat="server" Style="color: #8B0000;"/> 平安！請填寫您的信用卡資料以進行定期定額扣款授權
                <br /><br />
                
                <table border="0" cellpadding="0" cellspacing="1" class="table_v" width="80%">
                    <tr>
                        <th align="right" style="width: 200px">
                            <span class="necessary">信用卡種類： </span>
                        </th>
                        <td align="left">
                            <asp:Label ID="lblCardType" runat="server"></asp:Label>
                            <asp:Label ID="lblCardTypePoint" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <th align="right">
                            <span class="necessary">持卡人姓名：</span>
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtCardOwner" runat="server" Width="80%" MaxLength="20"></asp:TextBox>
                            &nbsp;&nbsp;<asp:Label ID="lblCardOwnerPoint" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <th align="right">
                            <span class="necessary">發卡銀行：</span>
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtCardBank" runat="server" Width="80%" MaxLength="20"></asp:TextBox>
                            &nbsp;&nbsp;<asp:Label ID="lblCardBankPoint" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <th align="right">
                            <span class="necessary">信用卡卡號： </span>
                        </th>
                        <td align="left">
                            <asp:TextBox runat="server" ID="tbxAccount_No1" CssClass="font9" Width="15%" OnKeyup="Account_No_Keyup('1');" MaxLength="4"></asp:TextBox> -
                            <asp:TextBox runat="server" ID="tbxAccount_No2" CssClass="font9" Width="15%" OnKeyup="Account_No_Keyup('2');" MaxLength="4"></asp:TextBox> -
                            <asp:TextBox runat="server" ID="tbxAccount_No3" CssClass="font9" Width="15%" OnKeyup="Account_No_Keyup('3');" MaxLength="4"></asp:TextBox> -
                            <asp:TextBox runat="server" ID="tbxAccount_No4" CssClass="font9" Width="15%" OnKeyup="Account_No_Keyup('4');" MaxLength="4"></asp:TextBox>
                            &nbsp;&nbsp;<asp:Label ID="lblCardNoPoint" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <th align="right">
                            <span class="necessary">有效期限： </span>
                        </th>
                        <td align="left">
                            <asp:DropDownList ID="ddlValidMonth" runat="server" Font-Size="18px" Width="15%">
                            </asp:DropDownList>
                            月/
                            <asp:DropDownList ID="ddlValidYear" runat="server" Font-Size="18px" Width="15%">
                            </asp:DropDownList>
                            年
                            &nbsp;&nbsp;<asp:Label ID="lblValidMonthYearPoint" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <th align="right" style="width: 30mm">
                            <span class="necessary">卡片背面後3碼： </span>
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtAuthorize" runat="server" Width="15%" MaxLength="3"></asp:TextBox><asp:Image ID="Image_Credit3no" ImageUrl="~/images/credit3no.gif" Width="110px" runat="server" />
                            &nbsp;&nbsp;<asp:Label ID="lblAuthorizePoint" runat="server"></asp:Label>
                        </td>
                    </tr>
                </table><br />
                <table border="0" cellpadding="0" cellspacing="1">
                    <tr>
                        <td align='center' colspan="2"><span style='color: blue; font-weight: bold;'>以上每一個欄位都是必填，請您填入有效之信用卡資訊。<br />
                            <br />請確認卡片有效日期及正確卡號。<br />&nbsp;</span></td>
                    </tr>
                    <tr>
                        <td><span style='color: red; font-weight: bold;'>我們使用SSL加密機制來保護傳輸您的信用卡資料，<br />並安全地儲存在我們嚴密防護的系統中。</span></td>
                        <td>
                            <script type="text/javascript" src="https://seal.godaddy.com/getSeal?sealID=DHpQNoGOKhhCpPbW2AxeTcxVj0p7gI3Yq0VXnSszQ1S6uUuKXOEwGDgvA"></script>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td height="30" colspan="2">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center">
                <table width="80%">
                    <tr>
                        <td>
                            <asp:Button ID="btnPrev" class="css_btn_class" runat="server" Text="回上一頁" OnClientClick="btnPrev_click();" />
                        </td>
                        <td align="right">
                            <asp:Button ID="btnConfirm" class="css_btn_class" runat="server" Text="完成授權填寫" OnClientClick="return Foolproof();" OnClick="btnConfirm_Click" />
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
