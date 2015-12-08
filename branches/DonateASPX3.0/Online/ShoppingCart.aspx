<%@ Page Language="C#" EnableEventValidation="false" CodeFile="ShoppingCart.aspx.cs" Inherits="Online_ShoppingCart" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>我的奉獻明細</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript">

        $(document).ready(function () {


            $("#hlCode").click(function () {

                document.forms[0].action = "VerificationCode.aspx";
                document.forms[0].submit();
                return false;

            });

            $("#hlAddlogin").click(function () {

                $("#HFD_AddNew").val('Y');
                $("#HFD_DonorID").val('');
                $('#HFD_LoginOK').val('');
                document.forms[0].action = "CheckOut.aspx";
                document.forms[0].submit();
                return false;

            });
            //20151029 拿掉「不註冊，我直接奉獻」的匿名功能
            //$("#hlNologin").click(function () {

            //    $("#HFD_AddNew").val('N');
            //    $("#HFD_DonorID").val('');
            //    $('#HFD_LoginOK').val('');
            //    document.forms[0].action = "CheckOut.aspx";
            //    document.forms[0].submit();
            //    return false;

            //});

            $("#ddlPeriod").change(function () {

                $("#HFD_PayType").val($("#ddlPeriod").val());

            });

            $("#btnContinue").click(function () {

                $("#HFD_AddNew").val('');
                $("#HFD_DonorID").val('');
                $('#HFD_LoginOK').val('');
                document.forms[0].action = "DonateOnlineAll.aspx";
                document.forms[0].submit();
                return false;

            });

            //帳號(Email)
            $('#txtAccount').blur(function () {
                $('#lblAccountPwdPoint').text('');
                $('#lblAccountPwdPoint').removeClass('blink');
                if ($('#txtAccount').val() == "") {
                    $('#lblAccountPoint').text('必填');
                    $('#lblAccountPoint').addClass('blink');
                }
                else {

                    //檢查Email格式
                    var patterns = /^[a-zA-Z0-9_-].+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$/;
                    if (!patterns.test($("#txtAccount").val())) {
                        $('#lblAccountPoint').text('格式錯誤');
                        $('#lblAccountPoint').addClass('blink');
                    }
                    else {
                        $('#lblAccountPoint').text('');
                        $('#lblAccountPoint').removeClass('blink');
                    }
                }
            });

            //密碼
            $('#txtPwd').blur(function () {
                $('#lblAccountPwdPoint').text('');
                $('#lblAccountPwdPoint').removeClass('blink');
                if ($('#txtPwd').val() == "") {
                    $('#lblPwdPoint').text('必填');
                    $('#lblPwdPoint').addClass('blink');
                }
                else {
                    $('#lblPwdPoint').text('');
                    $('#lblPwdPoint').removeClass('blink');
                }
            });


        });   //end of ready()


        function CheckLogin() {

            var error_cut = 0;

            //密碼
            if ($('#txtPwd').val() == "") {
                error_cut++;
                $('#lblPwdPoint').text('必填');
                $('#lblPwdPoint').addClass('blink');
                $('#txtPwd').focus();
            }
            else {
                $('#lblPwdPoint').text('');
                $('#lblPwdPoint').removeClass('blink');
            }

            //帳號(Email)
            if ($('#txtAccount').val() == "") {
                error_cut++;
                $('#lblAccountPoint').text('必填');
                $('#lblAccountPoint').addClass('blink');
                $('#txtAccount').focus();
            }
            else {
                $('#lblAccountPoint').text('');
                $('#lblAccountPoint').removeClass('blink');
            }

            //登入帳號
            if ($('#txtAccount').val() != "" && $('#txtPwd').val() != "") {

                var result_ok = "";
                $.ajax({
                    type: 'post',
                    url: "../common/ajax.aspx",
                    data: 'Type=11' + '&txtAccount=' + $('#txtAccount').val() + '&txtPwd=' + $('#txtPwd').val(),
                    async: false, //同步
                    success: function (result) {
                        result_ok = result;
                    },
                    error: function () { alert('ajax failed'); }
                });

                if (result_ok != "") {
                    $('#lblAccountPwdPoint').text('');
                    $('#lblAccountPwdPoint').removeClass('blink');
                    $('#HFD_DonorID').val(result_ok);
                    $('#HFD_LoginOK').val('ok');
                }
                else {
                    error_cut++;
                    $('#lblAccountPwdPoint').text('帳號(Email)或密碼不對');
                    $('#lblAccountPwdPoint').addClass('blink');
                    $('#txtPwd').focus();
                }
            }
            else {
                $('#lblAccountPwdPoint').text('');
                $('#lblAccountPwdPoint').removeClass('blink');
            }

            if (error_cut > 0) {
                return false;
            }
            else {
                document.forms[0].action = "CheckOut.aspx";
                document.forms[0].submit();
            }
            return false;

        }

        //在密碼欄位按Enter執行登入帳號
        function PwdSubmit() {
            if ($('#txtPwd').val() == "") return false;
            if (event.keyCode == 13) CheckLogin();
        }

    </script>
    <style type="text/css">
        body {
            background-color: #EEE;
            text-align: center;
            font-size: 18px;
        }

        td {
            color: black;
        }

        .table_h tr:hover {
            color: black;
            background-color: white;
        }

        select {
            font-size: 18px;
        }

</style>
</head>
<body onkeydown="if(event.keyCode==13) return false;" style="margin:0px;">
    <form id="form1" runat="server">
    <asp:HiddenField ID="HFD_DonorID" runat="server" />
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_Amount" runat="server" />
    <asp:HiddenField ID="HFD_PayType" runat="server" />
    <asp:HiddenField ID="HFD_AddNew" runat="server" />
    <asp:HiddenField ID="HFD_LoginOK" runat="server" />
    <table width="860" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td colspan="2">
                &nbsp;
            </td>
        </tr>       
        <tr>
            <td width="436">
                <img alt="aa" src="images/bar_logo.jpg" width="390" height="40" />
            </td>
            <td width="424" style="background-image:url('images/bar_item.jpg')">
                <table width="95%" border="0" align="center" cellpadding="5" cellspacing="5">
                    <tr>
                        <td align="left" valign="bottom">
                            <strong><font color="#003399">我的奉獻明細</font></strong>
                        </td>
                        <td align="right" valign="bottom">
                            <!--<strong><font color="#003399">Language：</font><a href="../English/ShoppingCart.aspx" target="_self"><font size="3">English</font></a></strong>-->
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td align="left" valign="bottom" height="30" colspan="2">
			</td>
        </tr>
        <tr valign="top">
            <td colspan="2" align="center">
                <table>
                    <tr>
                        <td align="center" style="white-space: nowrap;">
                            <span>我目前的位置：</span>
                            <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="Red" Font-Bold="True" />
                            <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="MediumBlue" />
                            <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="MediumBlue" />
                            <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="MediumBlue" />
                            <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="MediumBlue" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <br />
                            <asp:Label ID="lblGrid" runat="server"></asp:Label>
                            <br />
                            <asp:Label ID="PeriodMemo" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td align="center">
                            <table>
                                <tr>
                                    <td align="center">
                                        <div style="text-align: left; font-weight: bold;">※ 輸 入 您 的 帳 號</div>
                                        <table style="width: 100%; border: #6CC solid 1px;" bgcolor="White">
                                            <tr style="height: 30px; color: #666666">
                                                <td style="width: 200px;" align="right">帳號(Email)：
                                                </td>
                                                <td style="white-space: nowrap;">
                                                    <asp:TextBox ID="txtAccount" runat="server" Width="400px" Font-Size="18px"></asp:TextBox>
                                                    <asp:Label ID="lblAccountPoint" runat="server"></asp:Label>
                                                    <asp:Label ID="lblAccountPwdPoint" runat="server"></asp:Label><asp:Button ID="btnLogin" runat="server" CssClass="css_btn_class_login" Text="帳號登入" OnClientClick="return CheckLogin();" />
                                                </td>
                                            </tr>
                                            <tr style="height: 30px; color: #666666">
                                                <td align="right">密 &nbsp; 碼：
                                                </td>
                                                <td style="white-space: nowrap;">
                                                    <asp:TextBox ID="txtPwd" runat="server" Width="400px" TextMode="Password" OnKeyup="PwdSubmit();" Font-Size="18px"></asp:TextBox>
                                                    <asp:Label ID="lblPwdPoint" runat="server"></asp:Label>
                                                    <asp:HyperLink ID="hlCode" runat="server" NavigateUrl="#" ToolTip="從Email得知驗證碼，設定密碼。">忘記密碼</asp:HyperLink>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2" align="center"></td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left">
                                        <span>使用帳號登入，可以自動幫您帶出您的基本資料。</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left">
                                        <span>原有網路奉獻編號(AG開頭)已暫停使用，現已全面改用以您的Email為帳號。如您尚未註冊Email帳號，歡迎您註冊。</span>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td colspan="2">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2">
                <table width="100%">
                    <tr>
                        <td align="left">
                            <asp:Button ID="btnContinue" runat="server" class="css_btn_class" Text="回上一頁" />
                        </td>
                        <td align="right">
                            <asp:Button ID="hlAddlogin" runat="server" CssClass="css_btn_class" Text="我從未註冊過，要建立帳號" />
                            &nbsp;
                            <%-- <asp:Button ID="hlNologin" runat="server" CssClass="css_btn_class" Text="不註冊，我直接奉獻" />
                            &nbsp;--%>
                            
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
