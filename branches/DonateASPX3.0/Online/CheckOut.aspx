<%@ Page Language="C#" EnableEventValidation="false" CodeFile="CheckOut.aspx.cs" Inherits="Online_CheckOut" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>線上奉獻資料填寫</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">

        $(document).ready(function () {

            //帳號(E-mail),輸入密碼欄位顯示或隱藏
            if ($("#HFD_DonorID").val() != "") {
                //已登入帳號
                $('.trPassword').hide();
                $('#spEmail').hide();
                $('#spEmail_Member').show();
            }
            else if ($("#HFD_AddNew").val() == "Y") {
                //尚未註冊過，新增帳號
                $('.trPassword').show();
                $('#spEmail').hide();
                $('#spEmail_Member').show();
                $('#ConfirmPwd').addClass('necessary');
            }
            else {
                //不用註冊，直接奉獻
                $('.trPassword').hide();
                $('#spEmail').show();
                $('#spEmail_Member').hide();
            }
            //20150519 使用ajax取得此Donor_ID是否已有儲存收據寄送方式
            var InvoiceTypeIsExist_ok = InvoiceTypeIsExist($.trim($('#HFD_DonorID').val()));
            if (InvoiceTypeIsExist_ok) {
                $('.trInvoiceType').show();
                $('#lblReceipt').text('收據是重要報稅憑證，一個年度請單選一種方式，若要變更請完成此項捐款奉獻後來電02-8024-3911或e-mail通知。');
                $('#lblReceipt').addClass('blink');
            }
            else {
                $('.trInvoiceType').show();
            }

            //寄發收據欄位 不寄
            if ($('#HFD_Receipt').val() == "不寄") {
                $('.trTitle').hide();
                //$('.trAddressType').hide();
                //$('.trLocal').hide();
                //$('.trOverseas').hide();
                //$('.trSubscribe').hide();
                //$('.trIsFdc').hide();
            }
            else {

                //居住地區(海外)預設隱藏
                if ($('#HFD_LiveRegion').val() == "台灣") {
                    $('.trTitle').show();
                    $('.trAddressType').show();
                    $('.trLocal').show();
                    $('.trOverseas').hide();
                    $('.trSubscribe').show();
                    $('.trIsFdc').show();
                }
                else {
                    $('.trTitle').hide();
                    $('.trAddressType').show();
                    $('.trLocal').hide();
                    $('.trOverseas').show();
                    $('.trSubscribe').hide();
                    $('.trIsFdc').hide();
                }

            }

            //寄發收據欄位
            //年度證明
            $('#rdoReceipt1').bind("click", function (e) {
                $('#HFD_Receipt').val("年度證明");
                if ($('#HFD_LiveRegion').val() == "台灣") {
                    $('.trTitle').show();
                    $('.trAddressType').show();
                    $('.trLocal').show();
                    $('.trOverseas').hide();
                    $('.trSubscribe').show();
                    $('.trIsFdc').show();
                }
                else {
                    $('.trTitle').hide();
                    $('.trAddressType').show();
                    $('.trLocal').hide();
                    $('.trOverseas').show();
                    $('.trSubscribe').hide();
                    $('.trIsFdc').hide();
                }
            });

            //逐次寄發
            $('#rdoReceipt2').bind("click", function (e) {
                $('#HFD_Receipt').val("逐次寄發");
                if ($('#HFD_LiveRegion').val() == "台灣") {
                    $('.trTitle').show();
                    $('.trAddressType').show();
                    $('.trLocal').show();
                    $('.trOverseas').hide();
                    $('.trSubscribe').show();
                    $('.trIsFdc').show();
                }
                else {
                    $('.trTitle').hide();
                    $('.trAddressType').show();
                    $('.trLocal').hide();
                    $('.trOverseas').show();
                    $('.trSubscribe').hide();
                    $('.trIsFdc').hide();
                }
            });

            //不寄
            $('#rdoReceipt3').bind("click", function (e) {
                $('#HFD_Receipt').val("不寄");
                $('.trTitle').hide();
                //$('.trAddressType').hide();
                //$('.trLocal').hide();
                //$('.trOverseas').hide();
                //$('.trSubscribe').hide();
                //$('.trIsFdc').hide();
            });

            //初始化日曆
            initCalendar(['Birthday']);

            //地址下拉選項
            //縣市
            $('#ddlLiveCity').bind('change', function (e) {
                ChangeArea($(e.target).val(), 'ddlLiveArea');
                $('#HFD_LiveArea').val('');
                $('#txtLiveZip').val('');
            });

            //鄉鎮市區
            $('#ddlLiveArea').bind("change", function (e) {
                var ZipCode = $(this).val();
                if (ZipCode != '0') {
                    $('#txtLiveZip').val(ZipCode.substring(0, 3));
                    $('#HFD_LiveArea').val(ZipCode);
                }
            });

            //限制輸入數字
            $('#txtPhoneCode,#txtCorpPhone,#txtExt,#txtCellPhone,#txtFaxCode,#txtFax,#txtLiveZip,#txtLane,#txtAlley,#txtHouseNoSub,#txtFloor,#txtFloorSub').not('').bind("keyup", function (e) {
                var val = $(this).val();
                //if (isNaN(val)) {
                    val = val.replace(/[^0-9]/g, '');
                    $(this).val(val);
                //}
            });

            //性別
            $('input:radio[name=RDO_Gender]').bind("click", function (e) {
                $('#lblGenderPoint').removeClass('blink');
                $('#lblGenderPoint').text('');
            });

            //同意上傳國稅局申報
            $('input:radio[name=RDO_IsFdc]').bind("click", function (e) {
                if ($(this).val() == '是') {
                    //$('#spanIDNo').addClass('necessary');
                }
                else {
                    //$('#spanIDNo').removeClass('necessary');
                    $('#lblIDNoPoint').removeClass('blink');
                    $('#lblIDNoPoint').text('');
                }

            });

            //若勾選【居住地區】台灣，顯示trLocal，反之則顯示trOverseas
            $('input:radio[name=RDO_LiveRegion]').bind("click", function (e) {
                if ($(this).val() == '台灣') {
                    $('#HFD_LiveRegion').val("台灣");
                    $('.trLocal').show();
                    $('.trOverseas').hide();
                    $('.trSubscribe').show();
                    $('.trIsFdc').show();
                }
                else {
                    $('#HFD_LiveRegion').val("海外");
                    $('.trLocal').hide();
                    $('.trOverseas').show();
                    $('.trSubscribe').hide();
                    $('.trIsFdc').hide();
                }
            });

            $('#txtDonorName').blur(function () {
                if ($('#txtDonorName').val() == "") {
                    $('#lblDonorNamePoint').text('必填');
                    $('#lblDonorNamePoint').addClass('blink');
                }
                else {
                    $('#lblDonorNamePoint').text('');
                    $('#lblDonorNamePoint').removeClass('blink');
                }
            });

            $('#txtEmail').blur(function () {

                if ($('#HFD_LoginOK').val() != "ok") {

                    $('#lblEmailPoint').removeClass('blink');
                    if (!$.trim($('#txtEmail').val()).length) {
                        $('#lblEmailPoint').text('必填');
                        $('#lblEmailPoint').addClass('blink');
                    } else {

                        //檢查Email格式
                        var patterns = /^[a-zA-Z0-9_-].+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$/;
                        if (!patterns.test($("#txtEmail").val())) {
                            $('#lblEmailPoint').text('格式錯誤');
                            $('#lblEmailPoint').addClass('blink');
                        }
                        else {

                            // 2015/7/8 從"不註冊，我直接奉獻"進入，移除直接奉獻內對email帳號的檢核機制(保留格式檢查)
                            if ($("#HFD_AddNew").val() == "Y") {
                                //使用ajax取得email是否存在
                                var emailIsExist_ok = emailIsExist($.trim($('#txtEmail').val()));
                                if (emailIsExist_ok) {
                                    $('#lblEmailPoint').text('已註冊');
                                    $('#lblEmailPoint').addClass('blink');
                                }
                                else {
                                    $('#lblEmailPoint').text('');
                                    $('#lblEmailPoint').removeClass('blink');
                                }
                            }
                            else {
                                $('#lblEmailPoint').text('');
                                $('#lblEmailPoint').removeClass('blink');
                            }

                        }
                    }

                }

            });

            //道路街名或村名
            $('#txtLiveStreet').blur(function () {
                if ($('#txtLiveStreet').val() == "") {
                    $('#lblLiveStreetPoint').text('必填');
                    $('#lblLiveStreetPoint').addClass('blink');
                }
                else {
                    $('#lblLiveStreetPoint').text('');
                    $('#lblLiveStreetPoint').removeClass('blink');
                }
            });

            //國家/省/城市/區
            //$('#txtOverseasCountry').blur(function () {
            //    if ($('#txtOverseasCountry').val() == "") {
            //        $('#lblOverseasCountryPoint').text('必填');
            //        $('#lblOverseasCountryPoint').addClass('blink');
            //    }
            //    else {
            //        $('#lblOverseasCountryPoint').text('');
            //        $('#lblOverseasCountryPoint').removeClass('blink');
            //    }
            //});
            //國家下拉選單
            $('#ddlOverseasCountry').blur(function () {
                if ($('#ddlOverseasCountry').val() == "") {
                    $('#lblOverseasCountryPoint').text('必填');
                    $('#lblOverseasCountryPoint').addClass('blink');
                }
                else {
                    $('#lblOverseasCountryPoint').text('');
                    $('#lblOverseasCountryPoint').removeClass('blink');
                }
            });

            //地址
            $('#txtOverseasAddress').blur(function () {
                if ($('#txtOverseasAddress').val() == "") {
                    $('#lblOverseasAddressPoint').text('必填');
                    $('#lblOverseasAddressPoint').addClass('blink');
                }
                else {
                    $('#lblOverseasAddressPoint').text('');
                    $('#lblOverseasAddressPoint').removeClass('blink');
                }
            });

            //身份證字號
            $('#txtIDNo').blur(function () {
                if ($('#txtIDNo').val() != '') {
                    if (checkID($('#txtIDNo').val())) {
                        $('#lblIDNoPoint').text('');
                        $('#lblIDNoPoint').removeClass('blink');
                    }
                    else {
                        $('#lblIDNoPoint').text('驗證錯誤');
                        $('#lblIDNoPoint').addClass('blink');
                    }
                }
                else {

                    //同意上傳國稅局申報 必須填寫身份證字號
                    if ($('#rdoIsFdc1').prop("checked") == true) {
                        $('#lblIDNoPoint').text('必填');
                        $('#lblIDNoPoint').addClass('blink');
                    }
                    else {
                        $('#lblIDNoPoint').removeClass('blink');
                        $('#lblIDNoPoint').text('');
                    }

                }
            });

            //回上一頁
            $('#btnContinue').bind("click", function (e) {
                //history.back();
                document.forms[0].action = "ShoppingCart.aspx";
                document.forms[0].submit();
                return false;
            });

            /*
            //按上一頁回來的檢查
            if ($.trim($('#txtEmail').val()).length) {
                var emailIsExist_ok = emailIsExist($.trim($('#txtEmail').val()));
                if (emailIsExist_ok) {
                    $('#lblEmailIsExist').text('');
                }
                else
                    $('#lblEmailIsExist').text('可使用');

            }
            */

        });                   //end of ready()

        //以 CityID 取得其鄉鎮列表
        function ChangeArea(e, ddlArea) {
            $.ajax({
                type: 'post',
                url: "../common/ajax.aspx",
                data: 'Type=1' + '&CityID=' + e,
                success: function (result) {
                    $('select[id$=' + ddlArea + ']').html(result);
                    if (result == "<option value=''></option>") {
                        $('#lblLiveCityPoint').text('必填');
                        $('#lblLiveCityPoint').addClass('blink');
                    }
                    else {
                        $('#lblLiveCityPoint').removeClass('blink');
                        $('#lblLiveCityPoint').text('');
                    }
                },
                error: function () { alert('ajax failed'); }
            })
        } //end of ChangeArea()

        function checkInvalid() {

            var error_cut = 0;
            //20150511 新增 想對GOOD TV說的話 欄位長度防呆
            //var cnt = 0;
            //var txtToGoodTV = document.getElementById('txtToGoodTV');
            //下一步前的檢查
            //姓名
            if ($('#txtDonorName').val() == '') {
                error_cut++;
                $('#lblDonorNamePoint').text('必填');
                $('#lblDonorNamePoint').addClass('blink');
            }
            else {
                $('#lblDonorNamePoint').text('');
                $('#lblDonorNamePoint').removeClass('blink');
            }

            //性別
            if ($('#rdoSex1').prop("checked") == false && $('#rdoSex2').prop("checked") == false) {
                error_cut++;
                $('#lblGenderPoint').text('必填');
                $('#lblGenderPoint').addClass('blink');
            }
            else {
                $('#lblGenderPoint').text('');
                $('#lblGenderPoint').removeClass('blink');
            }

            //帳號(E-mail) for 未登入帳號
            if ($('#HFD_LoginOK').val() != "ok") {

                if (!$.trim($('#txtEmail').val()).length) {
                    error_cut++;
                    $('#lblEmailPoint').text('必填');
                    $('#lblEmailPoint').addClass('blink');
                } else {

                    //檢查Email格式
                    var patterns = /^[a-zA-Z0-9_-].+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$/;
                    if (!patterns.test($("#txtEmail").val())) {
                        error_cut++;
                        $('#lblEmailPoint').text('格式錯誤');
                        $('#lblEmailPoint').addClass('blink');
                        $("#txtEmail").focus();
                    }
                    else {

                        // 2015/7/8 從"不註冊，我直接奉獻"進入，移除直接奉獻內對email帳號的檢核機制(保留格式檢查)
                        if ($("#HFD_AddNew").val() == "Y") {
                            //使用ajax取得email是否存在
                            var emailIsExist_ok = emailIsExist($.trim($('#txtEmail').val()));
                            if (emailIsExist_ok) {
                                error_cut++;
                                $('#lblEmailPoint').text('已註冊');
                                $('#lblEmailPoint').addClass('blink');
                                $("#txtEmail").focus();
                            }
                            else {
                                $('#lblEmailPoint').text('');
                                $('#lblEmailPoint').removeClass('blink');
                            }
                        }
                    }
                }

            }

            //密碼 for 尚未註冊過，新增帳號
            if ($('#HFD_AddNew').val() == "Y") {
                //密碼
                if ($('#txtPassword').val() == "") {
                    error_cut++;
                    $('#lblPasswordPoint').text('必填');
                    $('#lblPasswordPoint').addClass('blink');
                    $('#txtPassword').focus();
                }
                else {
                    $('#lblPasswordPoint').text('');
                    $('#lblPasswordPoint').removeClass('blink');
                }
                //確認密碼
                if ($('#txtConfirmPwd').val() == "") {
                    error_cut++;
                    $('#lblConfirmPwdPoint').text('必填');
                    $('#lblConfirmPwdPoint').addClass('blink');
                    $('#txtConfirmPwd').focus();
                }
                else {
                    $('#lblConfirmPwdPoint').text('');
                    $('#lblConfirmPwdPoint').removeClass('blink');
                }
                //密碼與確認密碼比對
                if ($('#txtPassword').val() != "" && $('#txtConfirmPwd').val() != "") {
                    if ($('#txtPassword').val() == $('#txtConfirmPwd').val()) {
                        $('#lblConfirmPwdPoint').text('');
                        $('#lblConfirmPwdPoint').removeClass('blink');
                    }
                    else {
                        error_cut++;
                        $('#lblPasswordPoint').text('與確認密碼必須一致');
                        $('#lblPasswordPoint').addClass('blink');
                        $('#lblConfirmPwdPoint').text('必須一致');
                        $('#lblConfirmPwdPoint').addClass('blink');
                        $('#txtPassword').focus();
                    }
                }
            }

            // 若有身分證字號就檢核
            if ($('#txtIDNo').val() != '') {
                if (checkID($('#txtIDNo').val())) {
                    $('#lblIDNoPoint').text('');
                    $('#lblIDNoPoint').removeClass('blink');
                }
                else {
                    error_cut++;
                    $('#lblIDNoPoint').text('驗證錯誤');
                    $('#lblIDNoPoint').addClass('blink');
                }
            }
            else {

                //同意上傳國稅局申報 必須填寫身份證字號
                if ($('#rdoIsFdc1').prop("checked") == true) {
                    error_cut++;
                    $('#lblIDNoPoint').text('必填');
                    $('#lblIDNoPoint').addClass('blink');
                }
                else {
                    $('#lblIDNoPoint').removeClass('blink');
                    $('#lblIDNoPoint').text('');
                }

            }

            //居住地區(RdoLiveRegion1=台灣/RdoLiveRegion2=海外地區)
            if ($('#RdoLiveRegion1').prop('checked') == true) {
                //寄發收據(開立收據:rdoReceipt1=年度證明/rdoReceipt2=逐次寄發/rdoReceipt3=不寄)
                if ($('#rdoReceipt1').prop("checked") == true || $('#rdoReceipt2').prop("checked") == true) {
                    if ($('#ddlLiveCity').val() == "") {
                        error_cut++;
                        $('#lblLiveCityPoint').text('必填');
                        $('#lblLiveCityPoint').addClass('blink');
                    }
                    else {
                        $('#lblLiveCityPoint').removeClass('blink');
                        $('#lblLiveCityPoint').text('');
                    }
                    if ($('#txtLiveStreet').val() == "") {
                        error_cut++;
                        $('#lblLiveStreetPoint').text('必填');
                        $('#lblLiveStreetPoint').addClass('blink');
                    }
                    else {
                        $('#lblLiveStreetPoint').removeClass('blink');
                        $('#lblLiveStreetPoint').text('');
                    }
                }

                //號
                //if ($('#txtHouseNo').val() == "" && $('#txtHouseNoSub').val() != "") {
                //    error_cut++;
                //    $('#lblHousePoint').text('幾之幾號');
                //    $('#lblHousePoint').addClass('blink');
                //    $('#sHouse').addClass('blink');
                //}
                //else {
                //    $('#lblHousePoint').removeClass('blink');
                //    $('#lblHousePoint').text('');
                //    $('#sHouse').removeClass('blink');
                //}

                //樓
                if ($('#txtFloor').val() == "" && $('#txtFloorSub').val() != "") {
                    error_cut++;
                    $('#lblFloorPoint').text('幾樓之幾');
                    $('#lblFloorPoint').addClass('blink');
                    $('#sFloor').addClass('blink');
                }
                else {
                    $('#lblFloorPoint').removeClass('blink');
                    $('#lblFloorPoint').text('');
                    $('#sFloor').removeClass('blink');
                }

            }
            else {

                if ($('#ddlOverseasCountry').val() == "") {
                    error_cut++;
                    $('#lblOverseasCountryPoint').text('必填');
                    $('#lblOverseasCountryPoint').addClass('blink');
                }
                else {
                    $('#lblOverseasCountryPoint').text('');
                    $('#lblOverseasCountryPoint').removeClass('blink');
                }
                if ($('#txtOverseasAddress').val() == "") {
                    error_cut++;
                    $('#lblOverseasAddressPoint').text('必填');
                    $('#lblOverseasAddressPoint').addClass('blink');
                }
                else {
                    $('#lblOverseasAddressPoint').text('');
                    $('#lblOverseasAddressPoint').removeClass('blink');
                }

            }
            //cnt = txtToGoodTV.value.length;
            //if (cnt > 50) {
            //    alert('想對GOOD TV說的話 欄位長度超過限制！');
            //    return false;
            //}
            //確認是否有錯誤
            if (error_cut > 0)
                return false;
            else {

                //alert('ok');
                //return false;
                return true;
            }

        }

        //檢核Email帳號是否已存在
        function emailIsExist(Email) {
            var result_ok = false;
            $.ajax({
                type: 'post',
                url: "../common/ajax.aspx",
                data: 'Type=4' + '&Email=' + Email,
                async: false, //同步
                success: function (result) {
                    result_ok = (result == "Y") ? true : false;
                },
                error: function () { alert('ajax failed'); }
            });
            return result_ok;
        }

        //20150519 檢核捐款人收據寄送方式是否已存在
        function InvoiceTypeIsExist(DonorId) {
            var result_ok = false;
            $.ajax({
                type: 'post',
                url: "../common/ajax.aspx",
                data: 'Type=21' + '&DonorId=' + DonorId,
                async: false, //同步
                success: function (result) {
                    result_ok = (result == "Y") ? true : false;
                },
                error: function () { alert('ajax failed'); }
            });
            return result_ok;
        }

    </script>
    <style type="text/css">
        .table_v2 tr {
            font-size: 16px;
        }

        .table_v2 th {
            line-height: 30px;
        }

        .table_v2 td {
            line-height: 30px;
        }

        input {
            font-size: 16px;
        }

        select {
            font-size: 16px;
        }
    </style>
</head>
<body style="background-color: #EEE" onkeydown="if(event.keyCode==13) return false;">
    <form id="Form1" runat="server">
        <asp:HiddenField ID="HFD_DonorID" runat="server" />
        <asp:HiddenField ID="HFD_chkItem" runat="server" />
        <asp:HiddenField ID="HFD_Amount" runat="server" />
        <asp:HiddenField ID="HFD_PayType" runat="server" />
        <asp:HiddenField ID="HFD_AddNew" runat="server" />

        <asp:HiddenField ID="HFD_LiveArea" runat="server" />
        <asp:HiddenField ID="HFD_Gender" runat="server" />
        <asp:HiddenField ID="HFD_Receipt" runat="server" />
        <asp:HiddenField ID="HFD_Anony" runat="server" />
        <asp:HiddenField ID="HFD_Subscribe" runat="server" />
        <asp:HiddenField ID="HFD_IsFdc" runat="server" />
        <asp:HiddenField ID="HFD_LiveRegion" runat="server" />
        <asp:HiddenField ID="HFD_LoginOK" runat="server" />
        <table width="860" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td width="436">
                    <img alt="aa" src="images/bar_logo.jpg" width="390" height="40" />
                </td>
                <td width="424" style="background-image: url('images/bar_item.jpg')">
                    <table width="95%" border="0" align="center" cellpadding="5" cellspacing="5">
                        <tr>
                            <td align="left" valign="bottom">
                                <strong><font color="#003399">填寫您的基本資料</font></strong>
                            </td>
                            <td align="right" valign="bottom">
                                <!--<strong><font color="#003399">Language：</font><a href="../English/CheckOut.aspx" target="_self"><font size="3">English</font></a></strong>-->
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr valign="top">
                <td colspan="2" align="center" style="white-space: nowrap;">
                    <br />
                    <span style="font-size: 18px;">我目前的位置：
                    <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="MediumBlue" />
                        <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="Red" Font-Bold="True" />
                        <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="MediumBlue" />
                        <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="MediumBlue" />
                        <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="MediumBlue" />
                    </span></td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <br />
                </td>
            </tr>
        </table>
        <table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td>
                    <asp:Label ID="Label3" runat="server" Text=" ※代表必填欄位 " ForeColor="Red" Font-Bold="true" />
                </td>
            </tr>
            <tr>
                <td>
                    <table border="0" align="center" cellpadding="0" cellspacing="1" class="table_v2">
                        <tr>
                            <th align="right">
                                <span class="necessary">※姓名：</span>
                            </th>
                            <td>
                                <asp:TextBox ID="txtDonorName" runat="server" Width="80%" MaxLength="100"></asp:TextBox>
                                <asp:Label ID="lblDonorNamePoint" runat="server"></asp:Label>
                            </td>
                            <th align="right">
                                <span class="necessary">※性別：</span>
                            </th>
                            <td>
                                <asp:Label ID="lblGender" runat="server"></asp:Label>
                                <asp:Label ID="lblGenderPoint" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <th align="right">
                                <%--20130911Modify by GoodTV-Tanya:E-mail為必填--%>
                                <span id="spEmail" class="necessary">※E-mail：</span>
                                <span id="spEmail_Member" class="necessary" style="display: none">※帳號(E-mail)：</span>
                            </th>
                            <td style="white-space: nowrap;">
                                <asp:TextBox ID="txtEmail" runat="server" Width="80%" MaxLength="80"></asp:TextBox>
                                <asp:Label ID="lblEmailPoint" runat="server"></asp:Label>
                            </td>
                            <th align="right">生日：
                            </th>
                            <td>
                                <asp:TextBox ID="txtBirthday" Width="25mm" runat="server" onchange="CheckDateFormat(this, '生日'); ">
                                </asp:TextBox>
                                <img id="imgBirthday" alt="" src="../images/date.gif" />
                                <span style="color: blue; font-weight: bold;">格式：yyyy/mm/dd</span>
                            </td>
                        </tr>
                        <tr class="trPassword" style="display: none;">
                            <th align="right">
                                <span class="necessary">※密碼：</span>
                            </th>
                            <td>
                                <asp:TextBox ID="txtPassword" runat="server" Width="50mm" TextMode="Password" MaxLength="50"></asp:TextBox>
                                <asp:Label ID="lblPasswordPoint" runat="server"></asp:Label>
                            </td>
                            <th align="right">
                                <span id="ConfirmPwd" class="necessary">※確認密碼：</span>
                            </th>
                            <td>
                                <asp:TextBox ID="txtConfirmPwd" runat="server" Width="50mm" TextMode="Password" MaxLength="50"></asp:TextBox>
                                <asp:Label ID="lblConfirmPwdPoint" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr class="trInvoiceType" style="display: none;">
                            <th align="right">
                                <span class="necessary">※寄發收據：</span>
                            </th>
                            <td colspan="3">
                                <asp:Label ID="lblReceipt" runat="server"></asp:Label>
                            </td>
                            
                        </tr>
                        <tr class="trTitle" style="display: none;">
                            <th align="right">收據抬頭：
                            </th>
                            <td colspan="3">
                                <asp:TextBox ID="txtTitle" runat="server" Width="60mm" MaxLength="100"></asp:TextBox>
                            </td>
                        </tr>
                        <tr class="trAddressType">
                            <th align="right">
                                <span class="necessary">※居住地區：</span>
                            </th>
                            <td colspan="3">
                                <asp:Label ID="lblLiveRegion" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr class="trLocal">
                            <th align="right" nowrap="nowrap">
                                <span class="necessary">※聯絡地址：<br />
                                    (寄發收據請務必填寫地址)</span>
                            </th>
                            <td colspan="3">縣市：
                        <asp:DropDownList ID="ddlLiveCity" runat="server">
                        </asp:DropDownList>
                                <asp:Label ID="lblLiveCityPoint" runat="server"></asp:Label>
                                鄉鎮市區：
                        <asp:DropDownList ID="ddlLiveArea" runat="server">
                        </asp:DropDownList>
                                郵遞區號：
                        <asp:TextBox ID="txtLiveZip" runat="server" Width="15mm" MaxLength="5"></asp:TextBox>
                                <!--font color="blue" size="2"><b>範例：23號:請填 23之[空白]號，23號之1:請填 23之1號</b></font-->
                                <br />
                                道路街名或村名
                        <asp:TextBox ID="txtLiveStreet" runat="server" Width="30mm" MaxLength="10"></asp:TextBox>
                                <asp:Label ID="lblLiveStreetPoint" runat="server"></asp:Label>
                                <asp:DropDownList ID="ddlSection" runat="server">
                                </asp:DropDownList>
                                段
                        <asp:TextBox ID="txtLane" runat="server" Width="10mm" MaxLength="10"></asp:TextBox>
                                巷
                        <asp:TextBox ID="txtAlley" runat="server" Width="10mm" MaxLength="10"></asp:TextBox>
                                弄
                        <span id="sHouse">
                            <asp:TextBox ID="txtHouseNo" runat="server" Width="18mm" MaxLength="10"></asp:TextBox>
                            號之
                        <asp:TextBox ID="txtHouseNoSub" runat="server" Width="8mm" MaxLength="5"></asp:TextBox>
                        </span>
                                <asp:Label ID="lblHousePoint" runat="server"></asp:Label>
                                <span id="sFloor">
                                    <asp:TextBox ID="txtFloor" runat="server" Width="10mm" MaxLength="10"></asp:TextBox>
                                    樓之
                        <asp:TextBox ID="txtFloorSub" runat="server" Width="10mm" MaxLength="10"></asp:TextBox>
                                </span>
                                <asp:Label ID="lblFloorPoint" runat="server"></asp:Label>
                                -
                        <asp:TextBox ID="txtRoom" runat="server" Width="10mm" MaxLength="10"></asp:TextBox>
                                室
                        &nbsp;&nbsp;Attn(收件單位/收件人)：<asp:TextBox ID="txtAttn" runat="server" Width="40mm"></asp:TextBox>
                            </td>
                        </tr>
                        <tr class="trOverseas" style="display: none;">
                            <th align="right" nowrap="nowrap">
                                <span class="necessary">※聯絡地址：<br />
                                    (寄發收據請務必填寫地址)</span>
                            </th>
                            <td colspan="3">
                                <%-- <asp:TextBox ID="txtOverseasCountry" runat="server" Width="100mm" MaxLength="50"></asp:TextBox--%>
                                <asp:DropDownList ID="ddlOverseasCountry" runat="server"></asp:DropDownList>
                                國家
                                <asp:Label ID="lblOverseasCountryPoint" runat="server"></asp:Label>
                                <br />
                                <asp:TextBox ID="txtOverseasAddress" runat="server" Width="100mm" MaxLength="50"></asp:TextBox>
                                地址
                                <asp:Label ID="lblOverseasAddressPoint" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <th align="right">電話：
                            </th>
                            <td nowrap="nowrap">
                                <asp:TextBox ID="txtPhoneCode" runat="server" Width="8mm" MaxLength="5"></asp:TextBox>
                                ─
                                <asp:TextBox ID="txtCorpPhone" runat="server" Width="35mm" MaxLength="15"></asp:TextBox>
                                分機：
                                <asp:TextBox ID="txtExt" runat="server" Width="20mm" MaxLength="5"></asp:TextBox>
                            </td>
                            <th align="right">行動電話：
                            </th>
                            <td>
                                <asp:TextBox ID="txtCellPhone" runat="server" Width="30mm" MaxLength="40"></asp:TextBox>
                                <font color="blue"><b>格式：0987123456</b></font>
                            </td>
                        </tr>
                        <tr>
                            <%--<th align="right">
                        住家電話：
                                </th>
                                <td>
                                    <asp:TextBox ID="txtLocalPhone" runat="server" Width="40mm"></asp:TextBox>
                                </td>--%>
                            <th align="right">傳真：
                            </th>
                            <td>
                                <asp:TextBox ID="txtFaxCode" runat="server" Width="8mm" MaxLength="5"></asp:TextBox>
                                ─
                                <asp:TextBox ID="txtFax" runat="server" Width="35mm" MaxLength="15"></asp:TextBox>
                            </td>
                            <th align="right" style="white-space: nowrap;">是否刊登在捐款人名錄：</th>
                            <td style="white-space: nowrap;">
                                <asp:Label ID="lblAnony" runat="server"></asp:Label>
                                <font color="blue"><b>可在本台官網查詢</b></font>
                            </td>
                        </tr>
                        <tr class="trSubscribe">
                            <th align="right">
                                <span class="necessary">※訂閱紙本月刊：</span>
                            </th>
                            <td>
                                <asp:Label ID="lblSubscribe" runat="server"></asp:Label>
                                <span style="color: blue; font-weight: bold;">定期寄送好消息月刊</span>
                            </td>
                            <td align="right">
                                <!--捐款人編號：-->
                            </td>
                            <td>
                                <asp:TextBox ID="txtDonorID" runat="server" Width="30mm" ReadOnly="true" class="readonly" Visible="False"></asp:TextBox>
                            </td>

                        </tr>
                        <tr class="trIsFdc">
                            <th align="right">
                                <span class="necessary">※同意上傳國稅局申報：</span>
                            </th>
                            <td>
                                <asp:Label ID="lblIsFdc" runat="server"></asp:Label>
                                <span style="color: blue; font-weight: bold;">如同意請務必填寫身份證字號</span>
                            </td>
                            <th align="right">
                                <span id="spanIDNo">身份證字號：</span>
                            </th>
                            <td>
                                <asp:TextBox ID="txtIDNo" runat="server" Width="30mm" MaxLength="10"></asp:TextBox>
                                <asp:Label ID="lblIDNoPoint" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <%-- <th align="right" nowrap="nowrap">想對GOOD TV說的話：
                            </th>
                            <td>
                                <asp:TextBox ID="txtToGoodTV" runat="server" Width="90%" TextMode="MultiLine" Rows="5" Font-Size="16px" onkeydown="window.event? window.event.cancelBubble = true : e.stopPropagation();"></asp:TextBox>
                            </td>--%>
                            <th align="right">注意事項：
                            </th>
                            <td colspan="2">本基金會(好消息電視台)向您蒐集上開個人資料將妥善保護，只供本會營運業務相關服務使用，依個資法第3條所定權利，您可向本會請求查詢、閱覽、複製、補充、停止或刪除。<br />捐款服務組聯絡電話(02)8024-3911　　Email：ds@goodtv.tv</td>
                            <td>
                                <span style="font-size: 15px;font-weight:bold">歡迎上GOODTV官網閱讀好消息月刊 <a href="http://www.goodtv.tv/good-news" target="_blank">http://www.goodtv.tv/good-news</a></span>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <br />
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td>
                                <asp:Button ID="btnContinue" runat="server" class="css_btn_class" Text="回上一頁" />
                            </td>
                            <td>&nbsp;
                            </td>
                            <td align="right">
                                <asp:Button ID="btnCheckOut" runat="server" class="css_btn_class" Text="下一步" OnClientClick="return checkInvalid();" OnClick="btnCheckOut_Click" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
