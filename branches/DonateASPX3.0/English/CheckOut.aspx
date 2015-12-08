<%@ Page Language="C#" EnableEventValidation="false" AutoEventWireup="true" CodeFile="CheckOut.aspx.cs"
    Inherits="Online_CheckOut" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Donor Information</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/jquery.leanModal.js"></script>
    <script type="text/javascript" src="../include/jquery.address.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            //初始化日曆
            initCalendar(['Birthday']);
            //載入地址下拉選項
            $().CityAreaZip('ddlLiveCity', 'ddlLiveArea', 'txtLiveZip', 'HFD_LiveArea');

            //若勾選【新增帳號】，顯示輸入密碼，反之則隱藏
            $('#ChkAddNew').bind("click", function (e) {
                if ($(this).prop('checked') == true) {

                    if ($.trim($('#txtEmail').val()).length) {

                        if ($.trim($('#txtEmail').val()).length) {
                            var patterns = /^[a-zA-Z0-9_-].+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$/;
                            if (!patterns.test($("#txtEmail").val())) {
                                alert('The format of email address is wrong. Please fill in your correct email address.');
                                $("#txtEmail").val("");
                                $("#txtEmail").focus();
                                return false;
                            }
                        }
                        emailIsExist($.trim($('#txtEmail').val()));
                        if ($('#HFD_EmailIsExist').val() == 'Y') {
                            alert('The email address has already been used. Please create a new account(email address).');
                            $("#txtEmail").val("");
                            $("#txtEmail").focus();
                            return false;
                        }
                    }
                    $('.trhideclass1').show();
                }
                else {
                    $('.trhideclass1').hide();
                }
            });

            $('#txtEmail').blur(function () {

                if ($('#HFD_showAddNewCheckbox').val() != "N") {

                    if ($.trim($('#txtEmail').val()).length) {

                        if ($.trim($('#txtEmail').val()).length) {
                            var patterns = /^[a-zA-Z0-9_-].+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$/;
                            if (!patterns.test($("#txtEmail").val())) {
                                alert('The format of email address is wrong. Please fill in your correct email address.');
                                $("#txtEmail").val("");
                                $("#txtEmail").focus();
                                return false;
                            }
                        }
                        emailIsExist($.trim($('#txtEmail').val()));
                        if ($('#HFD_EmailIsExist').val() == 'Y') {
                            if ($('#ChkAddNew').prop('checked') == true) {
                                alert('The email address has already been used. Please create a new account(email address).');
                                $("#txtEmail").val("");
                                $("#txtEmail").focus();
                            }
                            else {
                                alert('The email address has already been used. Please sign in first.');
                                $("#txtEmail").val("");
                                $("#txtEmail").focus();
                            }
                            return false;
                        }
                    }
                }
            });

            //若勾選【居住地區】台灣，顯示trLocal，反之則顯示trOverseas
            $('input:radio[name=RDO_LiveRegion]').bind("click", function (e) {
                if ($(this).val() == '台灣') {
                    $('.trLocal').show();
                    $('.trOverseas').hide();
                    $('.trSubscribe').show();
                }
                else {
                    $('.trLocal').hide();
                    $('.trOverseas').show();
                    $('.trSubscribe').hide();
                    $('#RdoSubscribe2').prop('checked') = true;
                }
            });

            //輸入密碼預設隱藏
            $('.trhideclass1').hide();

            //居住地區(海外)預設隱藏
            if ($('#HFD_LiveRegion').val() == "台灣") {
                $('.trLocal').show();
                $('.trOverseas').hide();
            }
            else {
                $('.trLocal').hide();
                $('.trOverseas').show();
                $('.trSubscribe').hide();
            }

            if ($('#HFD_ShowLogin').val() == "N") {
                hideLoign();
            }
            else {
                showLoign();
            }
            if ($('#HFD_showAddNewCheckbox').val() == "N") {
                hideAddNewCheckbox();
            }
            else {
                showAddNewCheckbox();
            }
            if ($('#HFD_AddNewChecked').val() == "Y") {
                showAddNew();
            }

            //限制輸入數字
            $('#txtPhoneCode,#txtCorpPhone,#txtExt,#txtCellPhone,#txtFaxCode,#txtFax').not('').bind("keyup", function (e) {
                var val = $(this).val();
                if (isNaN(val)) {
                    val = val.replace(/[^0-9]/g, '');
                    $(this).val(val);
                }
            });
        });                   //end of ready()
        
        function showAddNew() {
            $("#reg").hide()
            $('#openLogin').hide();
            $('.trhideclass1').show();
            $('#ChkAddNew').prop('checked', true);
        }

        function showAddNewCheckbox() {
            $('#lblAddNew').show();
        }

        function hideAddNewCheckbox() {
            $('#lblAddNew').hide();
        }

        function hideAddNew() {
            $("#reg").hide()
            $('#openLogin').hide();
            $('.trhideclass1').hide();
            $('#ChkAddNew').prop('checked', false);
        }

        function showLoign() {
            $("#reg").fadeTo("fast", 0.6);
            $('#openLogin').show();
        }

        function hideLoign() {
            $("#reg").hide()
            $('#openLogin').hide();
        }
        //Refresh驗證碼
        /*function loadData() {
            $("#imgCheck").attr("src", '');
            $("#imgCheck").attr("src", '../images/code_img.ashx');
        }*/
        function checkLoginData() {
            if (!$.trim($('#txtAccount').val()).length) {
                alert('Please fill in your email address.');
                return false;
            }
            if (!$.trim($('#txtPwd').val()).length) {
                alert('Please fill in your password.');
                return false;
            }
        }

        function checkInvalid() {
            if (!$.trim($('#txtDonorName').val()).length) {
                alert('Please fill in donor name.');
                return false;
            }
            if ($('#rdoSex1').prop("checked") == false && $('#rdoSex2').prop("checked") == false) {
                alert('Please choose your gender.');
                return false;
            }
            //20140611新增
            var strRet = "";
            var ddlLiveCity = document.getElementById('ddlLiveCity');
            var txtLiveStreet = document.getElementById('txtLiveStreet');
            if ($('#RdoLiveRegion1').prop('checked') == true) {
                if ($('#rdoReceipt1').prop("checked") == true || $('#rdoReceipt2').prop("checked") == true) {
                    if (ddlLiveCity.value == "") {
                        strRet += "縣市 ";
                    }
                    if (txtLiveStreet.value == "") {
                        strRet += "道路街名或村名 ";
                    }
                    if (strRet != "") {
                        strRet += "欄位不可為空白！";
                        alert(strRet)
                        return false;
                    }
                }


                // 2014/7/8 增加數字欄位檢查 by Samson
                if (isNaN($('#txtLane').val())) {
                    alert('巷 欄位必須為數字！');
                    $('#txtLane').focus();
                    return false;
                }
                if (isNaN($('#txtAlley').val())) {
                    alert('弄 欄位必須為數字！');
                    $('#txtAlley').focus();
                    return false;
                }
                if (isNaN($('#txtHouseNo').val())) {
                    alert('號 欄位必須為數字！');
                    $('#txtHouseNo').focus();
                    return false;
                }
                if (isNaN($('#txtHouseNoSub').val())) {
                    alert('之號 欄位必須為數字！');
                    $('#txtHouseNoSub').focus();
                    return false;
                }
                if (isNaN($('#txtFloor').val())) {
                    alert('樓 欄位必須為數字！');
                    $('#txtFloor').focus();
                    return false;
                }
                if (isNaN($('#txtFloorSub').val())) {
                    alert('之樓 欄位必須為數字！');
                    $('#txtFloorSub').focus();
                    return false;
                }

                // 2014/7/8 增加檢查幾之幾號和幾之幾樓 by Samson
                if ($('#txtHouseNo').val() == "" && $('#txtHouseNoSub').val() != "") {
                    alert('門牌格式錯誤！正確格式範例：6之7號或6之(空白)號');
                    $('#txtHouseNo').focus();
                    return false;
                }
                if ($('#txtFloor').val() == "" && $('#txtFloorSub').val() != "") {
                    alert('門牌格式錯誤！正確格式範例：7樓之1或7樓之(空白)！');
                    $('#txtFloor').focus();
                    return false;
                }


                if (!checkTextLength($('#txtLiveStreet'), '道路街名或村名', 10)) {
                    return false;
                }
                if (!checkTextLength($('#txtLane'), '巷', 10)) {
                    return false;
                }
                if (!checkTextLength($('#txtAlley'), '弄', 10)) {
                    return false;
                }
                if (!checkTextLength($('#txtHouseNo'), '號', 10)) {
                    return false;
                }
                if (!checkTextLength($('#txtHouseNoSub'), '號之', 10)) {
                    return false;
                }
                if (!checkTextLength($('#txtFloor'), '樓', 10)) {
                    return false;
                }
                if (!checkTextLength($('#txtFloorSub'), '樓之', 10)) {
                    return false;
                }
                if (!checkTextLength($('#txtRoom'), '室', 10)) {
                    return false;
                }


            }
            else {

                if ($('#txtOverseasCountry').val() == "") {
                    alert('Country is required.');
                    $('#txtOverseasCountry').focus();
                    return false;
                }
                if ($('#txtOverseasAddress').val() == "") {
                    alert('City, State, Street address is required.');
                    $('#txtOverseasAddress').focus();
                    return false;
                }

                if (!checkTextLength($('#txtOverseasCountry'), 'Country', 50)) {
                    return false;
                }
                if (!checkTextLength($('#txtOverseasAddress'), 'City, State, Street address', 50)) {
                    return false;
                }

            }

            //20130911Modify by GoodTV-Tanya:E-mail為必填
            if (!$.trim($('#txtEmail').val()).length) {
                alert('Email address is required.');
                return false;
            }
            
            // 調整Email檢核 by Hilty 2013/12/18
            if ($.trim($('#txtEmail').val()).length) {
                var patterns = /^[a-zA-Z0-9_-].+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$/;
                if (!patterns.test($("#txtEmail").val())) {
                    alert('The format of email address is wrong. Please fill in your correct email address.');
                    $("#txtEmail").val("");
                    $("#txtEmail").focus();
                    return false;
                }
            }
            
            // 身分證字號檢核 by Hilty 2013/12/19
            //if ($('#txtIDNo').val() != '' && !checkID($('#txtIDNo').val())) {
            //    return false;
            //}        

			
            if ($('#ChkAddNew').prop("checked") == true) {
//                if (!$.trim($('#txtEmail').val()).length) {
//                    alert('新增Email帳號必須輸入E-mail！');
//                    return false;
//                }

                emailIsExist($.trim($('#txtEmail').val()));

                if ($('#HFD_EmailIsExist').val() == 'Y') {
                    alert('The email address has already been used. Please fill in other email address not been used.');
                    $("#txtEmail").val("");
                    $("#txtEmail").focus();
                    return false;
                }

                if (!checkTextLength($('#txtPassword'), 'Password', 50)) {
                    return false;
                }

                if (!$.trim($('#txtPassword').val()).length) {
                    alert('Please key in your password before you create a new account.');
                    $('#txtPassword').focus();
                    return false;
                }

                else {
                    if ($('#txtPassword').val() != $('#txtConfirmPwd').val()) {
                        alert('"Password" must be the same as "Confirm password".');
                        $('#txtConfirmPwd').focus();
                        return false;
                    }
                }
            }

            if (!checkTextLength($('#txtDonorName'), 'Name', 100)) {
                return false;
            }

            if (!checkTextLength($('#txtEmail'), 'E-mail', 80)) {
                return false;
            }

            if (!checkTextLength($('#txtTitle'), 'Name on receipt', 100)) {
                return false;
            } 

            //if (!checkTextLength($('#txtIDNo'), '身分證字號', 10)) {
            //    return false;
            //}

            if (!checkTextLength($('#txtLiveZip'), '郵遞區號', 5)) {
                return false;
            }
            if (!checkTextLength($('#txtPhoneCode'), '電話區碼', 5)) {
                return false;
            }
            if (!checkTextLength($('#txtCorpPhone'), '電話號碼', 15)) {
                return false;
            }
            if (!checkTextLength($('#txtExt'), '電話分機', 5)) {
                return false;
            }
            if (!checkTextLength($('#txtCellPhone'), '行動電話', 40)) {
                return false;
            }
            if (!checkTextLength($('#txtFaxCode'), '傳真區碼', 5)) {
                return false;
            }
            if (!checkTextLength($('#txtFax'), '傳真號碼', 15)) {
                return false;
            }

            return true;
        }

        //限制資料輸入長度
        function checkTextLength(txtTarget, fieldName, len) {
            if (txtTarget.val().length > len) {
                alert('The length of 【' + fieldName + '】 can not exceed ' + len + ' characters！');
                txtTarget.focus();
                return false;
            }
            return true;
        }

        //檢核Email帳號是否已存在
        function emailIsExist(Email) {
            $.ajax({
                type: 'post',
                url: "../common/ajax.aspx",
                data: 'Type=4' + '&Email=' + Email,
                async: false, //同步
                success: function (result) {
                    $('#HFD_EmailIsExist').val(result);
                },
                error: function () { alert('ajax failed'); }
            })
        }  
    </script>
</head>
<body style="background-color: #EEE">
<asp:Panel ID="Panel1" runat="server">
    <form id="Form1" runat="server">
    <asp:HiddenField ID="HFD_LiveArea" runat="server" />
    <asp:HiddenField ID="HFD_AddNew" runat="server" />
    <asp:HiddenField ID="HFD_DonorID" runat="server" />
    <asp:HiddenField ID="HFD_ShowLogin" runat="server" />
    <asp:HiddenField ID="HFD_Gender" runat="server" />
    <asp:HiddenField ID="HFD_Receipt" runat="server" />
    <asp:HiddenField ID="HFD_Anony" runat="server" />
    <asp:HiddenField ID="HFD_Subscribe" runat="server" />
    <asp:HiddenField ID="HFD_IsFdc" runat="server" />
    <asp:HiddenField ID="HFD_LiveRegion" runat="server" />
    <asp:HiddenField ID="HFD_AddNewChecked" runat="server" />
    <asp:HiddenField ID="HFD_showAddNewCheckbox" runat="server" />
    <asp:HiddenField ID="HFD_EmailIsExist" runat="server" />
    <table width="860" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td width="436">
                <img alt="aa" src="images/bar_logo.jpg" width="390" height="40" />
            </td>
            <td width="424" style="background-image: url('images/bar_item.jpg')">
                <table width="95%" border="0" align="center" cellpadding="5" cellspacing="5">
                    <tr>
                        <td align="left" valign="bottom">
                            <strong><font color="#003399">Donor Information</font></strong>
                        </td>
                        <td align="right" valign="bottom">
                            <strong><font color="#003399">Language：</font><a href="../Online/CheckOut.aspx" target="_self"><font size="3">中文</font></a></strong>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>        
    </table>
    <span>
        <asp:Label ID="lblCart" runat="server"></asp:Label>
    </span>
    <div style="text-align: center">                                     
        <ul>
            <asp:Button ID="btnOpenLogin" class="Online Online_Button10" runat="server" Text="Sign in here if you have an account."
                Width="258px" OnClientClick="showLoign();return false;" />
            <asp:Label ID="Label3" runat="server" Text=" ※required field " ForeColor="Red" Font-Size="Medium" Font-Bold="true" />
        </ul>
        <ul>
            <asp:Label ID="lblStep1" runat="server" Text=" Your Donation Items >> " ForeColor="MediumBlue" />
            <asp:Label ID="lblStep2" runat="server" Text=" Donor Information >> " ForeColor="Red" Font-Size="Medium" Font-Bold="true" />
            <asp:Label ID="lblStep3" runat="server" Text=" Confirm Donation >> " ForeColor="MediumBlue" />
            <asp:Label ID="lblStep4" runat="server" Text=" Pledge Information >> " ForeColor="MediumBlue" />
            <asp:Label ID="lblStep5" runat="server" Text=" Confirm Payment " ForeColor="MediumBlue" />             
        </ul>
    </div>
    <div id="reg" style="position: fixed; z-index: 100; top: 0px;
        left: 0px; height: 100%; width: 100%; background: #000; display: none;">
    </div>
    <asp:Panel ID="ShowLogin" runat="server" DefaultButton="btnLogin">
    <div id="openLogin" style="background: #ffffff; border-radius: 5px 5px 5px 5px; color: blue;
        font-size: large; display: none; padding-bottom: 2px; width: 500px; height: 300px;
        z-index: 11001; left: 30%; position: fixed; text-align: center; top: 200px;">
        <br />
        <table style="width: 500px;" align="center" cellpadding="2" cellspacing="4">            
            <tr>
                <td colspan="2" align="center">
                    <asp:Label ID="Label1" runat="server" Text="Sign in to your account" Font-Size="18px"></asp:Label></br>
                    <asp:Label ID="Label2" runat="server" ForeColor="MediumBlue" Text="An account can bring in your saved information"
                        Font-Size="18px"></asp:Label>
                </td>
            </tr>
			<tr>
                <td colspan="2" align="center">
                    
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <asp:Label ID="Label4" runat="server" ForeColor="Red" Text="The AG number is no longer available, please create a new account"
                        Font-Size="15px"></asp:Label>
                </td>
            </tr>
            <tr style="color: #666666">
                <td style="height: 25px" align="right">
                    Email：
                </td>
                <td align="left">
                    <asp:TextBox ID="txtAccount" runat="server" Width="200px" Font-Size="20px"></asp:TextBox>
                </td>
            </tr>
            <tr style="color: #666666">
                <td style="height: 25px" align="right">
                    Password：
                </td>
                <td align="left">
                    <asp:TextBox ID="txtPwd" runat="server" Width="200px" TextMode="Password" Font-Size="20px"></asp:TextBox>
                    <asp:LinkButton class="npoButton npoButton_pw" runat="server" Text="forgot password" Width="120px" OnClick="btnForgotPassword" /><br />
                </td>
            </tr>
            <%--<tr style="color: #666666">
                <td align="right">
                    Verification code：
                </td>
                <td align="left">
                    <asp:TextBox ID="txtCheckCode" runat="server" Width="60px" Font-Size="20px"></asp:TextBox>
                    <img id="imgCheck" alt="check" src="../images/code_img.ashx" style="display: inline"
                        align="middle" />
                    <asp:Button ID="btnRefresh" runat="server" Text="reset code" Height="20px" Width="100px"
                        OnClientClick="loadData();return false;" />
                </td>
            </tr>--%>
            
            <tr>
                <td style="width: auto; height: 19px" align="center" colspan="2">
                    <%-- <asp:Button ID="Button2" class="Online Online_Button5" runat="server" Text="I don't want to sign in"
                        Width="174px" OnClientClick="hideAddNew();return false;" />--%>
                    <asp:Button ID="btnLogin" class="Online Online_Button3" runat="server" Text="Sign in"
                        Width="90px" OnClientClick="return checkLoginData();" OnClick="btnLogin_Click" />
                    <asp:Button ID="btnRegister" class="Online Online_Button5" runat="server"
                        Text="Create an account" Width="174px" OnClientClick="showAddNew();" OnClick="btnRegister_Click" />
                </td>
            </tr>
        </table>
    </div>
    </asp:Panel>
    <div id="data" runat="server" onkeydown="if(event.keyCode==13) return false;">
    <table width="890" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td>
            <table width="890" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v2">
                <tr>
                    <th style="width: 40mm" align="right">
                        <span class="necessary">※Name：</span>
                    </th>
                    <td>
                        <asp:TextBox ID="txtDonorName" runat="server" Width="50mm"></asp:TextBox>
                    </td>
                    <th align="right">
                        <span class="necessary">※Gender：</span>
                    </th>
                    <td>
                        <asp:Label ID="lblGender" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        <%--20130911Modify by GoodTV-Tanya:E-mail為必填--%>
                        <span id="spEmail" class="necessary">※E-mail：</span>
                    </th>
                    <td width="35%">
                        <asp:TextBox ID="txtEmail" runat="server" Width="50mm"></asp:TextBox>
                        <asp:Label ID="lblAddNew" runat="server"></asp:Label>
                    </td>
                    <th align="right">
                        Name on receipt：
                    </th>
                    <td>
                        <asp:TextBox ID="txtTitle" runat="server" Width="50mm"></asp:TextBox>
                    </td>
                </tr>
                <tr class="trhideclass1">
                    <th align="right">
                        <span class="necessary">※Password：</span>
                    </th>
                    <td>
                        <asp:TextBox ID="txtPassword" runat="server" Width="30mm" TextMode="Password"></asp:TextBox>
                    </td>
                    <th align="right">
                        <span class="necessary">※Confirm password：</span>
                    </th>
                    <td>
                        <asp:TextBox ID="txtConfirmPwd" runat="server" Width="30mm" TextMode="Password"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        <span class="necessary">※Receipt：</span>
                    </th>
                    <td colspan="3">
                        <asp:Label ID="lblReceipt" runat="server"></asp:Label>
                    </td>
                    <!--% <th align="right">
                        是否刊登在捐款人名錄：
                    </th>
                    <td>
                        <asp:Label ID="lblAnony" runat="server"></asp:Label> <font color="blue"><b>可在本台官網查詢</b></font>
                    </td>%-->
                </tr>
                <tr>
                    <th align="right">
                        Birthday：
                    </th>
                    <td style="width: 70mm">
                        <asp:TextBox ID="txtBirthday" Width="25mm" runat="server" onchange="CheckDateFormat(this, '生日'); ">
                        </asp:TextBox>
                        <img id="imgBirthday" alt="" src="../images/date.gif" />
                        <span style="color: blue; font-weight: bold;">yyyy/mm/dd</span>
                    </td>
                    <td align="right">
                        <!--身份證字號：-->
                    </td>
                    <td>
                        <!--<asp:TextBox ID="txtIDNo" runat="server" Width="30mm"></asp:TextBox>-->
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        Place of Residence：
                    </th>
                    <td colspan="3">
                        <asp:Label ID="lblLiveRegion" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr class="trLocal">
                    <th align="right">
                        Contact Address：<br>
                    </th>
                    <td colspan="3" style="line-height: 25px">
                        縣市：
                        <asp:DropDownList ID="ddlLiveCity" runat="server">
                        </asp:DropDownList>
                        鄉鎮市區：
                        <asp:DropDownList ID="ddlLiveArea" runat="server">
                        </asp:DropDownList>
                        郵遞區號：
                        <asp:TextBox ID="txtLiveZip" runat="server" Width="15mm"></asp:TextBox> <font color="blue"><b>範例：23號:請填 23之[空白]號，23號之1:請填 23之1號</b></font>
                        <br />
                        道路街名或村名
                        <asp:TextBox ID="txtLiveStreet" runat="server" Width="30mm"></asp:TextBox>
                        <asp:DropDownList ID="ddlSection" runat="server">
                        </asp:DropDownList>
                        段
                        <asp:TextBox ID="txtLane" runat="server" Width="10mm"></asp:TextBox>
                        巷
                        <asp:TextBox ID="txtAlley" runat="server" Width="10mm"></asp:TextBox>
                        弄
                        <asp:TextBox ID="txtHouseNo" runat="server" Width="10mm" BackColor="#CCEEFF"></asp:TextBox>
                        之
                        <asp:TextBox ID="txtHouseNoSub" runat="server" Width="10mm" BackColor="#CCEEFF"></asp:TextBox>
                        號
                        <asp:TextBox ID="txtFloor" runat="server" Width="10mm" BackColor="#D1BBFF"></asp:TextBox>
                        樓之
                        <asp:TextBox ID="txtFloorSub" runat="server" Width="10mm" BackColor="#D1BBFF"></asp:TextBox>
                        -
                        <asp:TextBox ID="txtRoom" runat="server" Width="10mm"></asp:TextBox>
                        室
                        <br />
                        Attn(收件單位/收件人)：<asp:TextBox ID="txtAttn" runat="server" Width="40mm"></asp:TextBox>
                    </td>
                </tr>
                <tr class="trOverseas">
                    <th align="right">
                        Contact Address：<br>
                    </th>
                    <td colspan="3" style="line-height: 25px">
                        Country <asp:TextBox ID="txtOverseasCountry" runat="server" Width="100mm"></asp:TextBox>
                        <br />
                        City, State, Street address <asp:TextBox ID="txtOverseasAddress" runat="server" Width="100mm"></asp:TextBox>
                        
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        Office：
                    </th>
                    <td>
                        <asp:TextBox ID="txtPhoneCode" runat="server" Width="8mm" MaxLength="5"></asp:TextBox>
                        ─
                        <asp:TextBox ID="txtCorpPhone" runat="server" Width="20mm" MaxLength="15"></asp:TextBox>
                        ext.：
                        <asp:TextBox ID="txtExt" runat="server" Width="12mm" MaxLength="5"></asp:TextBox>
                    </td>
                    <th align="right">
                        Mobile：
                    </th>
                    <td>
                        <asp:TextBox ID="txtCellPhone" runat="server" Width="30mm" MaxLength="40"></asp:TextBox>
                        <font color="blue"><b>0987123456</b></font>
                    </td>
                </tr>
                <tr>
                    <%--            <th align="right">
                        住家電話：
                    </th>
                    <td>
                        <asp:TextBox ID="txtLocalPhone" runat="server" Width="40mm"></asp:TextBox>
                    </td>--%>
                    <th align="right">
                        Fax：
                    </th>
                    <td>
                        <asp:TextBox ID="txtFaxCode" runat="server" Width="8mm" MaxLength="5"></asp:TextBox>
                        ─
                        <asp:TextBox ID="txtFax" runat="server" Width="20mm" MaxLength="15"></asp:TextBox>
                    </td>
                    <td align="right">
                        <!--Donor ID：-->
                    </td>
                    <td>
                        <asp:TextBox ID="txtDonorID" runat="server" Width="30mm" ReadOnly="true" class="readonly" Visible="False"></asp:TextBox>
                    </td>
                </tr>
                <tr class="trSubscribe">
                    <th align="right">
                        <span class="necessary">※訂閱紙本月刊：</span>
                    </th>
                    <td>
                        <asp:Label ID="lblSubscribe" runat="server"></asp:Label> <span style="color: blue; font-weight: bold;">定期寄送好消息月刊</span>
                    </td>
                    <td colspan="2">
                        歡迎上GOODTV官網閱讀 <a href="http://www.goodtv.tv/good-news" target="_blank">http://www.goodtv.tv/good-news</a>
                        <!--span class="necessary">※是否上傳給國稅局：</span-->
                    </td>
                    <!--td>
                        <<asp:Label ID="lblIsFdc" runat="server"></asp:Label>>
                    </td-->
                </tr>
                <tr>
                    <th align="right">
                        Words to GOOD TV：
                    </th>
                    <td colspan="3">
                        <asp:TextBox ID="txtToGoodTV" runat="server" Width="120mm" TextMode="MultiLine" Rows="5"></asp:TextBox>
                        <script type="text/javascript" src="https://seal.godaddy.com/getSeal?sealID=DHpQNoGOKhhCpPbW2AxeTcxVj0p7gI3Yq0VXnSszQ1S6uUuKXOEwGDgvA"></script>
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        Privacy Policy：
                    </th>
                    <td colspan="3" width="82%">At GOODTV, we respect your privacy and are commited to protect the personal information provided on this website. Information you provided will be used for purposes that are merely relevant to our services. If you would like to change your information, please contact us at TEL:(02)8024-3911、Email:ds@goodtv.tv</td>
                </tr>        
            </table>
        </td>
    </tr>
    <tr>
        <td colspan="2">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                        <td width="100">
                            <asp:Button ID="btnContinue" runat="server" class="Online Online_Button3" Text="Previous"
                                OnClick="btnPrev_Click" Width="90px" />
                            <%--<asp:ImageButton ID="btnContinue" runat="server" OnClick="btnContinue_Click" ImageUrl="images/btn_繼續奉獻_O.jpg"
                                onmouseover="this.src='images/btn_繼續奉獻_down.jpg'" onmouseout="this.src='images/btn_繼續奉獻_O.jpg'" />--%>
                        </td>
                        <td width="10">
                            &nbsp;
                        </td>
                        <td width="180">
                            <asp:Button ID="btnCheckOut" runat="server" class="Online Online_Button5" Text="Continue to payment"
                                OnClientClick="return checkInvalid();" OnClick="btnCheckOut_Click" Width="174px" />
                            <%--<asp:ImageButton ID="btnCheckOut" runat="server" OnClientClick="return checkInvalid();" OnClick="btnCheckOut_Click" ImageUrl="images/btn_完成填寫_O.jpg"
                                onmouseover="this.src='images/btn_完成填寫_down.jpg'" onmouseout="this.src='images/btn_完成填寫_O.jpg'" />--%>
                        </td>
                    </tr>
                </table>
            </td>   
    </tr>
    <%--<table width="860" >
        <tr>
            <td width="860" align="right">
                <asp:ImageButton ID="btnContinue" runat="server" OnClick="btnContinue_Click" ImageUrl="images/btn_繼續奉獻_O.jpg"
                    onmouseover="this.src='images/btn_繼續奉獻_down.jpg'" onmouseout="this.src='images/btn_繼續奉獻_O.jpg'" />
                    &nbsp;&nbsp;
                <asp:ImageButton ID="btnCheckOut" runat="server" OnClientClick="return checkInvalid();" OnClick="btnCheckOut_Click" ImageUrl="images/btn_完成填寫_O.jpg"
                    onmouseover="this.src='images/btn_完成填寫_down.jpg'" onmouseout="this.src='images/btn_完成填寫_O.jpg'" />
            </td>
        </tr>--%>
    </table>
    </div>
    </form>
    </asp:Panel>
</body>
</html>
