<%@ Page Language="C#" EnableEventValidation="false" AutoEventWireup="true" CodeFile="CheckOut.aspx.cs"
    Inherits="Online_CheckOut" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>線上奉獻資料填寫</title>
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
                    $('.trhideclass1').show();
                }
                else {
                    $('.trhideclass1').hide();
                }
            });

            //若勾選【居住地區】台灣，顯示trLocal，反之則顯示trOverseas
            $('input:radio[name=RDO_LiveRegion]').bind("click", function (e) {
                if ($(this).val() == '台灣') {
                    $('.trLocal').show();
                    $('.trOverseas').hide();
                }
                else {
                    $('.trLocal').hide();
                    $('.trOverseas').show();
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
        });                //end of ready()

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
        function loadData() {
            $("#imgCheck").attr("src", '');
            $("#imgCheck").attr("src", '../images/code_img.ashx');
        }
        function checkLoginData() {
            if (!$.trim($('#txtAccount').val()).length) {
                alert('帳號(電子信箱)不可空白！');
                return false;
            }
            if (!$.trim($('#txtPwd').val()).length) {
                alert('密碼不可空白！');
                return false;
            }
        }

        function checkInvalid() {
            if (!$.trim($('#txtDonorName').val()).length) {
                alert('(奉獻者)姓名不可空白！');
                return false;
            }
            if ($('#rdoSex1').prop("checked") == false && $('#rdoSex2').prop("checked") == false) {
                alert('請選擇性別！');
                return false;
            }

            //20130911Modify by GoodTV-Tanya:E-mail為必填
            if (!$.trim($('#txtEmail').val()).length) {
                alert('E-mail(電子信箱)不可空白！');
                return false;
            }

            // 調整Email檢核 by Hilty 2013/12/18
            var patterns = /^[a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$/;

            if ($.trim($('#txtEmail').val()).length) {
                if (!patterns.test($("#txtEmail").val())) {
                    alert('Email格式有誤，請重新輸入！');
                    $("#txtEmail").focus();
                    return false;
                }
            }

            // 身分證字號檢核 by Hilty 2013/12/19
            if ($('#txtIDNo').val() != '' && !checkID($('#txtIDNo').val())) {
                return false;
            }        

            if ($('#ChkAddNew').prop("checked") == true) {
//                if (!$.trim($('#txtEmail').val()).length) {
//                    alert('新增Email帳號必須輸入E-mail！');
//                    return false;
//                }

                emailIsExist($.trim($('#txtEmail').val()));

                if ($('#HFD_EmailIsExist').val() == 'Y') {
                    alert('該Email帳號已存在，請重新輸入新帳號！');
                    return false;
                }

                if (!checkTextLength($('#txtPassword'), '密碼', 50)) {
                    return false;
                }

                if (!$.trim($('#txtPassword').val()).length) {
                    alert('新增帳號必須輸入密碼！');
                    $('#txtPassword').focus();
                    return false;
                }
                else {
                    if ($('#txtPassword').val() != $('#txtConfirmPwd').val()) {
                        alert('密碼與確認密碼必須一致！');
                        $('#txtConfirmPwd').focus();
                        return false;
                    }
                }
            }

            if (!checkTextLength($('#txtEmail'), 'Email帳號', 80)) {
                return false;
            }

            if (!checkTextLength($('#txtTitle'), '收據抬頭', 100)) {
                return false;
            }

            if (!checkTextLength($('#txtIDNo'), '身分證字號', 10)) {
                return false;
            }

            if (!checkTextLength($('#txtLiveZip'), '郵遞區號', 5)) {
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
            if (!checkTextLength($('#txtOverseasCountry'), '國家/省/城市/區', 50)) {
                return false;
            }
            if (!checkTextLength($('#txtOverseasAddress'), '海外地址', 50)) {
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
                alert('【' + fieldName + '】長度不可超過' + len + '！');
                txtTarget.focus();
                return false;
            }
            return true;
        }

        //檢核Email帳號是否已
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
                            <strong><font color="#003399">填寫您的基本資料</font></strong>
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
            <asp:Button ID="btnOpenLogin" class="npoButton npoButton_Single" runat="server" Text="已有帳號，請點此登入"
                Width="180px" OnClientClick="showLoign();return false;" />
        </ul>
        <ul>
            <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="Chocolate" />
            <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="Brown" />
            <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="Chocolate" />
            <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="Chocolate" />
            <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="Chocolate" />             
        </ul>
    </div>
    <div id="reg" style="position: fixed; z-index: 100; top: 0px;
        left: 0px; height: 100%; width: 100%; background: #000; display: none;">
    </div>
    <div id="openLogin" style="background: #ffffff; border-radius: 5px 5px 5px 5px; color: blue;
        font-size: large; display: none; padding-bottom: 2px; width: 500px; height: 250px;
        z-index: 11001; left: 30%; position: fixed; text-align: center; top: 200px;">
        <br />
        <table style="width: 400px;" align="center" cellpadding="2" cellspacing="4">            
            <tr>
                <td colspan="2" align="center">
                    <asp:Label ID="Label1" runat="server" Text="您有帳號了嗎?" Font-Size="18px"></asp:Label>
                </td>
            </tr>
			<tr>
                <td colspan="2" align="center">
                    <asp:Label ID="Label2" runat="server" ForeColor="#FF8000" Text="使用帳號登入可以自動將儲存的基本資料寫入"
                        Font-Size="15px"></asp:Label>
                </td>
            </tr>
            <tr style="color: #666666">
                <td style="height: 25px" align="right">
                    帳號(Email)：
                </td>
                <td align="left">
                    <asp:TextBox ID="txtAccount" runat="server" Width="200px" Font-Size="20px"></asp:TextBox>
                </td>
            </tr>
            <tr style="color: #666666">
                <td style="height: 25px" align="right">
                    密 &nbsp; 碼：
                </td>
                <td align="left">
                    <asp:TextBox ID="txtPwd" runat="server" Width="200px" TextMode="Password" Font-Size="20px"></asp:TextBox><br />
                </td>
            </tr>
            <tr style="color: #666666">
                <td align="right">
                    驗證碼：
                </td>
                <td align="left">
                    <asp:TextBox ID="txtCheckCode" runat="server" Width="60px" Font-Size="20px"></asp:TextBox>
                    <img id="imgCheck" alt="check" src="../images/code_img.ashx" style="display: inline"
                        align="middle" />
                    <asp:Button ID="btnRefresh" runat="server" Text="更新驗證碼" Height="20px" Width="80px"
                        OnClientClick="loadData();return false;" />
                </td>
            </tr>
            
            <tr>
                <td style="width: 380px; height: 19px" align="center" colspan="2">
                    <asp:Button ID="btnRegister" class="npoButton npoButton_PrevStep" runat="server"
                        Text="歡迎註冊" Width="100px" OnClientClick="showAddNew();" OnClick="btnRegister_Click" />
                    <asp:Button ID="Button2" class="npoButton npoButton_NextStep" runat="server" Text="不想登入帳號但繼續奉獻"
                        Width="190px" OnClientClick="hideAddNew();return false;" />
                    <asp:Button ID="btnLogin" class="npoButton npoButton_Submit" runat="server" Text="登入"
                        Width="80px" OnClientClick="return checkLoginData();" OnClick="btnLogin_Click" />
                </td>
            </tr>
        </table>
    </div>
    <table width="860" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td>
            <table width="860" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
                <tr>
                    <th style="width: 40mm" align="right">
                        <span class="necessary">※姓名：</span>
                    </th>
                    <td>
                        <asp:TextBox ID="txtDonorName" runat="server" Width="30mm"></asp:TextBox>
                    </td>
                    <th style="width: 40mm" align="right">
                        <span class="necessary">※性別：</span>
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
                    <td>
                        <asp:TextBox ID="txtEmail" runat="server" Width="30mm"></asp:TextBox>
                        <asp:Label ID="lblAddNew" runat="server"></asp:Label>
                    </td>
                    <th align="right">
                        收據抬頭：
                    </th>
                    <td>
                        <asp:TextBox ID="txtTitle" runat="server" Width="30mm"></asp:TextBox>
                    </td>
                </tr>
                <tr class="trhideclass1">
                    <th align="right">
                        <span class="necessary">※密碼：</span>
                    </th>
                    <td>
                        <asp:TextBox ID="txtPassword" runat="server" Width="30mm" TextMode="Password"></asp:TextBox>
                    </td>
                    <th align="right">
                        <span class="necessary">※確認密碼：</span>
                    </th>
                    <td>
                        <asp:TextBox ID="txtConfirmPwd" runat="server" Width="30mm" TextMode="Password"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        <span class="necessary">※寄發收據：</span>
                    </th>
                    <td>
                        <asp:Label ID="lblReceipt" runat="server"></asp:Label>
                    </td>
                    <th align="right">
                        徵信錄：
                    </th>
                    <td>
                        <asp:Label ID="lblAnony" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        生日：
                    </th>
                    <td style="width: 70mm">
                        <asp:TextBox ID="txtBirthday" Width="25mm" runat="server" onchange="CheckDateFormat(this, '生日'); ">
                        </asp:TextBox>
                        <img id="imgBirthday" alt="" src="../images/date.gif" />
                        格式：yyyy/mm/dd
                    </td>
                    <th align="right">
                        身份證字號：
                    </th>
                    <td>
                        <asp:TextBox ID="txtIDNo" runat="server" Width="30mm"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        居住地區：
                    </th>
                    <td colspan="3">
                        <asp:Label ID="lblLiveRegion" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr class="trLocal">
                    <th align="right">
                        聯絡地址：
                    </th>
                    <td colspan="3" style="line-height: 25px">
                        縣市：
                        <asp:DropDownList ID="ddlLiveCity" runat="server">
                        </asp:DropDownList>
                        鄉鎮市區：
                        <asp:DropDownList ID="ddlLiveArea" runat="server">
                        </asp:DropDownList>
                        郵遞區號：
                        <asp:TextBox ID="txtLiveZip" runat="server" Width="15mm"></asp:TextBox>
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
                        <asp:TextBox ID="txtHouseNo" runat="server" Width="10mm"></asp:TextBox>
                        號之
                        <asp:TextBox ID="txtHouseNoSub" runat="server" Width="10mm"></asp:TextBox>
                        <br />
                        <asp:TextBox ID="txtFloor" runat="server" Width="10mm"></asp:TextBox>
                        樓之
                        <asp:TextBox ID="txtFloorSub" runat="server" Width="10mm"></asp:TextBox>
                        、
                        <asp:TextBox ID="txtRoom" runat="server" Width="10mm"></asp:TextBox>
                        室
                    </td>
                </tr>
                <tr class="trOverseas">
                    <th align="right">
                        聯絡地址：
                    </th>
                    <td colspan="3" style="line-height: 25px">
                        <asp:TextBox ID="txtOverseasCountry" runat="server" Width="100mm"></asp:TextBox>
                        國家/省/城市/區
                        <br />
                        <asp:TextBox ID="txtOverseasAddress" runat="server" Width="100mm"></asp:TextBox>
                        地址
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        電話：
                    </th>
                    <td>
                        <asp:TextBox ID="txtPhoneCode" runat="server" Width="8mm" MaxLength="5"></asp:TextBox>
                        ─
                        <asp:TextBox ID="txtCorpPhone" runat="server" Width="20mm" MaxLength="15"></asp:TextBox>
                        分機：
                        <asp:TextBox ID="txtExt" runat="server" Width="12mm" MaxLength="5"></asp:TextBox>
                    </td>
                    <th align="right">
                        行動電話：
                    </th>
                    <td>
                        <asp:TextBox ID="txtCellPhone" runat="server" Width="30mm" MaxLength="40"></asp:TextBox>
                        格式：0987123456
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
                        傳真：
                    </th>
                    <td>
                        <asp:TextBox ID="txtFaxCode" runat="server" Width="8mm" MaxLength="5"></asp:TextBox>
                        ─
                        <asp:TextBox ID="txtFax" runat="server" Width="20mm" MaxLength="15"></asp:TextBox>
                    </td>
                    <th align="right">
                        捐款人編號：
                    </th>
                    <td>
                        <asp:TextBox ID="txtDonorID" runat="server" Width="30mm" ReadOnly="true" class="readonly"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        <span class="necessary">※訂閱月刊：</span>
                    </th>
                    <td>
                        <asp:Label ID="lblSubscribe" runat="server"></asp:Label>
                    </td>
                    <th align="right">
                        <span class="necessary">※是否上傳給國稅局：</span>
                    </th>
                    <td>
                        <asp:Label ID="lblIsFdc" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <th align="right">
                        想對GOOD TV說的話：
                    </th>
                    <td colspan="3">
                        <asp:TextBox ID="txtToGoodTV" runat="server" Width="120mm" TextMode="MultiLine" Rows="5"></asp:TextBox>
                    </td>
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
                        <td width="133">
                            <asp:Button ID="btnContinue" runat="server" class="npoButton npoButton_Submit" Text="繼續奉獻"
                                OnClick="btnContinue_Click" Width="120px" height="30px" />
                            <%--<asp:ImageButton ID="btnContinue" runat="server" OnClick="btnContinue_Click" ImageUrl="images/btn_繼續奉獻_O.jpg"
                                onmouseover="this.src='images/btn_繼續奉獻_down.jpg'" onmouseout="this.src='images/btn_繼續奉獻_O.jpg'" />--%>
                        </td>
                        <td width="10">
                            &nbsp;
                        </td>
                        <td width="154">
                            <asp:Button ID="btnCheckOut" runat="server" class="npoButton npoButton_Submit" Text="完成填寫"
                                OnClientClick="return checkInvalid();" OnClick="btnCheckOut_Click" Width="120px" height="30px" />
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
        </tr>
    </table>--%>
    </form>
</body>
</html>
