<%@ Page Language="C#" EnableEventValidation="false" AutoEventWireup="true" CodeFile="DonateCreditCard.aspx.cs"
    Inherits="Online_DonateCreditCard" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>線上定期定額奉獻</title>
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
            $('#txtAccountNo,#txtAuthorize').not('').bind("keyup", function (e) {
                var val = $(this).val();
                if (isNaN(val)) {
                    val = val.replace(/[^0-9]/g, '');
                    $(this).val(val);
                }
            });

            //$('#txtAccountNo').maxlength({ maxCharacters: 16 });
            //$('#txtAuthorize').maxlength({ maxCharacters: 3 });
        });                     //end of ready()

        function Foolproof() {
            if ($('#txtCardBank').val().length > 20) {
                alert('【發卡銀行】長度不可超過20！');
                $('#txtCardBank').focus();
                return false;
            }
            if ($('#ddlCardType').val() == '') {
                alert('請選擇【信用卡種類】！');
                $('#ddlCardType').focus();
                return false;
            }
            if ($('#txtAccountNo').val().trim().length < 16) {
                alert('【信用卡卡號】長度需為16碼！');
                $('#txtAccountNo').focus();
                return false;
            }
            if ($('#txtAuthorize').val().length < 3) {
                alert('【授權碼】長度需為3碼！');
                $('#txtAuthorize').focus();
                return false;
            }
            if ($('#ddlValidMonth').val() == '') {
                alert('請選擇【有效期限-月】！');
                $('#ddlValidMonth').focus();
                return false;
            }
            if ($('#ddlValidYear').val() == '') {
                alert('請選擇【有效期限-年】！');
                $('#ddlValidYear').focus();
                return false;
            }

            if ($('#txtRelation').val().length > 20) {
                alert('【與發卡人關係】長度不可超過20！');
                $('#txtRelation').focus();
                return false;
            }
            if ($('#txtCardOwner').val().length > 20) {
                alert('【持卡人姓名】長度不可超過20！');
                $('#txtCardOwner').focus();
                return false;
            }
            if ($('#txtOwnerIdNo').val().length > 10) {
                alert('【持卡人身分證字號】長度不可超過10！');
                $('#txtOwnerIdNo').focus();
                return false;
            }

            // 有效期限驗證是否已過期 by Hilty 2013/12/18
            var today = new Date();
            var tMonth = today.getMonth() + 1;
            var tYear = today.getFullYear().toString().substring(2);

            if ($('#ddlValidYear').val() <= tYear && $('#ddlValidMonth').val() < tMonth) 
            {
                alert('請確認卡片有效期限！');
                return false;
            }

            // 身分證字號檢核 by Hilty 2013/12/19
            if ($('#txtOwnerIdNo').val() != '' && !checkID($('#txtOwnerIdNo').val()))
            {
                return false;
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
                            <strong><font color="#003399">填寫定期定額授權資料</font></strong>
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
                <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="Chocolate" />
                <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="Chocolate" />
                <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="Chocolate" />
                <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="Brown" />
                <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="Chocolate" />
                <p />
                            <asp:Label ID="lblTitle" runat="server" Style="color: #8B0000;"> 平安！請填寫您的信用卡資料以進行定期定額扣款授權</asp:Label>
                <br />
                
                <table align="center" border="0" cellpadding="0" cellspacing="1" class="table_v" width="70%">
                    <tr>
                        <th align="right" style="width: 30mm">
                            發卡銀行：
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtCardBank" runat="server" Width="40mm" MaxLength="20"></asp:TextBox>
                        </td>
                        <th align="right" style="width: 30mm">
                            <span class="necessary">信用卡種類： </span>
                        </th>
                        <td align="left">
                            <asp:DropDownList ID="ddlCardType" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <th align="right" style="width: 30mm">
                            <span class="necessary">信用卡卡號： </span>
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtAccountNo" runat="server" Width="40mm" MaxLength="16"></asp:TextBox>
                        </td>
                        <th align="right" style="width: 30mm">
                            <span class="necessary">授權碼： </span>
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtAuthorize" runat="server" Width="40mm" MaxLength="3"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <th align="right" style="width: 30mm">
                            <span class="necessary">有效期限： </span>
                        </th>
                        <td align="left">
                            <asp:DropDownList ID="ddlValidMonth" runat="server">
                            </asp:DropDownList>
                            月/
                            <asp:DropDownList ID="ddlValidYear" runat="server">
                            </asp:DropDownList>
                            年
                        </td>
                        <th align="right" style="width: 30mm">
                            與持卡人關係：
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtRelation" runat="server" Width="40mm" MaxLength="20"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <th align="right" style="width: 30mm">
                            持卡人姓名：
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtCardOwner" runat="server" Width="40mm" MaxLength="20"></asp:TextBox>
                        </td>
                        <th align="right" style="width: 30mm">
                            持卡人身份證字號：
                        </th>
                        <td align="left">
                            <asp:TextBox ID="txtOwnerIdNo" runat="server" Width="40mm" MaxLength="10"></asp:TextBox>
                        </td>
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
                <asp:Button ID="Button1" class="npoButton npoButton_PrevStep" runat="server" Text="上一步"
                    OnClick="btnPrev_Click" height="40px"/>
                <asp:Button ID="Button2" class="npoButton npoButton_Submit" runat="server" Text="完成授權填寫"
                    OnClientClick="return Foolproof();" OnClick="btnConfirm_Click" Width="120px" height="40px" />
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
