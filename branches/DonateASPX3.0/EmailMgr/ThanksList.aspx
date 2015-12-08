<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ThanksList.aspx.cs" Inherits="EmailMgr_ThanksList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>首捐感謝函列印</title>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            InitMenu();
        });
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxDonate_DateS",   // id of the input field
                button: "imgDonate_DateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonate_DateE",   // id of the input field
                button: "imgDonate_DateE"     // 與觸發動作的物件ID相同
            });
        }
        function alertCon(condition) {
            if ($("#ddlEmail_subject").val() == "") {
                alert("感謝函內容 欄位不可為空白！");
                $("#ddlEmail_subject").focus();
                return false;
            }
            else {
                if (condition == "1") {
                    if (window.confirm('※請注意：若資料無地址感謝函將不會列印\n\n您是否確定要將查詢結果列印？') == true) {
                        return true;
                    }
                    else {
                        return false;
                    }
                }
                else if (condition == "2") {
                    if ($("#ddlFormat").val() == "") {
                        alert("名條格式 欄位不可為空白！");
                        $("#ddlFormat").focus();
                        return false;
                    }
                    else {
                        if (window.confirm('您是否確定要列印地址名條？') == true) {
                            return true;
                        }
                        else {
                            return false;
                        }
                    }
                }
                else if (condition == "3") {
                    return true;
                }
            }
        }
        function MailPrint() {
            window.open('../EmailMgr/Thanks_MailPrint.aspx', 'Thanks_MailPrint', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
        }
        function Address() {
            window.open('ThankList_Print/ThankList_Print_Style' + $("#ddlFormat").val() + '.aspx', 'ThankList_Print_Style' + $("#ddlFormat").val(), 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
        }
        function Preview() {
            window.open('../EmailMgr/Thanks_Preview.aspx', 'Thanks_Preview', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
        }
    </script>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        首捐感謝函列印 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款人：
            </th>
            <td align="left" colspan="2" >
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
                &nbsp;&nbsp;&nbsp;
                <asp:CheckBox ID="cbxThanks" text = "感謝函重印" runat="server" />
                <asp:CheckBox ID="cbxAddress" text = "地址名條重印" runat="server" />
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款人編號： 
            </th>
            <td align="left" colspan="2" >
                <asp:TextBox runat="server" ID="tbxDonor_IdS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="tbxDonor_IdE" CssClass="font9" Width="70px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
               捐款日期：
            </th>
            <td align="left" colspan="2" >
                <asp:TextBox runat="server" ID="tbxDonate_DateS" onchange="CheckDateFormat(this, '捐款日期');" CssClass="font9" 
                    Width="70px"></asp:TextBox>
                <img id="imgDonate_DateS" alt="" src="../images/date.gif" />
                ~<asp:TextBox runat="server" ID="tbxDonate_DateE" onchange="CheckDateFormat(this, '捐款日期');" CssClass="font9" 
                    Width="70px"></asp:TextBox>
                <img id="imgDonate_DateE" alt="" src="../images/date.gif" />
             </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlActName" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                感謝函內容：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlEmail_subject" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                名條內容：
            </th>
            <td align="left" colspan="2">
                 <asp:RadioButtonList ID="rblContent" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Value="1" Selected="True">收據地址</asp:ListItem>
                    <asp:ListItem Value="2">通訊地址 </asp:ListItem>
                </asp:RadioButtonList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                名條格式：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlFormat" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                排序方式：
            </th>
            <td align="left" colspan="1">
                 <asp:RadioButtonList ID="rblSort" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Value="1" Selected="True">郵遞區號</asp:ListItem>
                    <asp:ListItem Value="2">捐款人編號</asp:ListItem>
                </asp:RadioButtonList>
            </td> 
        </tr>
        <tr>
            <td align="right" colspan="3" class="style1">
                <asp:Button ID="btnMailPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="23mm"
                    Text="感謝函" OnClientClick="return alertCon('1');" onclick="btnMailPrint_Click"/>
                <asp:Button ID="btnAddress" runat="server"  Width="20mm"
                    Text="名條" CssClass="npoButton npoButton_Print" 
                    OnClientClick="return alertCon('2');"  onclick="btnAddress_Click"/>
                <asp:Button ID="btnPreview" CssClass="npoButton npoButton_Modify" runat="server"  Width="25mm"
                    Text="樣本預覽"  OnClientClick="return alertCon('3');" onclick="btnPreview_Click"/>
            </td>
        </tr>
        <tr>
            <td align="center" width="100%" colspan="3">
                 <asp:Label ID="lblGridList" runat="server" ></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>

