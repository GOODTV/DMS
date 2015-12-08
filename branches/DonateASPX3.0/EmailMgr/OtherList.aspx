<%@ Page Language="C#" AutoEventWireup="true" CodeFile="OtherList.aspx.cs" Inherits="EmailMgr_OtherList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>其他信函列印</title>
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
        function Print(PrintType) {
            window.open('../DonorMgr/DonorQry_Print.aspx', 'DonorQry_Print', 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
        }
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxDonate_DateS",   // id of the input field
                button: "imgtbxDonate_DateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonate_DateE",   // id of the input field
                button: "imgtbxDonate_DateE"     // 與觸發動作的物件ID相同
            });
        }
        function alertCon(condition) {
            if ($("#ddlEmail_subject").val() == "") {
                alert("其他信函內容 欄位不可為空白！");
                $("#ddlEmail_subject").focus();
                return false;
            }
            else {
                if (condition == "1") {
                    if (window.confirm('※請注意：若資料無地址將不會列印\n\n您是否確定要將查詢結果列印？') == true) {
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
            window.open('../EmailMgr/Other_MailPrint.aspx', 'Other_MailPrint', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
        }
        function Address() {
            window.open('OtherList_Print/OtherList_Print_Style' + $("#ddlFormat").val() + '.aspx', 'OtherList_Print_Style' + $("#ddlFormat").val(), 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
        }
        function Preview() {
            window.open('../EmailMgr/Other_Preview.aspx', 'Other_Preview', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
        }
    </script>
    <style type="text/css">
        .style1
        {
            height: 35px;
        }
     </style>
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
        其他信函列印</h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
<%--        <tr>
            <th align="right" colspan="1">
                 類別：
            </th>
            <td align="left" colspan="1">
                <asp:CheckBoxList ID="cblDonor_Category" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td>
        </tr>--%>
        <tr>
            <th align="right" colspan="1">
                身份別：
            </th>
            <td align="left" colspan="7">
                <asp:CheckBoxList ID="cblDonor_Type" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                通訊地址：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlCity" CssClass="font9" 
                     AutoPostBack="True" onselectedindexchanged="ddlCity_SelectedIndexChanged"></asp:dropdownlist>
                <asp:dropdownlist runat="server" ID="ddlArea" CssClass="font9"></asp:dropdownlist>
             </td>
            <th align="right" colspan="1">
                收據地址：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlCity2" CssClass="font9" 
                     AutoPostBack="True" 
                    onselectedindexchanged="ddlCity2_SelectedIndexChanged"></asp:dropdownlist>
                <asp:dropdownlist runat="server" ID="ddlArea2" CssClass="font9"></asp:dropdownlist>
             </td>     
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款日期：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonate_DateS" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgtbxDonate_DateS" alt="" src="../images/date.gif" />
                ~<asp:TextBox runat="server" ID="tbxDonate_DateE" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgtbxDonate_DateE" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                捐款總金額：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonate_TotalS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="tbxDonate_TotalE" CssClass="font9" 
                    Width="70px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlActName" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                捐款總次數： 
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonate_NoS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="tbxDonate_NoE" CssClass="font9" Width="70px"></asp:TextBox>
            </td>
        </tr>
        <tr> 
            <th align="right" colspan="1">
                捐款人：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                捐款人編號： 
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonor_IdS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="tbxDonor_IdE" CssClass="font9" Width="70px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1" >
                生日月份：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlBirthMonth" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                多久未捐款：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlHowLong" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                文宣品：
            </th>
            <td align="left" colspan="3">
                <asp:CheckBoxList ID="cblPropaganda" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                其他：
            </th>
            <td align="left" colspan="1" >
                <asp:checkbox runat="server" ID="cbxIsAbroad" Text="海外地址" CssClass="font9"></asp:checkbox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                其他信函內容：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlEmail_subject" CssClass="font9"></asp:dropdownlist>
            </td>
            
        </tr>
        <tr>
            <th align="right" colspan="1">
                名條內容：
            </th>
            <td align="left" colspan="1">
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
            <td align="left" colspan="1">
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
            <td align="right" colspan="4" class="style1">
                <asp:Button ID="btnMailPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="23mm"
                    Text="勸募函" OnClientClick="return alertCon('1');" onclick="btnMailPrint_Click"/>
                <asp:Button ID="btnAddress" runat="server"  Width="20mm"
                    Text="名條" CssClass="npoButton npoButton_Print" 
                    OnClientClick="return alertCon('2');"  onclick="btnAddress_Click"/>
                <asp:Button ID="btnPreview" CssClass="npoButton npoButton_Modify" runat="server"  Width="25mm"
                    Text="樣本預覽"  OnClientClick="return alertCon('3');" onclick="btnPreview_Click"/>
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="8">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
