<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PledgeTodate.aspx.cs" Inherits="DonateMgr_PledgeTodate" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>捐款授權到期列印</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            InitMenu();
        });
        function alertCon(condition) {
            if ($("#ddlContent").val() == "") {
                alert("到期信函內容 欄位不可為空白！");
                $("#ddlContent").focus();
                return false;
            }
            else {
                if (condition == "1") {
                    if (window.confirm('※請注意：若資料無地址到期函將不會列印\n\n您是否確定要將查詢結果列印？') == true) {
                        return true;
                    }
                    else {
                        return false;
                    }
                }
                else if (condition == "2") {
                    if ($("#ddlFormat").val() == "") {
                        alert("名條欄位 欄位不可為空白！");
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
           window.open('../DonateMgr/PledgeTodate_MailPrint.aspx', 'PledgeTodate_MailPrint', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=850,height=600', '');
        }
        function Address() {
           window.open('PledgeTodate_Print/PledgeTodate_Print_Style' + $("#ddlFormat").val() + '.aspx', 'PledgeTodate_Print_Style' + $("#ddlFormat").val(), 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
       }
       function Preview() {
           window.open('../DonateMgr/PledgeTodate_Preview.aspx', 'PledgeTodate_Preview', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=850,height=600', '');
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
        捐款授權到期列印 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right"colspan="1" >
                捐款到期年月：
            </th>
            <td align="left" colspan="2">
                <asp:dropdownlist runat="server" ID="ddlYear_Donate_ToDate" CssClass="font9"></asp:dropdownlist>
                年
                <asp:dropdownlist runat="server" ID="ddlMonth_Donate_ToDate" CssClass="font9"></asp:dropdownlist>
                月
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                授權狀態：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlStatus" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                授權類別：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Type" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                授權捐款人：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                授權編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPledge_Id_From" CssClass="font9"></asp:TextBox>~
                <asp:TextBox runat="server" ID="tbxPledge_Id_To" CssClass="font9"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                到期信函內容：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlContent" CssClass="font9"></asp:dropdownlist>
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
                    <asp:ListItem Value="3">授權編號 </asp:ListItem>
                </asp:RadioButtonList>
            </td> 
        </tr>
        <tr>
            <td align="right" colspan="2" class="style1">
                <asp:Button ID="btnMailPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="23mm"
                    Text="到期函" onclick="btnMailPrint_Click"
                    OnClientClick="return alertCon('1');" />
                <asp:Button ID="btnAddress" runat="server"  Width="20mm"
                    Text="名條" CssClass="npoButton npoButton_Print" 
                    OnClientClick="return alertCon('2');"  onclick="btnAddress_Click"/>
                <asp:Button ID="btnPreview" CssClass="npoButton npoButton_Modify" runat="server"  Width="25mm"
                    Text="樣本預覽"  OnClientClick="return alertCon('3');" onclick="btnPreview_Click"/>
            </td>
        </tr>
        <tr>
            <td align="center" width="100%" colspan="2">
                 <asp:Label ID="lblGridList" runat="server" ></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
