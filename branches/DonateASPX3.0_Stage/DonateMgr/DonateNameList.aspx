<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateNameList.aspx.cs" Inherits="DonateMgr_DonateNameList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>捐款報表(捐款資料)</title>
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
            if (document.getElementById("HFD_Query_Flag").value == '') {
                $("#btnPrint").attr("disabled", "disabled");
                $("#btnFinancePrint").attr("disabled", "disabled");
                $("#btnToxls").attr("disabled", "disabled");
                $("#btnToFinancexls").attr("disabled", "disabled");
            }
        });
        function Print() {
            if (window.confirm('您是否確定要將查詢結果列表') == true) {
                window.open('../DonateMgr/DonateNameList_Print.aspx', 'DonateNameList_Print', 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
            }
            else {
                return;
            }
        }
        function FinancePrint() {
            if (window.confirm('您是否確定要將查詢結果列表') == true) {
                window.open('../DonateMgr/DonateMonthQry_Print.aspx', 'DonateMonthQry_Print', 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
            }
            else {
                return;
            }
        }
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
            //Calendar.setup({
            //    inputField: "tbxAccoun_DateS",   // id of the input field
            //    button: "imgAccoun_DateS"     // 與觸發動作的物件ID相同
            //});
            //Calendar.setup({
            //    inputField: "tbxAccoun_DateE",   // id of the input field
            //    button: "imgAccoun_DateE"     // 與觸發動作的物件ID相同
            //});
        }
        function CheckFieldMustFillBasic() {
            var uiSDate = new Date($("#tbxDonate_DateS").val());
            var uiEDate = new Date($("#tbxDonate_DateE").val());
            var v1 = uiSDate - uiEDate;
            if (uiSDate - uiEDate > 0) {
                alert("捐款日期 開始日至結束日之間範圍有誤！");
                $("#uiSDate").focus();
                return false;
            }
            return true;
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
    <asp:HiddenField runat="server" ID="HFD_Query_Flag"/>
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        捐款報表(捐款資料)</h1>
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
                身份別：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblDonor_Type" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                通訊地址：
            </th>
            <td align="left" colspan="2" >
                <asp:dropdownlist runat="server" ID="ddlCity" CssClass="font9" 
                     AutoPostBack="True" onselectedindexchanged="ddlCity_SelectedIndexChanged"></asp:dropdownlist>
                <asp:dropdownlist runat="server" ID="ddlArea" CssClass="font9"></asp:dropdownlist>
             </td>
            <th align="right" colspan="1">
                收據地址：
            </th>
            <td align="left" colspan="2" >
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
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxDonate_DateS" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgDonate_DateS" alt="" src="../images/date.gif" />
                ~<asp:TextBox runat="server" ID="tbxDonate_DateE" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgDonate_DateE" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                捐款金額(大於)：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxDonate_AmtS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="tbxDonate_AmtE" CssClass="font9" 
                    Width="70px"></asp:TextBox>
             </td>
            
        </tr>
        <tr>
            <th align="right" colspan="1">
               捐款人編號： 
            </th>
            <td align="left" colspan="2" >
                <asp:TextBox runat="server" ID="txtDonor_IdS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="txtDonor_IdE" CssClass="font9" Width="70px"></asp:TextBox>
            </td>
            <!--th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="2" >
                <asp:dropdownlist runat="server" ID="ddlActName" CssClass="font9"></asp:dropdownlist>
            </td-->
            <th align="right" colspan="1">
                收據編號： 
            </th>
            <td align="left" colspan="2" >
                <asp:TextBox runat="server" ID="tbxInvoice_NoS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="tbxInvoice_NoE" CssClass="font9" Width="70px"></asp:TextBox>
             </td>
        </tr>
        <!--tr>
            <th align="right" colspan="1">
                捐款人：
            </th>
            <td align="left" colspan="5">
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
            </td>
        </tr-->
        <tr>
            <th align="right" colspan="1">
                捐款方式：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblDonate_Payment" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                線上付款方式：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblPayment_type" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款用途：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblDonate_Purpose" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款類別：
            </th>
            <td align="left" colspan="2">
                <asp:CheckBoxList ID="cblDonate_Type" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td> 
            <th align="right" colspan="1">
                區域：
            </th>
            <td align="left" colspan="2">
                 <asp:RadioButtonList ID="rblAddress" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Value="1">國內</asp:ListItem>
                    <asp:ListItem Value="2">海外 </asp:ListItem>
                    <asp:ListItem Value="3" Selected="True">全部 </asp:ListItem>
                </asp:RadioButtonList>
            </td> 
            
            <!--th align="right" colspan="1">
                收據狀態：
            </th>
            <td align="left" colspan="2">
                <asp:CheckBoxList ID="cblExp_Pre" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td--> 
            <!--th align="right" colspan="1">
                入帳銀行：
            </th>
            <td align="left" colspan="2">
                <asp:CheckBoxList ID="cblAccoun_Bank" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td--> 
        </tr>
        <!--tr>
            <th align="right" colspan="1">
                沖帳日期：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxAccoun_DateS" runat="server" onchange="CheckDateFormat(this, '沖帳日期');" Width="70px"></asp:TextBox>
                <img id="imgAccoun_DateS" alt="" src="../images/date.gif" />
                <asp:TextBox ID="tbxAccoun_DateE" runat="server" onchange="CheckDateFormat(this, '沖帳日期');" Width="70px"></asp:TextBox>
                <img id="imgAccoun_DateE" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                會計科目：
            </th>
            <td align="left" colspan="2">
                <asp:CheckBoxList ID="cblAccounting_Title" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td>
        </tr-->
        <tr>
            <th align="right" colspan="1">
                其他：
            </th>
            <td align="left" colspan="5">
                <asp:checkbox runat="server" ID="cbxIsErrAddress" Text="不含錯址" CssClass="font9"></asp:checkbox>
                <asp:checkbox runat="server" ID="cbxNoDie" Text="不含歿" CssClass="font9"></asp:checkbox>
                <asp:checkbox runat="server" ID="cbxIsContact" Text="不主動聯絡" CssClass="font9"></asp:checkbox>
            </td> 
        </tr>
        <tr>
            <td align="left" colspan="2" class="style1">
                捐款金額總計：
                <asp:Label ID="lblAmt" runat="server" Text=""></asp:Label>
                元</td>
            <td align="right" colspan="4" class="style1">
                <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="20mm"
                    Text="查詢" onclick="btnQuery_Click" OnClientClick= "return CheckFieldMustFillBasic()"/>
                <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" OnClientClick="Print();"/>
                <asp:Button ID="btnFinancePrint" CssClass="npoButton npoButton_Print" runat="server"  Width="25mm"
                    Text="財務報表" OnClientClick="FinancePrint();"/>
                <asp:Button ID="btnToxls" CssClass="npoButton npoButton_Excel" runat="server"  Width="20mm"
                    Text="匯出"  OnClientClick="return confirm('您是否確定要將查詢結果匯出？')" onclick="btnToxls_Click"/>
                <asp:Button ID="btnToFinancexls" CssClass="npoButton npoButton_Excel" runat="server"  Width="30mm"
                    Text="財務報表匯出"  OnClientClick="return confirm('您是否確定要將查詢結果匯出？')" onclick="btnToFinancexls_Click"/>
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="8" align="center">
                 <asp:Label ID="lblGridList" runat="server" ></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>
