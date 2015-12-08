<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonorQry.aspx.cs" Inherits="DonorMgr_DonorQry" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>捐款報表(捐款人)</title>
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
                $("#btnToxls").attr("disabled", "disabled");
            }
        });
        function Print(PrintType) {
            window.open('../DonorMgr/DonorQry_Print.aspx', 'DonorQry_Print', 'width=800,height=600,scrollbars=yes,menubar=yes,location =no,status=no,toolbar=no', '');
        }
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "txtDonateDateS",   // id of the input field
                button: "imgtxtDonateDateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtDonateDateE",   // id of the input field
                button: "imgtxtDonateDateE"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtLastDonateDateS",   // id of the input field
                button: "imgtxtLastDonateDateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtLastDonateDateE",   // id of the input field
                button: "imgtxtLastDonateDateE"     // 與觸發動作的物件ID相同
            });
        }
        function CheckFieldMustFillBasic() {
            var uiSDate = new Date($("#txtDonateDateS").val());
            var uiEDate = new Date($("#txtDonateDateE").val());
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
        捐款報表(捐款人)</h1>
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
            <th align="right" colspan="1">
                身份別：
            </th>
            <td align="left" colspan="7">
                <asp:CheckBoxList ID="cblDonor_Type" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList><div style="color: blue">PS.此處的「讀者」曾有「奉獻記錄」，不是「單純」的讀者，若想搜尋讀者有哪些人，請至「捐款人基本資料維護」或「客製報表」搜尋。</div>
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
            <td align="left" colspan="3" >
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
                <asp:TextBox runat="server" ID="txtDonateDateS" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgtxtDonateDateS" alt="" src="../images/date.gif" />
                ~<asp:TextBox runat="server" ID="txtDonateDateE" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgtxtDonateDateE" alt="" src="../images/date.gif" />
             </td>
            <th align="right" colspan="1">
                    捐款累積金額(大於)： 
                </th>
                <td align="left" colspan="3">
                    <asp:TextBox ID="tbxDonate_Total_Amt" runat="server" Width="90px"></asp:TextBox>
                </td>
        </tr>
        <tr>
            <!--th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlActName" CssClass="font9"></asp:dropdownlist>
            </td-->
            <th align="right" colspan="1" >
                生日月份：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlBirthMonth" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                捐款金額(大於)：
            </th>
            <td align="left" colspan="3" >
                <asp:TextBox runat="server" ID="txtDonate_AmtS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="txtDonate_AmtE" CssClass="font9" 
                    Width="70px"></asp:TextBox>
             </td>
        </tr>
        <tr>
            <!--th align="right" colspan="1">
                捐款人：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="txtDonor_Name" CssClass="font9"></asp:TextBox>
            </td-->
            <th align="right" colspan="1">
                文宣品：
            </th>
            <td align="left" colspan="1">
                <asp:CheckBoxList ID="cblPropaganda" runat="server" 
                    RepeatDirection="Horizontal">
                </asp:CheckBoxList>
            </td>
            <th align="right" colspan="1">
                捐款總累計金額：
            </th>
            <td align="left" colspan="3" >
                <asp:TextBox runat="server" ID="txtDonate_TotalS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="txtDonate_TotalE" CssClass="font9" 
                    Width="70px"></asp:TextBox>
                <font style="color: blue">此條件為搜尋捐款人歷年來的捐款金額總和</font>
             </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                多久未捐款：
            </th>
            <td align="left" colspan="1" >
                <asp:dropdownlist runat="server" ID="ddlHowLong" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                捐款總累計次數： 
            </th>
            <td align="left" colspan="3" >
                <asp:TextBox runat="server" ID="txtDonate_NoS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="txtDonate_NoE" CssClass="font9" Width="70px"></asp:TextBox>
                <font style="color: blue">此條件為搜尋捐款人歷年來的捐款次數總和</font>
             </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
               末捐日期：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="txtLastDonateDateS" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '末捐日期');"></asp:TextBox>
                <img id="imgtxtLastDonateDateS" alt="" src="../images/date.gif" />
                ~<asp:TextBox runat="server" ID="txtLastDonateDateE" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '末捐日期');"></asp:TextBox>
                <img id="imgtxtLastDonateDateE" alt="" src="../images/date.gif" />
             </td>
            <th align="right" colspan="1">
               捐款人編號： 
            </th>
            <td align="left" colspan="3" >
                <asp:TextBox runat="server" ID="txtDonor_IdS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="txtDonor_IdE" CssClass="font9" Width="70px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                區域：
            </th>
            <td align="left" colspan="1">
                 <asp:RadioButtonList ID="rblAddress" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Value="1">國內</asp:ListItem>
                    <asp:ListItem Value="2">海外 </asp:ListItem>
                    <asp:ListItem Value="3" Selected="True">全部 </asp:ListItem>
                </asp:RadioButtonList>
            </td> 
            <th align="right" colspan="1">群組類別：</th>
            <td align="left" colspan="1">
                <asp:DropDownList ID="ddlGroupClass" runat="server" CssClass="font9">
                </asp:DropDownList>
            </td>
            <th align="right" colspan="1">群組代表：</th>
                <td align="left" colspan="1">
                    <asp:TextBox runat="server" ID="txtGroupItemName" Width="60mm" CssClass="font9"></asp:TextBox>
             </td>
        </tr>
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
                    Text="列表" OnClientClick="if (confirm('您是否確定要將查詢結果匯出？')==false) {return false;} Print('');"/>
                <asp:Button ID="btnToxls" CssClass="npoButton npoButton_Excel" runat="server"  Width="20mm"
                    Text="匯出" OnClientClick="return confirm('您是否確定要將查詢結果匯出？')" onclick="btnToxls_Click"/>
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="8" align="center">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
</body>
</html>