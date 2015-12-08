<%@ Page Language="C#" AutoEventWireup="true" CodeFile="AddressList.aspx.cs" Inherits="DonorMgr_AddressList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>地址名條列印</title>
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
                inputField: "txtLast_DonateDateS",   // id of the input field
                button: "imgtxtLast_DonateDateS"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "txtLast_DonateDateE",   // id of the input field
                button: "imgtxtLast_DonateDateE"     // 與觸發動作的物件ID相同
            });
        }
        function Address_OnClick(i) {
            var ddlFormat = document.getElementById('ddlFormat');
            if (ddlFormat.value == "") {
                alert('標籤格式 欄位不可為空')
                return false;
            }
            if (window.confirm('您是否確定要將查詢結果列印？') == true) {
                window.open('AddressList_Print/AddressList_Print_Style' + i + '.aspx', 'AddressList_Print_Style' + i, 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600', '');
            }
        }
        function Post_OnClick() {
            if (window.confirm('您是否確定要將查詢結果列印？') == true) {
                window.open('../DonorMgr/AddressList_Post.aspx', 'AddressList_Post', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
            }
        }
        function Help_OnClick() {
            window.open('../DonorMgr/AddressList_Help.htm', 'AddressList_Help', 'toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=300');

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
        地址名條列印</h1>
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
                <asp:dropdownlist runat="server" ID="ddlArea2" CssClass="font9" ></asp:dropdownlist>
             </td>
                         
        </tr>
        <tr>
            
            <th align="right" colspan="1">
               捐款日期：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="txtLast_DonateDateS" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgtxtLast_DonateDateS" alt="" src="../images/date.gif" />
                ~<asp:TextBox runat="server" ID="txtLast_DonateDateE" CssClass="font9" 
                    Width="70px" onchange="CheckDateFormat(this, '捐款日期');"></asp:TextBox>
                <img id="imgtxtLast_DonateDateE" alt="" src="../images/date.gif" />
             </td>
                         <th align="right" colspan="1">
                             捐款總金額：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="txtDonate_TotalS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="txtDonate_TotalE" CssClass="font9" 
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
                <asp:TextBox runat="server" ID="txtDonate_NoS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="txtDonate_NoE" CssClass="font9" Width="70px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款人：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="txtDonor_Name" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                             捐款人編號： 
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="txtDonor_IdS" CssClass="font9" Width="70px"></asp:TextBox>
                ~<asp:TextBox runat="server" ID="txtDonor_IdE" CssClass="font9" Width="70px"></asp:TextBox>
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
                收據編號：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxInvoice_NoS" CssClass="font9"></asp:TextBox>~
                <asp:TextBox runat="server" ID="tbxInvoice_NoE" CssClass="font9"></asp:TextBox>
                <font color="red">( 機構前置碼不需輸入 )</font>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1" >
                 捐款方式：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblDonate_Payment" runat="server"  RepeatDirection="Horizontal"></asp:CheckBoxList>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1" >
                 捐款用途：
            </th>
            <td align="left" colspan="5">
                <asp:CheckBoxList ID="cblDonate_Purpose" runat="server"  RepeatDirection="Horizontal"></asp:CheckBoxList>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1" >
                 捐款類別：
            </th>
            <td align="left" colspan="1">
                <asp:CheckBoxList ID="cblDonate_Type" runat="server"  RepeatDirection="Horizontal"></asp:CheckBoxList>
            </td>
        </tr>
        <tr>
            <td colspan="6">

                

            </td>
        </tr>
        <tr>
             <th align="right" colspan="1">
                名條用途：
            </th>
            <td align="left" colspan="5">
                 <asp:RadioButton ID="rbPrint_Type_1" runat="server" GroupName="1" text ="紙本月刊" Checked/>
                 &nbsp;&nbsp;
                 <%--20140425 修改 by Ian_Kao--%>
                 <asp:RadioButton ID="rbPrint_Type_2" runat="server" GroupName="1" text = "生日卡"/>
                 <%--<asp:TextBox ID="tbxBirth_Month" runat="server" CssClass="font9" Width="20px" text='13'></asp:TextBox>--%>
                 <asp:dropdownlist runat="server" ID="ddlBirthMonth2" CssClass="font9"></asp:dropdownlist>
                 &nbsp;&nbsp;
                 <asp:RadioButton ID="rbPrint_Type_3" runat="server" GroupName="1" text="DVD"/>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                名條內容：
            </th>
            <td align="left" colspan="1">
                 <asp:RadioButtonList ID="rblContent" runat="server" 
                    RepeatDirection="Horizontal">
                     <asp:ListItem Value="1" Selected="True">通訊地址 </asp:ListItem>
                     <asp:ListItem Value="2">收據地址</asp:ListItem>
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
            <td align="right" colspan="8" class="style1">
                <asp:Button ID="btnAddress" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="名條" onclick="btnAddress_Click"/>
                <asp:Button ID="btnPost" CssClass="npoButton npoButton_Print" runat="server"  Width="25mm"
                    Text="大宗掛號" onclick="btnPost_Click"/>
                <asp:Button ID="btnHelp" CssClass="npoButton npoButton_Single" runat="server"  Width="20mm"
                    Text="說明" OnClientClick="Help_OnClick();"/>
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="4">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
    </table>
    </div>
    </form>
    <p>
        &nbsp;</p>
</body>
</html>