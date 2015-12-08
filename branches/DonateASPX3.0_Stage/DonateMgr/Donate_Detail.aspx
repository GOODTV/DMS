<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Donate_Detail.aspx.cs" Inherits="DonateMgr_Donate_Detail" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>捐款資料維護</title>
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
        function Export_Y() {
            $('#btnEdit').hide();
            $('#btnPrint').hide();
            $('#btnRePrint').hide();
            $('#btn_Export').hide();
            $('#btn_Re_No').hide();
        }
        //20140509 修改 by Ian_Kao 判斷是否是補印收據
        function Export_N(Invoice_Print) {
            if (Invoice_Print == 1) {
                $('#btnRePrint').hide();
            }
            else if (Invoice_Print == 2) {
                $('#btnPrint').hide();
            }
            $('#btn_ReExport').hide();
        }
        //20140515 新增 by Ian_Kao 若已關帳則不能修改和作廢，只能列印或補印
        function Close(Invoice_Print) {
            $('#btnEdit').hide();
            $('#btn_Export').hide();
            $('#btn_ReExport').hide();
            $('#btn_Re_No').hide();
            if (Invoice_Print == 1) {
                $('#btnRePrint').hide();
            }
            else if (Invoice_Print == 2) {
                $('#btnPrint').hide();
            }
        }
        //--------------------------------------------------------------------
        function Post_OnClick() {
            var HFD_Uid = document.getElementById('HFD_Uid'); 
            if (window.confirm('您是否確定要將查詢結果列印？') == true) {
                window.open('../DonateMgr/Donate_InvoicePrn.aspx?Donate_Id='+ HFD_Uid.value, 'Donate_InvoicePrn', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
            }
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
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <asp:HiddenField runat="server" ID="HFD_Donor_Id" />
    <asp:HiddenField runat="server" ID="HFD_Donate_Accou" />
    <asp:HiddenField runat="server" ID="HFD_Invoice_No" />
    <asp:HiddenField runat="server" ID="HFD_Donate_Amt_Total" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        收據維護紀錄 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <%--<tr>
            <th align="right" colspan="1">
                收據編號：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxInvoice_No_Top" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>--%>
        <tr>
            <th align="right" colspan="1">
                捐款人：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <%--<th align="right"colspan="1" >
                類別：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxCategory" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>--%>
            <th align="right"colspan="1" >
                捐款人編號：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxDonor_Id" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                收據地址：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxAddress" CssClass="font9"  Width="350px" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                身分別：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxDonor_Type" CssClass="font9"
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                身分證/統編：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIDNo" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        
        <tr>
            <td colspan="8">

            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Date" runat="server" onchange="CheckDateFormat(this, '捐款日期');"
                BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
                &nbsp;</td>
            <th align="right" colspan="1">
                捐款方式：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Payment" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                捐款用途：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Purpose" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                收據開立：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxInvoice_Type" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐款金額：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Amt" runat="server" OnPropertyChange="SumAccou()" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                手續費：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Fee" Text="0" runat="server" OnPropertyChange="SumAccou()" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                實收金額：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Accou" Text="0" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <asp:Panel ID="PIEPayType" runat="server" Visible ="false">
            <th align="right" colspan="1">
                線上付款方式：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxIEPayType" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
           </asp:Panel>
        </tr>
        <tr>
            <th align="right" colspan="1">
                外幣幣別：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Forign" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                外幣金額：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox ID="tbxDonate_ForignAmt" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <asp:Panel ID="PIEPayOrderid" runat="server" Visible ="false">
            <th align="right" colspan="1">
                線上訂單編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxOrderid" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
           </asp:Panel>
        </tr>
        <tr>
            <td colspan="8">

            </td>
        </tr>
        <asp:Panel ID="PanelCheck" runat="server"  Visible ="false">
        <tr>
            <th align="right"colspan="1" >
                支票號碼：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCheck_No" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                到期日：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxCheck_ExpireDate" runat="server" onchange="CheckDateFormat(this, '到期日');"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <td colspan="4">
                
            </td>
        </tr>
        </asp:Panel>
        <asp:Panel ID="PanelCreditCard" runat="server" Visible ="false">
        <tr>
            <th align="right" colspan="1">
                銀行別：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCard_Bank" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                信用卡別：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCard_Type" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                信用卡號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxAccount_No1" CssClass="font9" Width="28px"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>-
                <asp:TextBox runat="server" ID="tbxAccount_No2" CssClass="font9" Width="28px"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>-
                <asp:TextBox runat="server" ID="tbxAccount_No3" CssClass="font9" Width="28px"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>-
                <asp:TextBox runat="server" ID="tbxAccount_No4" CssClass="font9" Width="28px"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                有效月年：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxValid_Date" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                持卡人：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCard_Owner" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True" ></asp:TextBox>
                <asp:checkbox runat="server" ID="cbxSame_Owner" Text="同捐款人" CssClass="font9" Enabled="False"></asp:checkbox>
            </td>
            <th align="right" colspan="1">
                持卡人身分證：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxOwner_IDNo" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                與捐款人關係：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxRelation" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                授權碼：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxAuthorize" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        </asp:Panel>
        <asp:Panel ID="PanelAccount" runat="server"  Visible ="false">
        <tr>
             <th align="right" colspan="1">
                存簿戶名：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_Name" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
                <asp:checkbox runat="server" ID="cbxSame_Owner2" Text="同捐款人" CssClass="font9"  Enabled="False"></asp:checkbox>
            </td>
            <th align="right" colspan="1">
                持有人身分證：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_IDNo" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                存簿局號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_SavingsNo" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                存簿帳號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_AccountNo" CssClass="font9"  BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        </asp:Panel>
        <tr>
            <td colspan="8">

            </td>
        </tr>
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDept" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                收據抬頭：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxInvoice_Title" CssClass="font9" Width="300px" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <td align="center" colspan="1">
                <asp:checkbox runat="server" ID="cbxInvoice_Pre" Text="手開收據" CssClass="font9" 
                    OnClick="cbxInvoice_Pre_OnClick()" Enabled="False"></asp:checkbox>
                
            </td>
            <th align="right" colspan="1">
                收據編號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxInvoice_No" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                請款日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxRequest_Date" runat="server" onchange="CheckDateFormat(this, '請款日期');"
                BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
                &nbsp;</td>
            <th align="right" colspan="1">
                入帳銀行：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxAccoun_Bank" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                沖帳日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxAccoun_Date" runat="server" onchange="CheckDateFormat(this, '沖帳日期');"
                BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
                &nbsp;</td>
            <th align="right" colspan="1">
                捐款類別：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxDonate_Type" runat="server" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                傳票號碼：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonation_NumberNo" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                劃撥 / 匯款單號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonation_SubPoenaNo" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                會計科目：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxAccounting_Title" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                收據寄送：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxInvoceSend_Date" runat="server" onchange="CheckDateFormat(this, '收據寄送');"
                BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
                &nbsp;</td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                專案活動：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxAct_Id" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
         </tr>
         <tr>
            <th colspan="1" align="right">
                捐款備註：
            </th>
            <td colspan="3">
                <asp:Textbox runat="server" ID="tbxComment" CssClass="font9" 
                    TextMode="MultiLine" Width="450px" Height="100px" BackColor="#FFE1AF" ReadOnly="True"></asp:Textbox> 
            </td>
            <th colspan="1" align="right">
                收據備註：<br />
               (列印用) 
            </th>
            <td colspan="3">
                <asp:Textbox runat="server" ID="tbxInvoice_PrintComment" CssClass="font9" 
                    TextMode="MultiLine" Width="450px" Height="100px" BackColor="#FFE1AF" ReadOnly="True"></asp:Textbox> 
            </td>
         </tr>
         <tr>
            <th align="right" colspan="1">
                首印日期：
            </th>
            <td align="left">
                <asp:TextBox runat="server" ID="tbxInvoice_Print_Date" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
             <th align="right" colspan="1">
                首印經手人：
            </th>
            <td align="left">
                <asp:TextBox runat="server" ID="tbxInvoice_FirstPrint_User" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                最後補印日期：
            </th>
            <td align="left">
                <asp:TextBox runat="server" ID="tbxInvoice_RePrint_Date" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
             <th align="right" colspan="1">
                最後補印&nbsp;&nbsp;&nbsp;&nbsp;<br />經手人：
            </th>
            <td align="left">
                <asp:TextBox runat="server" ID="tbxInvoice_LastPrint_User" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                年度證明&nbsp;&nbsp;&nbsp;&nbsp;<br />首印日期：
            </th>
            <td align="left">
                <asp:TextBox runat="server" ID="tbxInvoice_Yearly_Print_Date" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
             <th align="right" colspan="1">
                首印經手人：
            </th>
            <td align="left">
                <asp:TextBox runat="server" ID="tbxInvoice_Yearly_FirstPrint_User" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                最後補印日期：
            </th>
            <td align="left">
                <asp:TextBox runat="server" ID="tbxInvoice_Yearly_RePrint_Date" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
             <th align="right" colspan="1">
                最後補印&nbsp;&nbsp;&nbsp;&nbsp;<br />經手人：
            </th>
            <td align="left">
                <asp:TextBox runat="server" ID="tbxInvoice_Yearly_LastPrint_User" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                資料建檔&nbsp;&nbsp;&nbsp;&nbsp;<br />人員：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCreate_User" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                資料建檔日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCreate_Date" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                最後異動人員：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxLastUpdate_User" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                最後異動&nbsp;&nbsp;&nbsp;&nbsp;<br />日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxLastUpdate_Date" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
    </table>
    </div>
    <div class="function">
        <asp:Button ID="btnEdit" class="npoButton npoButton_Modify" runat="server" 
            Text="修改捐款資料" Width="110px" onclick="btnEdit_Click"/>
        <asp:Button ID="btnPrint" class="npoButton npoButton_Print" runat="server" 
            Text="收據列印" OnClientClick = "Post_OnClick();"/>
        <asp:Button ID="btnRePrint" class="npoButton npoButton_Print" runat="server" 
            Text="收據補印" OnClientClick = "Post_OnClick();"/>
        <asp:Button ID="btn_Export" class="npoButton npoButton_Export" runat="server" 
            Text="收據作廢" onclick="btnExport_Click" OnClientClick= "return confirm('您是否確定要作廢收據？');"/>
        <asp:Button ID="btn_ReExport" class="npoButton npoButton_Export" runat="server" 
            Text="還原作廢收據" Width="120px" onclick="btn_ReExport_Click" OnClientClick= "return confirm('您是否確定要還原作廢收據？');" />
        <asp:Button ID="btn_Re_No" class="npoButton npoButton_Modify" runat="server" 
            Text="重新取號" onclick="btn_Re_No_Click"/>
        <asp:Button ID="btnExit" class="npoButton npoButton_Cancel" runat="server"    
            Text="取消" Width="80px" onclick="btnExit_Click" />
        
        <asp:Button ID="btn_Add_self" class="npoButton npoButton_New" runat="server" 
            Text="續增本人捐贈資料" Width="140px" onclick="btn_Add_self_Click" />
        <asp:Button ID="btn_Add_other" class="npoButton npoButton_New" runat="server" 
            Text="新增他人捐款資料" Width="140px" onclick="btn_Add_other_Click"/>
        <br />
    </div>
    </form>
</body>
</html>
