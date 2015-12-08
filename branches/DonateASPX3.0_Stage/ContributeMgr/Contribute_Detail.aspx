<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Contribute_Detail.aspx.cs" Inherits="ContributeMgr_Contribute_Detail" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>實物奉獻捐贈維護</title>
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
    </script>
    <script type="text/javascript">
        function Export_Y() {
            $('#btnEdit').hide();
            $('#btnPrint').hide();
            $('#btn_Export').hide();
            $('#btnPrint').hide();
            $('#btnRePrint').hide();
        }
        function Export_N_Invoice_Print_N() {
            $('#btn_ReExport').hide();
            $('#btnRePrint').hide();
        }
        function Export_N_Invoice_Print_Y() {
            $('#btn_ReExport').hide();
            $('#btnPrint').hide();
        }
        function Post_OnClick() {
            var HFD_Uid = document.getElementById('HFD_Uid');
            if (window.confirm('您是否確定要將查詢結果列印？') == true) {
                window.open('../ContributeMgr/Contribute_InvoicePrn.aspx?Contribute_Id=' + HFD_Uid.value, 'Contribute_InvoicePrn', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600', '');
            }
        }
    </script>
    <style type="text/css">
    </style>
    </head>
<body class="body">
    <form id="Form2" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <asp:HiddenField runat="server" ID="HFD_Donor_Id" />
    <asp:HiddenField runat="server" ID="HFD_Contribute_Amt" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        實物奉獻捐贈維護
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1">
                捐贈人：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right"colspan="1" >
                類別：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxCategory" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                收據地址：
            </th>
            <td align="left" colspan="4">
                <asp:TextBox runat="server" ID="tbxAddress" CssClass="font9"  Width="250px" 
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
            <td colspan="9">

            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                捐贈日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxContribute_Date" runat="server" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
                &nbsp;</td>
            <th align="right" colspan="1">
                捐贈方式：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxContribute_Payment" runat="server" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                捐贈用途：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxContribute_Purpose" runat="server" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                收據開立：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox ID="tbxInvoice_Type" runat="server" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td> 
        </tr> 
        <tr>
            <td colspan="8">

            </td>
        </tr>
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDept" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                收據抬頭：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxInvoice_Title" CssClass="font9" Width="210px" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td>
            <td align="center" colspan="1">
                <asp:checkbox runat="server" ID="cbxInvoice_Pre" Text="手開收據" CssClass="font9" 
                    Enabled="False" ></asp:checkbox>
            </td>
            <th align="right" colspan="1">
                收據編號：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxInvoice_No" CssClass="font9" ReadOnly="True"
                    BackColor="#FFE1AF"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                沖帳日期：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxAccoun_Date" runat="server" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
                &nbsp;</td>
            <th align="right" colspan="1">
                會計科目：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxAccounting_Title" runat="server" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="4">
                <asp:TextBox ID="tbxAct_Id" runat="server" Width="250px" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th colspan="1" align="right">
                捐款備註：
            </th>
            <td colspan="3">
                <asp:Textbox runat="server" ID="tbxComment" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True" TextMode="MultiLine" Width="450px" Height="100px"></asp:Textbox> 
            </td>
            <th colspan="1" align="right">
                收據備註：<br />
               (列印用) 
            </th>
            <td colspan="4">
                <asp:Textbox runat="server" ID="tbxInvoice_PrintComment" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True" TextMode="MultiLine" Width="450px" Height="100px"></asp:Textbox> 
            </td>
         </tr>
         <tr>
            <th colspan="1" align="right">
                捐贈內容：
            </th>
            <td  align="center" colspan="7">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
         </table>
    </div>
    <div class="function">
        <asp:Button ID="btnEdit" class="npoButton npoButton_Modify" runat="server" 
            Text="修改捐贈資料" Width="110px" onclick="btnEdit_Click"/>
        <asp:Button ID="btnPrint" class="npoButton npoButton_Print" runat="server" 
            Text="收據列印" OnClientClick="Post_OnClick();"/>
            <asp:Button ID="btnRePrint" class="npoButton npoButton_Print" runat="server" 
            Text="收據補印" OnClientClick="Post_OnClick();"/>
        <asp:Button ID="btn_Export" class="npoButton npoButton_Export" runat="server" 
            Text="收據作廢" onclick="btnExport_Click" OnClientClick= "return confirm('您是否確定要作廢收據？');"/>
        <asp:Button ID="btn_ReExport" class="npoButton npoButton_Export" runat="server" 
            Text="還原作廢收據" Width="120px" onclick="btn_ReExport_Click" OnClientClick= "return confirm('您是否確定要還原作廢收據？');" />
        <asp:Button ID="btnExit" class="npoButton npoButton_Cancel" runat="server"    
            Text="取消" Width="80px" onclick="btnExit_Click" />
        <asp:Button ID="btn_Add_self" class="npoButton npoButton_New" runat="server" 
            Text="續增本人捐贈資料" Width="140px" onclick="btn_Add_self_Click" />
        <asp:Button ID="btn_Add_other" class="npoButton npoButton_New" runat="server" 
            Text="新增他人捐贈資料" Width="140px" onclick="btn_Add_other_Click"/>
        <br />
    </div>
    </form>
</body>
</html>
