<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Gift_Detail.aspx.cs" Inherits="DonorMgr_Gift_Detail" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>公關贈品維護</title>
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
        公關贈品維護
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
            <%--<th align="right"colspan="1" >
                類別：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxCategory" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>--%>
            
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
       <%-- <tr>
            <th align="right" colspan="1">
                收據地址：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxAddress" CssClass="font9"  Width="250px" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>--%>
        <tr>
            <td colspan="4">

            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                贈送日期：
            </th>
            <td>
                <asp:TextBox runat="server" ID="tbxContribute_Date" CssClass="font9" BackColor="#FFE1AF" 
                        ReadOnly="True"></asp:TextBox>
            </td>
        </tr> 
        <tr>
            <td colspan="4">

            </td>
        </tr>
        <tr>
            <th colspan="1" align="right">
                備註：
            </th>
            <td colspan="3">
                <asp:Textbox runat="server" ID="tbxComment" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True" TextMode="MultiLine" Width="450px" Height="100px"></asp:Textbox> 
            </td>
         </tr>
         <tr>
            <th colspan="1" align="right">
                捐贈內容：
            </th>
            <td  align="center" colspan="3">
                 <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
            </td>
        </tr>
         </table>
    </div>
    <div class="function">
        <asp:Button ID="btnEdit" class="npoButton npoButton_Modify" runat="server" 
            Text="修改贈品資料" Width="110px" onclick="btnEdit_Click"/>
        <asp:Button ID="btnExit" class="npoButton npoButton_Cancel" runat="server"    
            Text="取消" Width="80px" onclick="btnExit_Click" />
    </div>
    </form>
</body>
</html>
