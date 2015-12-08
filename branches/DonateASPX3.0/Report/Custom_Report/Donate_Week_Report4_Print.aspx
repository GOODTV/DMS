﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Donate_Week_Report4_Print.aspx.cs" Inherits="Report_Custom_Report_Donate_Week_Report4" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html  xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>週報用大額捐款人明細</title>
    <link href="/include/main.css" rel="stylesheet" type="text/css" />
    <link href="/include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="/include/jquery-1.7.js"></script>
   <style type="text/css">
       #Form1
       {
           text-align: center;
       }
   </style>
</head>
<body>
     <script type="text/javascript">
         function SelfClose() {
             window.opener = null;
             window.open('', '_self', '');
             window.close();
         }
    </script>
   <form id="Form1" runat="server">
    <asp:HiddenField runat="server" ID="HFD_Donate_Amt_week4" />
    <asp:HiddenField runat="server" ID="HFD_DonateDateS_week4" />
    <asp:HiddenField runat="server" ID="HFD_DonateDateE_week4" />
    <asp:HiddenField runat="server" ID="HFD_AccumulateDateS_week4" />
    <asp:HiddenField runat="server" ID="HFD_AccumulateDateE_week4" />
    <asp:HiddenField runat="server" ID="HFD_Is_Abroad_week4" />
    <asp:HiddenField runat="server" ID="HFD_Is_ErrAddress_week4" />
    <asp:HiddenField runat="server" ID="HFD_Sex_week4" />
       <table id="PrnTable" width="100%" border="0" cellpadding="0" cellspacing="0" runat="server">
       <tr>
           <td align="center" style="width:1500px">
               <asp:Literal ID="GridList" runat="server"></asp:Literal>
           </td>
       </tr>
   </table>
   </form>
</body>
</html>
