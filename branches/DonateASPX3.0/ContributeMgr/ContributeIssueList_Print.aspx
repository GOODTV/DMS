﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ContributeIssueList_Print.aspx.cs" Inherits="ContributeMgr_ContributeIssue_Print" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html  xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>實物奉獻領用作業</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
   <style type="text/css">
       @media print
       {
           .btnPrint
           {
               display: none;
           }
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
   <table id="PrnTable" width="100%" border="0" cellpadding="0" cellspacing="0" runat="server">
<%--       <tr>
           <td align="right" style="font-size: 16px">
               <input id="btnPrint" class="btnPrint" type="button" value="列印本頁" onclick=" window.print(); " />
           </td>
       </tr>--%>
       <tr>
           <td align="center" style="width:800px">
               <asp:Literal ID="GridList" runat="server"></asp:Literal>
           </td>
       </tr>
   </table>
   </form>
</body>
</html>
