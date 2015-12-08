<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Invoice_Yearly_Print_Style5.aspx.cs" Inherits="DonateMgr_Invoice_Year_Print_Invoice_Year_Print_Style5" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html  xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>年度捐款證明地址名條列印</title>
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
<%--<!--input id="btnPrint" class="btnPrint" type="button" value="列印本頁" onclick=" window.print(); " /-->--%>
       <asp:Literal ID="GridList" runat="server"></asp:Literal>
   </form>
</body>
</html>
