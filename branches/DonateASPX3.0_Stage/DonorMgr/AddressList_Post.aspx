<%@ Page Language="C#" AutoEventWireup="true" CodeFile="AddressList_Post.aspx.cs" Inherits="DonorMgr_AddressList_Post" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html  xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>郵局掛號函件執據</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <object id="factory" viewastext style="display:none" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="http://mis.npois.com.tw/smsx.cab"></object>
    <script language="javascript">
        function window.onload() {
            factory.printing.header = '';          //頁首
            factory.printing.footer = '';          //頁尾
            factory.printing.portrait = true;    //直印(true),橫印(false)
            factory.printing.leftMargin = 8.0;   //左邊界
            factory.printing.topMargin = 8.0;    //上邊界
            factory.printing.rightMargin = 8.0;  //右邊界
            factory.printing.bottomMargin = 8.0; //下邊界
            //factory.printing.print();
        }
    </script>
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
           text-align: left;
       }
      .style1 {font-size: 15px}
      .style3 {font-size: 10pt}
      .PageBreak { page-break-after:always; }
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