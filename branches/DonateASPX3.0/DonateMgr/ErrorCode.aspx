<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ErrorCode.aspx.cs" Inherits="DonateMgr_ErrorCode" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>錯誤代碼查詢</title>
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
     <script type="text/javascript">
         function SelfClose() {
             window.opener = null;
             window.open('', '_self', '');
             window.close();
         }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table id="PrnTable" width="80%" border="0" cellpadding="0" cellspacing="0" runat="server">
       <tr>
           <td style="width:1500px; text-align: center;">
               <asp:Literal ID="GridList" runat="server"></asp:Literal>
           </td>
       </tr>
   </table>
    </div>
    </form>
</body>
</html>
