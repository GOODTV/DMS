<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Invoice_Yearly_Print_Data.aspx.cs" Inherits="DonateMgr_Invoice_Yearly_Print_Address" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html  xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>年度捐款證明列印</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <%--<OBJECT id=factory style="DISPLAY: none" codebase="http://mis.npois.com.tw/smsx.cab" classid=clsid:1663ed61-23eb-11d2-b92f-008048fdd814 VIEWASTEXT>
    </OBJECT>
    <script type="text/javascript">
        function window.onload(){
          factory.printing.header='' ;          //頁首
          factory.printing.footer='' ;          //頁尾
          factory.printing.portrait = true ;    //直印(true),橫印(false)
          factory.printing.leftMargin = 8.0 ;   //左邊界
          factory.printing.topMargin = 8.0 ;    //上邊界
          factory.printing.rightMargin = 8.0 ;  //右邊界
          factory.printing.bottomMargin = 8.0 ; //下邊界
          factory.printing.print();
        }
    </script>--%>
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
<%--   <table id="PrnTable"  align="center" width="800" border="0" cellpadding="1" cellspacing="0" runat="server">
       <tr>
           <td align="center">
               
           </td>
       </tr>
   </table>--%>

               <asp:Literal ID="GridList" runat="server" ></asp:Literal>
 
   </form>
</body>
</html>
