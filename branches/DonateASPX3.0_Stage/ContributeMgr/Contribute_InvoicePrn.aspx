<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Contribute_InvoicePrn.aspx.cs" Inherits="ContributeMgr_Contribute_InvoicePrn" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>標準三聯式單筆收據列印</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <object id="factory" viewastext style="display:none" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="http://mis.npois.com.tw/smsx.cab"></object>
    <script type="text/javascript">
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
       table.TableGrid{
	    font-size: 12px;
	    width:194mm;
        }
        table.TableGrid1{
	        width:194mm;
	        height:86mm;
        }
        table.TableGrid2{
	        width:194mm;
	        height:84mm;
        }
        table.TableGrid3{
	        width:194mm;
	        height:84mm;
        }    
        .Comp_Name{ 
	        width:194mm;
	        height:12mm;
          font-size:20px;
          font-family:"標楷體";
        }
        .Invoice_No{
          height:5mm;
          font-size:15px;
          font-family:"標楷體";
        }
        .Donate_Desc{
          height:8mm;
          font-size:15px;
          font-family:"標楷體";
        }
        .Donate_Right{
          font-size:12px;
          font-family:"標楷體";
        }
        .Donate_Seal{
          height:10mm;
          font-size:15px;
          font-family:"標楷體";
        }    
        .Donate_Account{
          height:8mm;
          font-size:15px;
          font-family:"標楷體";
        }    
        .Donate_Foot{
          height:5mm;
          font-size:12px;
          font-family:"標楷體";
        }
        .Line{
          border-bottom: 1px dashed #000000;
          height:3mm;
        }
        .CellMiddle{
          height:8mm;
        }
        .PageBreak {
          page-break-after:always;
        }
        .GridList
        {
            margin-bottom:8mm;
            margin-left:8mm;
            margin-right:8mm;
            margin-top:8mm;
        }
   </style>
</head>
   <body class=tool>
   <form id="Form1" runat="server"  class="GridList">
<%--<!--input id="btnPrint" class="btnPrint" type="button" value="列印本頁" onclick=" window.print(); " /-->--%>
       <asp:Literal ID="GridList" runat="server"></asp:Literal>
   </form>
</body>
</html>