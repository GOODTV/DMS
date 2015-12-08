<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateDone.aspx.cs" Inherits="Online_DonateDone" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>線上定期定額奉獻</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/jquery.address.js"></script>
    <script type="text/javascript" src="../include/jquery.field.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {

        });                     //end of ready()

    </script>
</head>
<body style="background-color: #EEE">
    <form id="Form1" runat="server">
    <asp:HiddenField ID="HFD_chkItem" runat="server" />
    <asp:HiddenField ID="HFD_ItemList" runat="server" />
    <table width="860" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td width="436">
                <img alt="aa" src="images/bar_logo.jpg" width="390" height="40" />
            </td>
            <td width="424" style="background-image: url('images/bar_item.jpg')">
                <table width="95%" border="0" align="center" cellpadding="5" cellspacing="5">
                    <tr>
                        <td align="left" valign="bottom">
                            <strong><font color="#003399">定期定額奉獻完成</font></strong>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td height="20" colspan="3">
            </td>
        </tr>
        <tr valign="center">
            <td height="405" colspan="3" style="background-image: url('images/bk_fish.jpg')"
                align="center">
                <br />
                <table width="860">
                    <tr align="center">
                        <td height="230" colspan="3" valign="bottom">
                           <asp:Label ID="lblTitle" runat="server" Style="color: #8B0000;"> 平安！
                           <p />感謝您的奉獻！我們已收到您的定期授權資料，
                           <br />將會按您所指定的授權周期進行扣款。<br />
                           <br />再一次感謝您對 GOODTV 的奉獻支持，願神大大賜福給您！</asp:Label>                
                           <br />  
                        </td>
                    </tr>
                    <tr align="center">                        
                        <td height="145" valign="bottom">
                            <font color='red' size='2'>您有任何捐款問題請和我們聯絡：
                            <br />總機：02-8024-3911<br />傳真：02-8024-3933<br />
                            地址：235 新北市中和區中正路911號6樓 捐款服務組</font>
                        </td>
                    </tr>
                </table>
                              
                 
                <%--<table width="860">
                    <tr align="center">
                        <td height="80" width="430" valign="bottom">
                           
                        </td>
                    </tr>
                </table>   --%>             
            </td>            
        </tr>       
        <tr>
            <td height="10" colspan="2">
                &nbsp;
            </td>         
        </tr>
        <tr>
            <td colspan="2" align="center">
                <asp:Button ID="btnBackDefault" class="npoButton npoButton_Submit" runat="server"
                    Text="回奉獻首頁" OnClick="btnBackDefault_Click" Width="120px" />
            </td>
        </tr>
    </table>
   <%-- <div class="function">
        <asp:Button ID="btnBackDefault" class="npoButton npoButton_Submit" runat="server"
            Text="回奉獻首頁" OnClick="btnBackDefault_Click" Width="120px" />
    </div>--%>
    </form>
</body>
</html>
