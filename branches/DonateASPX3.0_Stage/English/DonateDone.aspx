<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateDone.aspx.cs" Inherits="Online_DonateDone" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Recurring Donation Accepted</title>
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
            $('#lblDonateContent td:nth-child(1)').css('text-align', 'center');
            $('#lblDonateContent td:nth-child(2)').css('text-align', 'center');
            $('#lblDonateContent td:nth-child(3)').css('text-align', 'right');
        });                     //end of ready()

    </script>
    <style type="text/css">
    td,th {
    font-size: 14px;
    color: #000;
    }
    </style>
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
                            <strong><font color="#003399">Recurring Donation Accepted</font></strong>
                        </td>
                        <td align="right" valign="bottom">
                            <strong><font color="#003399">language：</font><a href="../Online/DonateDone.aspx" target="_self"><font size="3">中文</font></a></strong>
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
                        <%--<td height="230" colspan="3" valign="bottom">
                           <span Style="color: #8B0000;">親愛的 </span><asp:Label ID="lblTitle" runat="server" Style="color: #8B0000;"> 平安！
                           <p />感謝您的奉獻！我們已收到您的定期授權資料，
                           <br />將會按您所指定的授權周期進行扣款。<br /><br /></asp:Label>   
                           <asp:Label ID="lblDonateContent" runat="server" Style="color:red;font-weight: bold;"></asp:Label>
                           <br /><asp:Label ID="lblend" runat="server" Style="color: #8B0000;">再一次感謝您對 GOODTV 的奉獻支持，願神大大賜福給您！</asp:Label>    
                        </td>--%>
                        <td height="230" colspan="3" valign="bottom">
                           <span Style="color: #0000CC; font-size: medium; font-family:微軟正黑體;">Dear <asp:Label ID="lblTitle" runat="server" Style="color: #0000CC;">,
                           <p />We have acknoledged the acceptence of the plan of your recurring donation.
                           <br />THe initial payment will be processed on the designated date shortly.</asp:Label></span>   
                           <table align="center" border="0" cellpadding="0" cellspacing="1" class="table_v2" width="70%">
                               <tr>
                                   <td>
                                       <asp:Label ID="lblDonateContent" runat="server"  Style="font-weight: bold;"></asp:Label>
                                   </td>
                               </tr>
                           </table>
                           
                           <br /><asp:Label ID="lblend" runat="server" Style="color: #8B0000; font-family:新細明體;">Thank you for your support. May God bless you greatly!</asp:Label>    
                        </td>
                    </tr>
                    <tr align="center">                        
                        <td height="145" valign="bottom">
                            <font color='red' size='2'>If you have any questions regarding your donation, Please call us at:
                            <br />Tel：02-8024-3911<br />Fax：02-8024-3933<br />
                            6F., No.911 Zhongzheng Rd., Zhonghe Dist., New Taipei City 23544, Taiwan, R.O.C.</font>
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
                <asp:Button ID="btnBackDefault" class="Online Online_Button5" runat="server"
                    Text="Back to Home" OnClick="btnBackDefault_Click" Width="174px" />
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
