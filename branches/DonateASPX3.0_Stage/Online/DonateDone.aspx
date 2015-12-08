<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateDone.aspx.cs" Inherits="Online_DonateDone" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>線上定期定額奉獻</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            if ($("#HFD_Question").val() != "") {
                $('.Questionaire').show();
            } else {
                $('.Questionaire').hide();
            }
        });
    </script>
    <style type="text/css">

        body {
            font-size: 18px;
        }

        .table_h tr {
            color: black;
            height: 30px;
        }

        .table_h tr:hover {
            color: black;
            background-color: white;
        }

    </style>
</head>
<body style="background-color: #EEE">
    <form id="Form1" runat="server">
    <asp:HiddenField ID="HFD_UID" runat="server" />
    <asp:HiddenField ID="HFD_Question" runat="server" value=""/>
    <table width="860" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td colspan="2">
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
                        <td align="right" valign="bottom">
                            <!--<strong><font color="#003399">Language：</font><a href="../English/DonateDone.aspx" target="_self"><font size="3">English</font></a></strong>-->
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td height="20" colspan="2">
            </td>
        </tr>
        <tr valign="top">
            <td colspan="2" align="center" style="white-space: nowrap;">                                                    
                <br />
                <span>我目前的位置：</span>
                <asp:Label ID="lblStep1" runat="server" Text=" 我的奉獻明細 >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep2" runat="server" Text=" 填寫天使資料 >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep3" runat="server" Text=" 確認奉獻明細 >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep4" runat="server" Text=" 填寫定期授權資料 >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep5" runat="server" Text=" 確認奉獻捐款 " ForeColor="Red" Font-Bold="True" />
            </td>
        </tr>
        <tr valign="middle">
            <td height="405" colspan="2" align="center">
                <br />
                <table width="860">
                    <tr align="center">
                        <td>
                           <span style="color: #0000CC; font-family:微軟正黑體;">親愛的 <asp:Label ID="lblTitle" runat="server" Style="color: #0000CC;"/> 平安！
                           <br />感謝您的奉獻！我們已收到您的定期授權資料，
                           <br />將會按您所指定的授權周期進行扣款。</span>   
                           <br /><br />
                           <table align="center" border="0" cellpadding="0" cellspacing="1" width="80%">
                               <tr>
                                   <td>
                                       <asp:Label ID="lblDonateContent" runat="server"></asp:Label>
                                   </td>
                               </tr>
                           </table>
                           
                           <br /><asp:Label ID="lblend" runat="server" Style="color: #8B0000;">再一次感謝您對 GOODTV 的奉獻支持，願神大大賜福給您！</asp:Label>
                        </td>
                    </tr>
                    <tr align="center">                        
                        <td>
                            <font size='3'><br /><br />有任何關於定期授權問題請和我們聯絡：
                            <br /><br />總機：02-8024-3911&nbsp;&nbsp;傳真：02-8024-3933<br /><br />
                            地址：235 新北市中和區中正路911號6樓 捐款服務組</font>
                        </td>
                    </tr>
                </table>
  
            </td>            
        </tr>       
        <tr>
            <td height="10" colspan="2">
                &nbsp;
            </td>         
        </tr>
        <tr class="Questionaire" style="display: none;" align="center">
            <td colspan="2">
               <table width="530" border="0" cellpadding="0" cellspacing="1" class="table_v2">
                    <tr>
                        <th colspan="3"><font color="blue">為更有效服務捐款人，請務必勾選下列問題，或惠賜寶貴意見</font></th>
                    </tr>
                    <tr>
                        <th align="center" style="width:150px;">奉獻動機(可複選)：</th>
                        <td>
                            <asp:CheckBox ID="DonateMotive1" runat="server" Text="支持媒體宣教大平台，可廣傳福音" /><br />
                            <asp:CheckBox ID="DonateMotive2" runat="server" Text="個人靈命得造就" /><br />
                            <asp:CheckBox ID="DonateMotive3" runat="server" Text="支持優質節目製作" /><br />
                            <asp:CheckBox ID="DonateMotive4" runat="server" Text="支持GOOD TV家庭事工" /><br />
                            <%--<asp:CheckBox ID="DonateMotive5" runat="server" Text="抵扣稅額" /><br />--%>
                            <asp:CheckBox ID="DonateMotive5" runat="server" Text="感恩奉獻" /><br />
                            <asp:CheckBox ID="DonateMotive6" runat="server" Text="其他" />(如有請寫入〝想對GOODTV說的話〞)
                        </td>
                        </tr>
                        <tr>
                        <th align="center" style="width:150px;">收看管道(可複選)：</th>
                        <td>
                            <asp:CheckBox ID="WatchMode1" runat="server" Text="GOOD TV電視頻道" />
                            <br />
                            <asp:CheckBox ID="WatchMode9" runat="server" Text="報章雜誌" />
                            <br />新媒體平台：
                            <asp:CheckBox ID="WatchMode2" runat="server" Text="官網" />
                            <asp:CheckBox ID="WatchMode3" runat="server" Text="Facebook" />
                            <asp:CheckBox ID="WatchMode4" runat="server" Text="Youtube" />
                            <br />平面：　　　
                            <asp:CheckBox ID="WatchMode5" runat="server" Text="好消息月刊" />
                            <asp:CheckBox ID="WatchMode6" runat="server" Text="GOOD TV簡介刊物" />
                            <br />親友推薦：　
                            <asp:CheckBox ID="WatchMode7" runat="server" Text="教會牧者" />
                            <asp:CheckBox ID="WatchMode8" runat="server" Text="親友" />
                        </td>
                    </tr>
                   <tr>
                        <th align="right" nowrap="nowrap">想對GOOD TV說的話：
                        </th>
                        <td>
                            <asp:TextBox ID="txtToGoodTV" runat="server" Width="90%" TextMode="MultiLine" Rows="5" Font-Size="16px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" align="center">
                            <asp:Button ID="btnSubmit" runat="server" Text="感謝您提供寶貴意見，送出"
                                OnClick="btnSubmit_Click" />
                        </td>
                    </tr>
                </table>         
            </td>
        </tr>
        <tr>
            <td colspan="2">
               &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center">
                <asp:Button ID="btnBackDefault" class="css_btn_class" runat="server" Text="回奉獻首頁" OnClick="btnBackDefault_Click" />
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
