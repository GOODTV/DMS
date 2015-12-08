<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Questionnaire_Add.aspx.cs" Inherits="DonateMgr_Questionnaire_Add" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>問卷資料維護【新增】</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        問卷維護紀錄【新增】 
    </h1>
    <table width="100%" border="0" cellpadding="0" cellspacing="1" class="table_v">
        <tr>
            <th align="right" colspan="1">
                捐款人：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                捐款人編號：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonor_Id" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                收據地址：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxAddress" CssClass="font9"  Width="350px" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                身分別：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxDonor_Type" CssClass="font9"
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                收據開立：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxInvoice_Type" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                身分證/統編：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIDNo" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                最近捐款日：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxLast_DonateDate" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
         <tr>
            <td colspan="4">
               <table width="530" border="0" cellpadding="0" cellspacing="1" class="table_v">
                    <tr>
                        <th colspan="3"><font color="blue">奉獻動機與收看管道調查</font></th>
                    </tr>
                    <tr>
                        <th align="right">奉獻動機(可複選)：</th>
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
                        <th align="right">收看管道(可複選)：</th>
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
                            <asp:TextBox ID="txtToGoodTV" runat="server" Width="90%" TextMode="MultiLine" Rows="5" Font-Size="14px"></asp:TextBox>
                        </td>
                    </tr>
                   <tr>
                       <th align="right" nowrap="nowrap">問卷來源：
                        </th>
                        <td>
                            <asp:TextBox ID="tbxDonateWay" runat="server" Width="100px" Text="電話客服" BackColor="#FFE1AF" ReadOnly="true"></asp:TextBox>
                        </td>
                   </tr>
                </table>         
            </td>
        </tr>
        <tr>
            <td colspan="8" align="center">
                <asp:Button ID="btnSubmit" class="npoButton npoButton_New" runat="server" Text="存檔" OnClick="btnSubmit_Click" />
                <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" Text="取消" onclick="btnExit_Click"/>
            </td>
            </tr>    
        </table>         
    </div>
    </form>
</body>
</html>
