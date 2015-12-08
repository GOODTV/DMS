<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonateOnline.aspx.cs" Inherits="Online_DonateOnline" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>線上奉獻</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {

//            //所有TextBox的開頭是txt都限制輸入數字，除了txtDataDate和txtDeptName
//            $('input:text[id*=txt]').not('').bind("keyup", function (e) {
//                var val = $(this).val();
//                if (isNaN(val)) {
//                    val = val.replace(/[^0-9\.]/g, '');
//                    $(this).val(val);
//                }
//            });

        });      //end of ready()

    </script>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <h1>
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        <asp:Literal ID="litTitle" runat="server">Good TV 線上奉獻</asp:Literal>
    </h1>
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <table align="center" border="0" cellpadding="0" cellspacing="1" 
                class="table_v" width="100%">
                <tr>
                    <td align="center" colspan="2" 
                        style="background-color: #D3D3D3; color: Black; font-weight: bold">
                        線 上 奉 獻 
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="color: #8B0000;">
                        平安！歡迎您使用線上奉獻，請選擇奉獻項目： 
                    </td>
                </tr>
                <tr>
                    <th align="right" style="width:30mm">
                        奉獻項目： 
                    </th>
                    <td valign="middle">
                        <asp:CheckBox ID="chkItem1" Text="經常奉獻" runat="server" AutoPostBack="True" 
                            oncheckedchanged="chkItem1_CheckedChanged" />
                        ：新台幣
                        <asp:TextBox ID="txtAmount1" runat="server" Width="20mm" 
                            style="text-align:right" ontextchanged="txtAmount1_TextChanged" 
                            AutoPostBack="True" Height="19px"></asp:TextBox>
                        元
                        <asp:CompareValidator ID="cvAmount1" runat="server"
                          ErrorMessage="只允許數字" Operator="DataTypeCheck"
                          ControlToValidate="txtAmount1" Type="Integer" Display="Dynamic">
                        </asp:CompareValidator>
                        <br />
                        <asp:CheckBox ID="chkItem2" Text="優質媒體奉獻" runat="server" AutoPostBack="True" 
                            oncheckedchanged="chkItem2_CheckedChanged" />
                        ：新台幣
                        <asp:TextBox ID="txtAmount2" runat="server" Width="20mm" 
                            style="text-align:right" AutoPostBack="True" 
                            ontextchanged="txtAmount2_TextChanged"></asp:TextBox>
                        元
                        <asp:CompareValidator ID="cvAmount2" runat="server"
                          ErrorMessage="只允許數字" Operator="DataTypeCheck"
                          ControlToValidate="txtAmount2" Type="Integer" Display="Dynamic">
                        </asp:CompareValidator>
                        <br />
                        <asp:CheckBox ID="chkItem3" Text="布永康事工" runat="server" AutoPostBack="True" 
                            oncheckedchanged="chkItem3_CheckedChanged" />
                        ：新台幣
                        <asp:TextBox ID="txtAmount3" runat="server" Width="20mm" 
                            style="text-align:right" AutoPostBack="True" 
                            ontextchanged="txtAmount3_TextChanged"></asp:TextBox>
                        元
                        <asp:CompareValidator ID="cvAmount3" runat="server"
                          ErrorMessage="只允許數字" Operator="DataTypeCheck"
                          ControlToValidate="txtAmount3" Type="Integer" Display="Dynamic">
                        </asp:CompareValidator>
                        <br />
                        <asp:CheckBox ID="chkItem4" Text="林書豪佈道會" runat="server" AutoPostBack="True" 
                            oncheckedchanged="chkItem4_CheckedChanged" />
                        ：新台幣
                        <asp:TextBox ID="txtAmount4" runat="server" Width="20mm" 
                            style="text-align:right" AutoPostBack="True" 
                            ontextchanged="txtAmount4_TextChanged"></asp:TextBox>
                        元
                        <asp:CompareValidator ID="cvAmount4" runat="server"
                          ErrorMessage="只允許數字" Operator="DataTypeCheck"
                          ControlToValidate="txtAmount4" Type="Integer" Display="Dynamic">
                        </asp:CompareValidator>
                        <br />
                        <asp:CheckBox ID="chkItem5" Text="以色列事工" runat="server" AutoPostBack="True" 
                            oncheckedchanged="chkItem5_CheckedChanged" />
                        ：新台幣
                        <asp:TextBox ID="txtAmount5" runat="server" Width="20mm" 
                            style="text-align:right" AutoPostBack="True" 
                            ontextchanged="txtAmount5_TextChanged"></asp:TextBox>
                        元
                        <asp:CompareValidator ID="cvAmount5" runat="server"
                          ErrorMessage="只允許數字" Operator="DataTypeCheck"
                          ControlToValidate="txtAmount5" Type="Integer" Display="Dynamic">
                        </asp:CompareValidator>
                        <br/>
                        <br/>
                        <span  style="color: #8B0000;">
                        備註
                        <br/>
                        1：線上刷卡、WEB ATM，每筆金額最低 200 元，最高60000元。
                        <br/>
                        2：7-11、ibon 、便利商店等電子帳單，每筆金額最低 500 元。 
                        <br/>
                        3：奉獻金額請填寫純數字如 10000 
                        </span>
                    </td>
                </tr>
            </table>
            <br/>
            <table align="center" border="0" cellpadding="0" cellspacing="0" width="50%">
                <tr>
                    <td>
                        <asp:Label ID="lblGridList" runat="server"></asp:Label>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    </form>
</body>
</html>
