<%@ Page Language="C#" AutoEventWireup="true" CodeFile="GroupClassQry.aspx.cs" Inherits="DonorGroupMgr_GroupClassQry" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>群組類別設定</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function() {
            InitMenu();
            $("#drp_CaseType").bind("change", function(e) {
                ChangeSubType(e, 'drp_CaseTypeList');
            });

            $("#drp_CaseSpecies").bind("change", function(e) {
                ChangeCaseSpeciesList(e, 'drp_CaseSpeciesList');
            });
        });	
    </script>
</head>
<body class="body">
    <form id="Form1" name="form" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
        <h1>
            <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
            群組類別設定
        </h1>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
            <tr style="height: 30px">
                <th align="right" width="10%">群組類別：</th>
                <td align="left" nowrap="nowrap">
                    <asp:TextBox runat="server" ID="txtGroupClassName" Width="60mm" CssClass="font9"></asp:TextBox>
                    <span style="color:red; font-weight: bold; font-size: medium;">輸入名稱如家庭，教會等</span>
                 </td>
                 <td rowspan="2" align="right">
                     <asp:Button ID="btnQuery" class="npoButton npoButton_Search" runat="server" Text="查詢" OnClick="btnQuery_Click"/>
                     <asp:Button ID="btnAdd" class="npoButton npoButton_New" runat="server" Text="新增" OnClick="btnAdd_Click"/>
                 </td>
            </tr>
            <tr style="height: 30px">
                 <th width="10%" align="right">備註：</th>
                 <td colspan="1" align="left">
                     <asp:TextBox runat="server" ID="txtSupplement" Width="60mm" CssClass="font9"></asp:TextBox>
                 </td>
            </tr>
            <tr>
                <td colspan="3">
                    <span style="color:red; font-weight: bold; font-size: medium;">
                        如何設定群組類別：
                        <br />1.要把一些天使關聯起來，要幫他們先設想一個關聯起來後這個群體的類別如家庭或顧問等
                        <br />2.按新增並依畫面填入群組類別名稱儲存即可生效
                        <br />3.可以先查詢看看有哪些已定義類別參考
                    </span>
                </td>
            </tr>
            <tr>
                <td width="100%" colspan="3">
                    <br/>
                    <asp:Label ID="lblGridList" runat="server" Text=""></asp:Label>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
