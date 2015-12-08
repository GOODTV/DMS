<%@ Page Language="C#" AutoEventWireup="true" CodeFile="GroupItemQry.aspx.cs" Inherits="DonorGroupMgr_GroupItemQry" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript">


        $(document).ready(function() {

            // menu控制
            InitMenu();

        });

        function CheckField() {
            var tbxDonor_Id = document.getElementById('txtDonorID');

            if (isNaN(Number(tbxDonor_Id.value)) == true) {
                alert('捐款人編號 欄位必須為數字！');
                tbxDonor_Id.focus();
                return false;
            }
            return true;
        }

        function openEditDonor(e, GroupItemUid) {

            $('img[title="增加捐款人"]').hide();
            $(e).after('<div id="divInputDonor"><span style="color: blue;" >輸入捐款人編號：</span><input type="text" id="GroupItemDonorID" name="GroupItemDonorID" style="width: 100px" onkeydown="if(event.keyCode==13) getDonor(this,' + GroupItemUid + ');" /><img onclick="saveDonor(this,' + GroupItemUid + ');" src="../images/toolbar_submit.gif" alt="確定" title="確定" border=0 /><img onclick="cancelDonor(this);" src="../images/toolbar_cancel.gif" alt="取消" title="取消" border=0 /></div>');
            return false;
        }

        function getDonor(e, GroupItemUid) {
            alert('尚未完成取得捐款人資料...');
            return false;
        }

        function saveDonor(e, GroupItemUid ) {
            
            var DonorID = $('#GroupItemDonorID').val();
            //儲存群組的捐款人
            $(e).parent().parent().prev().append(', <span style="color: blue;" >' + DonorID + '(ok)</span>');
            $('#divInputDonor').remove();
            $('img[title="增加捐款人"]').show();
            alert(DonorID);
            return false;
        }

        function cancelDonor(e, GroupItemUid) {

            $('img[title="增加捐款人"]').show();
            $('#divInputDonor').remove();
            return false;

        }

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
            捐款人群組代表管理
        </h1>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
            <tr>
                <th align="right" width="15%">群組類別：</th>
                <td width="20%" align="left">
                    <asp:DropDownList ID="ddlGroupClass" runat="server" AutoPostBack="False" >
                    </asp:DropDownList>
                 </td>
                 <td align="right" rowspan="4">
                     <table><tr><td>
                     <asp:Button ID="btnQuery" class="npoButton npoButton_Search" runat="server" Text="查詢" OnClick="btnQuery_Click" OnClientClick="return CheckField();"/>
                         </td><td>
                     <asp:Button ID="btnAdd" class="npoButton npoButton_New" runat="server" Text="新增" OnClick="btnAdd_Click"/>
                         </td><td>
                     <asp:Button ID="btnExcel" class="npoButton npoButton_Excel" runat="server" Text="匯出Excel" OnClick="btnExcel_Click" OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" Width="100px" Height="34px"/>
                         </td></tr></table>
                 </td>
            </tr>
            <tr>
                <th align="right" width="15%">群組代表：</th>
                <td width="30%" align="left">
                    <asp:TextBox runat="server" ID="txtGroupItemName" Width="60mm" CssClass="font9"></asp:TextBox>
                    <asp:TextBox ID="txtGroupItemUid" runat="server" Visible="False"></asp:TextBox>
                 </td>
            </tr>
            <tr>
                 <th width="20%" align="right">備註：</th>
                 <td align="left">
                     <asp:TextBox runat="server" ID="txtSupplement" Width="60mm" CssClass="font9"></asp:TextBox>
                 </td>
            </tr>
            <tr>
                 <th width="20%" align="right">捐款人編號：</th>
                 <td align="left">
                     <asp:TextBox runat="server" ID="txtDonorID" Width="60mm" CssClass="font9"></asp:TextBox>
                 </td>
            </tr>
            <tr>
                <td width="100%" colspan="3">
                    <span style="color:red; font-weight: bold; font-size: medium;">
                        如何設定群組代表：
                        <br />1.群組代表人就像家長或地址一樣，協助在這麼多相同群組類別中，能找到對的群組然後才能加進新成員，代表人是關鍵
                        <br />2.所以要新增代表人，就要先選好群組類別，如果是新類別，就要先到群組類別設定去新增定義後，再回來設代表人
                        <br />3.請按新增鍵依畫面指示即可
                    </span>
                </td>
            </tr>
            <tr>
                <td width="100%" colspan="3">
                    <br/>
                    <asp:Label ID="lblGridList" runat="server"></asp:Label>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
