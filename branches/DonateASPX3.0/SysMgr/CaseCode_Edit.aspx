<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CaseCode_Edit.aspx.cs" Inherits="SysMgr_CaseCode_Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>修改代碼選項</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        function CheckFieldMustFillBasic() {
            var cnt = 0;
            var sName = '';
            var tbxCaseItems = document.getElementById('tbxCaseItems');
            var tbxCaseID = document.getElementById('tbxCaseID');
            if (tbxCaseItems.value == "") {
                alert('代碼選項 欄位不可為空白');
                tbxCaseItems.focus();
                return false;
            }
            cnt = 0;
            sName = tbxCaseItems.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 50) {
                alert('代碼選項 欄位長度超過限制！');
                return false;
            }
            if (tbxCaseID.value == '') {
                alert('排序 欄位不可為空白！');
                tbxCaseID.focus();
                return false;
            }
            if (isNaN(Number(tbxCaseID.value)) == true) {
                alert('排序 欄位必須為數字！');
                tbxCaseID.focus();
                return false;
            }
//            if (window.confirm('您是否要修改？') == true) {
//                return true;
//            }
//            else {
//                return false;
//            }
        }
    </script>
</head>
<body class="body">
    <form id="form1" runat="server">
    <asp:HiddenField ID="HFD_Uid" runat="server" />
        <div align="center">
            <center>
                <table border="0" width="100%" class="table_v">
                    <tr>
                        <td width="100%">
                            <div align="center">
                                <center>
                                    <table width="100%" border="1" cellspacing="0" cellpadding="2">
                                        <tr>
                                            <th align="right" width="15%" height="22">
                                                代碼名稱：
                                            </th>
                                            <td align="left" width="85%">
                                                <asp:TextBox ID="tbxCaseFolder" runat="server" MaxLength="30" ReadOnly="true" CssClass="readonly"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <th align="right" width="15%" height="22">
                                                代碼選項：
                                            </th>
                                            <td align="left" width="85%">
                                                <asp:TextBox ID="tbxCaseItems" Width="240" runat="server" MaxLength="30"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <th align="right" width="25%" height="22">
                                                 排序：
                                            </th>
                                            <td align="left" width="75%">
                                                <asp:TextBox ID="tbxCaseID" Width="24px" CssClass="font9" runat="server"></asp:TextBox>
                                            </td>
                                        </tr>
                                    </table>
                                    <div class="function">
                                        <asp:Button ID="btnEdit" runat="server" Text="修改" OnClick="btnEdit_Click" 
                                             CssClass="npoButton npoButton_Modify"  OnClientClick= "return CheckFieldMustFillBasic() ;"/>
                                        <asp:Button ID="btnCancle" runat="server" Text="離開" OnClientClick="window.close();"
                                             CssClass="npoButton npoButton_Exit" />
                                    </div>
                                </center>
                            </div>
                        </td>
                    </tr>
                </table>
            </center>
        </div>
    </form>
</body>
</html>
