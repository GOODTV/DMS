<%@ Page Language="C#" AutoEventWireup="true" CodeFile="LinkedList_Add.aspx.cs" Inherits="ContributeMgr_LinkedList_Add" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>新增物品類別</title>
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
            var tbxLinked_Name = document.getElementById('tbxLinked_Name');
            var tbxLinked_Seq = document.getElementById('tbxLinked_Seq');
            if (tbxLinked_Name.value == "") {
                alert('物品類別 欄位不可為空白');
                tbxLinked_Name.focus();
                return false;
            }
            cnt = 0;
            sName = tbxLinked_Name.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 50) {
                alert('物品類別 欄位長度超過限制！');
                return false;
            }
            if (tbxLinked_Seq.value == '') {
                alert('排序 欄位不可為空白！');
                tbxLinked_Seq.focus();
                return false;
            }
            if (isNaN(Number(tbxLinked_Seq.value)) == true) {
                alert('排序 欄位必須為數字！');
                tbxLinked_Seq.focus();
                return false;
            }
            return true;
        }
    </script>
</head>
<body class="body">
    <form id="form1" runat="server">
        <div align="center">
            <center>
                <table border="0" width="100%" class="table_v">
                    <tr>
                        <td width="100%">
                            <div align="center">
                                <center>
                                    <table width="100%" border="1" cellspacing="0" cellpadding="2">
                                        <tr>
                                            <th align="right" width="20%" height="22">
                                                物品類別：
                                            </th>
                                            <td align="left" width="80%">
                                                <asp:TextBox ID="tbxLinked_Name" Width="240" runat="server" MaxLength="30"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <th align="right" width="15%" height="22">
                                                 排序：
                                            </th>
                                            <td align="left" width="85%">
                                                <asp:TextBox ID="tbxLinked_Seq" Width="24px" CssClass="font9" runat="server"></asp:TextBox>
                                            </td>
                                        </tr>
                                    </table>
                                    <div class="function">
                                        <asp:Button ID="btnSave" runat="server" Text="存檔" OnClick="btnSave_Click" 
                                             CssClass="npoButton npoButton_Save" />
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