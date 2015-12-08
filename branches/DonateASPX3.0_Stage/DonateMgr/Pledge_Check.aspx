<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Pledge_Check.aspx.cs" Inherits="DonateMgr_Pledge_Check" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>確認</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        function ReturnOpener() {
            var DonateDate = $("#HFD_DonateDate").val();
            opener.document.getElementById('HFD_LastTransfer_Check').value = '1';
            opener.document.getElementById('tbxDonateDate').value = DonateDate;
            window.close();
        }
        function Confirm_Data() {
            var Total_Row = "";
            var Pledge_Id = "";
            jQuery("input[type=checkbox]:checked").each(function () { //判斷checkbox勾選了幾個
                Pledge_Id += "1";
            });
            Total_Row = Pledge_Id.length;
            if (Total_Row != 0) {
                if (confirm('您是否確定要將授權資料轉入捐款資料？\n\n※請注意轉入過程中請勿『關閉視窗』') == false) {
                    return false;
                }
                else {
                    //2014/6/4 新增by Ian_Kao 確認送出後hide按鈕
                    $("#btnImport").hide();
                    $("#lblWarming").html('正在確認中，請稍等.... ');
                    return true;
                }
            } 
            else if (Total_Row == 0) {
                alert('查無相關授權資料無法轉入！');
                return false;
            }
        }
    </script>
    <style type="text/css">
        .style1
        {
            height: 35px;
        }
        .npoButton_Export
        {}
    </style>
</head>
<body class="body">
    <form id="Form1" runat="server">
    <div id="container">
    <asp:HiddenField runat="server" ID="HFD_Pledge_Id"/>
    <asp:HiddenField runat="server" ID="HFD_DonateDate"/>
    <asp:HiddenField runat="server" ID="HFD_Total_Row"/>
    <asp:HiddenField runat="server" ID="HFD_Total_Amt"/>
    <asp:HiddenField runat="server" ID="HFD_PledgeBatchFileName"/>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
          <tr>
                    <td align="center" width="100%" colspan="8">
                         <asp:Label ID="lblCount_Amount" runat="server" ></asp:Label>
                    </td>
          </tr>
          <tr>
                    <td align="center" width="100%" colspan="8">
                         <asp:Label ID="lblGridList" runat="server" ></asp:Label>
                    </td>
          </tr>
                    
          <tr>
            <td align="center" colspan="1" >
                <asp:Button ID="btnImport" CssClass="npoButton npoButton_Export" runat="server"  Width="25mm" 
                    Text="確認送出" OnClientClick="return Confirm_Data()" OnClick="btnInput_Click" />
            </td>
          </tr>
          <tr>
                    <td align="center" width="100%" colspan="8">
                         <asp:Label ID="lblWarming" runat="server" ></asp:Label>
                    </td>
          </tr>
    </table>             
    </div>
    </form>
</body>
</html>

