<%@ Page Language="C#" AutoEventWireup="true" CodeFile="OnlineQuestionnaire.aspx.cs" Inherits="Online_OnlineQuestionnaire" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>奉獻問卷調查</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            //$('input:checkbox[id*=CHK_DonateMotive]').bind("click", function (e) {
            //    var ret = '';
            //    $('input:checkbox:checked[name="DonateMotive"]').each(function (i) { ret += this.value + ','; });
            //    $('#HFD_DonateMotive').val(ret);
            //});
            //$('input:checkbox[id*=ChildWatchMode2]').bind("click", function (e) {
            //    if ($('#ChildWatchMode2_1').prop('checked') == true || $('#ChildWatchMode2_2').prop('checked') == true || $('#ChildWatchMode2_3').prop('checked') == true) {
            //        $("#CHK_WatchMode2").attr("checked", "checked");
            //    }
            //    else {
            //        $("#CHK_WatchMode2").removeAttr("checked");
            //    }
            //});
            //$('input:checkbox[id*=ChildWatchMode3]').bind("click", function (e) {
            //    if ($('#ChildWatchMode3_1').prop('checked') == true || $('#ChildWatchMode3_2').prop('checked') == true) {
            //        $("#CHK_WatchMode3").attr("checked", "checked");
            //    }
            //    else {
            //        $("#CHK_WatchMode3").removeAttr("checked");
            //    }
            //});
            //$('input:checkbox[id*=ChildWatchMode4]').bind("click", function (e) {
            //    if ($('#ChildWatchMode4_1').prop('checked') == true || $('#ChildWatchMode4_2').prop('checked') == true) {
            //        $("#CHK_WatchMode4").attr("checked", "checked");
            //    }
            //    else {
            //        $("#CHK_WatchMode4").removeAttr("checked");
            //    }
            //});
            //$('#Journalism').bind('change', function (e) {
            //    if ($('#Journalism').prop('checked') == true) {
            //        $("#CHK_WatchMode5").attr("checked", "checked");
            //    }
            //    else {
            //        $("#CHK_WatchMode5").removeAttr("checked");
            //    }
            //});
            //$('input:checkbox[id*=CHK_WatchMode]').bind("click", function (e) {
            //    var ret = '';
            //    $('input:checkbox:checked[name="WatchMode"]').each(function (i) { ret += this.value + ','; });
            //    $('#HFD_WatchMode').val(ret);
            //});
        });
     </script>
    <style type="text/css">
        .table_v2 tr {
            font-size: 16px;
        }

        .table_v2 th {
            line-height: 30px;
        }

        .table_v2 td {
            line-height: 30px;
        }

        input {
            font-size: 16px;
        }

        select {
            font-size: 16px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:HiddenField ID="HFD_OrderId" runat="server" value=""/>
    <asp:HiddenField ID="HFD_UID" runat="server" value=""/>
    <asp:HiddenField ID="HFD_Device" runat="server" value=""/>
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
    </div>
    </form>
</body>
</html>
