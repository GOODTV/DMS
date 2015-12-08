<%@ Page Language="C#" AutoEventWireup="true" CodeFile="GroupItemMember.aspx.cs" Inherits="DonorGroupMgr_GroupItemMember" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>群組代表成員管理</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <link href="../include/jquery-ui-1.8.18.custom/css/ui-lightness/jquery-ui-1.8.18.custom.css" rel="stylesheet" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery-ui-1.8.18.custom/js/jquery-ui-1.8.18.custom.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript">

        $(document).ready(function () {

            // menu控制
            InitMenu();

            /*
            //點選未顯示的說明文字
            $("#divhelp span").click(function () {
                $("#hfHelp").val('1');
                //說明文字區塊顯示
                ShowHide();
            });

            //點選已顯示的說明文字
            $("#tbhelp span").click(function () {
                $("#hfHelp").val('0');
                //說明文字區塊隱藏
                ShowHide();
            });
            */

            // 取得表格大小
            $(".table_v").width(Math.round(($(window).width() - 50) / 2));
            if ($("#divGroupItem").length > 0) {
                $("#divGroupItem").height($(window).height() - Math.round($("#divGroupItem").position().top) - 10);
            }
            if ($("#divDonor").length > 0) {
                $("#divDonor").height($(window).height() - Math.round($("#divDonor").position().top) - 10);
            }

            //改變視窗大小時，重算表格大小
            $(window).resize(function () {
                var window_width = $(window).width();
                var window_height = $(window).height();
                $(".table_v").width(Math.round((window_width - 50) / 2));
                if ($("#divGroupItem").length > 0) {
                    $("#divGroupItem").height(window_height - Math.round($("#divGroupItem").position().top) - 10);
                }
                if ($("#divDonor").length > 0) {
                    $("#divDonor").height(window_height - Math.round($("#divDonor").position().top) - 10);
                }
            });

            //產出拖拉功能
            DragDrop();

            //捲軸
            if ($('#hfGroupItemScrollTop').val() > 0) {
                //$("#divGroupItem").scrollTop(Number($('#hfGroupItemScrollTop').val()));
                $("#divGroupItem").scrollTop($('#hfGroupItemScrollTop').val());
            }
            if ($('#hfDonorScrollTop').val() > 0) {
                $("#divDonor").scrollTop($('#hfDonorScrollTop').val());
            }

            //說明文字區塊顯示與否
            ShowHide();

            //High light 已拖移成員的文字
            if ($('#hfHighLight').val() != '') {
                var findword1 = $("#hfHighLight").val().split(",")[0];
                var findword2 = $("#hfHighLight").val().split(",")[1];
                var findword = $('#divGroupItem tr[id=' + findword1 + ']').find('span:contains("' + findword2 + '")');
                $('#hfHighLight').val('');
                //閃爍
                $(findword).fadeOut(1000).fadeIn(1000).fadeOut(1000).fadeIn(1000);
                //變色
                $(findword).animate({ 'backgroundColor': '#00FF00' }, 1000).animate({ 'backgroundColor': '#fbf0b3' }, 3000);
            }

        });

        function ShowHide() {

            //儲存Cookie
            setCookie('GroupItemHelp', $("#hfHelp").val());
            //說明文字區塊顯示與否
            if ($("#hfHelp").val() == '1') {
                $("#tbhelp").show();
                $('#tbhelp').offset($('#divhelp').offset());
            }
            else {
                $("#tbhelp").hide();
            }

        }

        //檢查群組代表的成員(捐款人)是否存在
        function checkGroupItemMember(Donor_Id,GroupItemUid) {

            var result = false;
            $.ajax({
                type: 'post',
                url: "../common/ajax.aspx",
                async: false,
                data: "Type=18&DonorID=" + Donor_Id + "&GroupItemUid=" + GroupItemUid,
                success: function (data) {
                    if (data != '') {
                        result = true;
                    }
                },
                error: function () {
                    alert('ajax failed');
                }

            });
            return result;
        }

        //儲存捐款人的群組代表資料
        function saveGroupItemMember(Donor_Id, GroupItemUid) {

            var result = '';
            $.ajax({
                type: 'post',
                url: "../common/ajax.aspx",
                async: false,
                data: "Type=17&DonorID=" + Donor_Id + "&GroupItemUid=" + GroupItemUid,
                success: function (data) {
                    if (data != '') {
                        result = data;
                    }
                },
                error: function () {
                    alert('ajax failed');
                }

            });
            return result;

        }

        //群組代表區塊的查詢前檢查
        function CheckField() {

            var ddlGroupClass = $('#ddlGroupClass').val();
            var txtGroupItemName = $('#txtGroupItemName').val();
            var txtSupplement = $('#txtSupplement').val();
            var txtDonorID = $('#txtDonorID').val();

            //畫面有輸入查詢條件時，清除從捐款人修改頁面過來的參數
            if (ddlGroupClass != '' || txtGroupItemName != '' || txtSupplement != '' || txtDonorID != '') {
                $('#hfGroupItemUid').val('');
            }

            if (isNaN(Number(txtDonorID)) == true) {
                /*
                $("#dialog-message").dialog({
                    modal: true,
                    resizable: false,
                    buttons: {
                        Ok: function () {
                            $(this).dialog("close");
                            $('#txtDonorID').focus();
                        }
                    }
                });
                */
                alert('捐款人編號 欄位必須為數字！');
                $('#txtDonorID').focus();
                return false;
            }
            else {
                $('#hfGroupItemScrollTop').val($("#divGroupItem").scrollTop());
                $('#hfDonorScrollTop').val($("#divDonor").scrollTop());
                $('#btnQuery').css('cursor', 'progress');
                document.body.style.cursor = 'progress';
                return true;
            }
        }

        //捐款人區塊的查詢前檢查
        function CheckFieldMustFillBasic() {

            if (isNaN(Number($('#tbxDonor_Id').val())) == true) {
                alert('捐款人編號 欄位必須為數字！');
                $('#tbxDonor_Id').focus();
                return false;
            }
            else {
                $('#hfGroupItemScrollTop').val($("#divGroupItem").scrollTop());
                $('#hfDonorScrollTop').val($("#divDonor").scrollTop());
                $('#btnDonorQuery').css('cursor', 'progress');
                document.body.style.cursor = 'progress';
                return true;
            }
        }

        //產出拖拉功能
        function DragDrop() {

            //從捐款人拖移到群組代表的功能
            //begin
            $("#lblDonorGridList .table_h tr:not(:first)").draggable({

                cursor: 'move',     //cell,move
                cursorAt: { bottom: 14, right: 0 },     //top: 5, left: 26
                revert: "invalid",
                containment: "#container",
                helper: function (event) {
                    var result = $('<div class="drag_row">' + $(event.target).closest('tr').children().first().text() + '</div>');
                    return result;
                },
                appendTo: 'body'

            }).addClass('Donor_dragd');

            var Donor_Id;
            var Donor_info;
            $("#lblGroupItemGridList .table_h tr:not(:first)").droppable({

                /*tolerance: "touch",
                */
                accept: ".Donor_dragd",
                activeClass: "Donor-dragd-active",
                hoverClass: "Donor-dragd-hover",
                over: function (event, ui) {
                    Donor_Id = $(ui.draggable).children().first().text();
                    Donor_info = $(ui.draggable).find('td:eq(0)').text() + ' , ' + $(ui.draggable).find('td:eq(1)').text();

                },
                drop: function (event, ui) {
                    var GroupItemUid = $(this).attr('id');
                    var GroupItemMember = $(this).find('td:last');
                    var GroupItem = $(this).find('td:eq(1)').text() + ' , ' + $(this).find('td:eq(2)').text();
                    if (confirm('群組類別與代表：' + GroupItem + '\n準備要加入成員：' + Donor_info)) {
                        Donor_info
                        //儲存群組代表的成員(捐款人)
                        if (checkGroupItemMember(Donor_Id, GroupItemUid)) {
                            alert('拖移的捐款人已存在於該群組代表的成員內！');
                        }
                        else {
                            var data = saveGroupItemMember(Donor_Id, GroupItemUid);
                            if (data != '') {
                                $('#hfHighLight').val(GroupItemUid + ',' + Donor_Id);
                                $('#btnQuery').click();
                            }
                        }

                    }
                    else {
                        //alert('false');
                    }
                    return false;
                }

            });
           //end

            //從群組代表內拖出成員(捐款人)的功能
            //begin
            $("#lblGroupItemGridList .table_h tr td:last-child span").draggable({

                cursor: 'move',     //move
                cursorAt: { bottom: 17, left: -4 },     //top: 5, right: 26
                revert: "invalid",
                containment: "#container",
                helper: function (event) {
                    var result = $('<div class="drag_row">' + $(event.target).closest('span').children().first().text() + '</div>');
                    return result;
                },
                appendTo: 'body'

            }).addClass('GroupItem_dragd');

            $("body").droppable({

                accept: '.GroupItem_dragd',
                /*tolerance: "touch",*/
                drop: function (event, ui) {
                    var Donor_Id = $(ui.draggable).children().first().text();
                    var Donor_info = $(ui.draggable).text();
                    var GroupItemUid = $(ui.draggable).closest('tr').attr('id');
                    var GroupItemMember = $(ui.draggable);
                    var GroupItem = $(ui.draggable).closest('tr').find('td:eq(1)').text() + ' , ' + $(ui.draggable).closest('tr').find('td:eq(2)').text();
                    if (confirm('群組類別與代表：' + GroupItem + '\n準備要移除成員：' + Donor_info)) {

                        //刪除(update刪除註記)群組代表的成員(捐款人)
                        $.ajax({
                            type: 'post',
                            url: "../common/ajax.aspx",
                            data: "Type=15&DonorID=" + Donor_Id + "&GroupItemUid=" + GroupItemUid,
                            success: function (data) {
                                if (data == 'Y') {
                                    $('#btnQuery').click();
                                }
                            },
                            error: function () {
                                alert('ajax failed');
                            }

                        });

                    }
                    else {
                        //alert('false');
                    }
                    return false;

                }

            });

            /*
            //資源回收桶
            $("#Recycling").droppable({

                accept: ".GroupItem_dragd",
                over: function (event, ui) {
                    Donor_Id = $(ui.draggable).children().first().text();
                    Donor_info = $(ui.draggable).text();
                },
                drop: function (event, ui) {
                    var GroupItemUid = $(ui.draggable).closest('tr').attr('id');
                    var GroupItemMember = $(ui.draggable);
                    var GroupItem = $(ui.draggable).closest('tr').find('td:eq(1)').text() + ' , ' + $(ui.draggable).closest('tr').find('td:eq(2)').text();
                    if (confirm('群組類別與代表：' + GroupItem + '\n準備要移除成員：' + Donor_info)) {

                        //刪除(update刪除註記)群組代表的成員(捐款人)
                        $.ajax({
                            type: 'post',
                            url: "../common/ajax.aspx",
                            data: "Type=15&DonorID=" + Donor_Id + "&GroupItemUid=" + GroupItemUid,
                            success: function (data) {
                                if (data == 'Y') {
                                    $('#btnQuery').click();
                                }
                            },
                            error: function () {
                                alert('ajax failed');
                            }

                        });

                    }
                    else {
                        //alert('false');
                    }
                    return false;
                }

            });
            */

            //end

        }

        //儲存Cookie
        function setCookie(key, value) {
            var expires = new Date();
            expires.setTime(expires.getTime() + (7 * 24 * 60 * 60 * 1000));
            document.cookie = key + '=' + value + ';expires=' + expires.toUTCString();
        }

        //取得Cookie
        function getCookie(key) {
            var keyValue = document.cookie.match('(^|;) ?' + key + '=([^;]*)(;|$)');
            return keyValue ? keyValue[2] : null;
        }

    </script>
    <style type="text/css">

        .Donor-dragd-active {
           border: 1px solid #00FF00;
            /*cursor: copy;*/
        }

        .Donor-dragd-hover {
           border: 1px solid #00FF00;
            /*cursor: copy;*/
        }

        HTML {
            overflow: hidden;
        }

        .body {
            color: black;
            scrollbar-face-color: whitesmoke;
        }

        .table_h td {
            word-break: break-all;
            height: 25px;
            color: black;
        }

        .table_h td:first-child {
            white-space: nowrap;
        }

        .table_v td {
            line-height: none;
        }

        #lblDonorGridList .table_h td:hover {
            cursor: pointer;
        }

        .drag_row {
            font-size: 14px;
            padding: 5px 10px;
            background-color: #b6ff00;
            color: #0026ff;
        }

        .GroupItem_dragd {
            background-color: #fbf0b3;
            line-height: 20px;
            margin: 3px;
            display: inline-block;
        }
        /*
            .GroupItem_dragd font {
                color: red;
            }
            */
        .GroupItem_dragd:hover {
            cursor: pointer;
        }

        /*
        #divhelp span {
            color: red;
            font-weight: bold;
            font-size: medium;
            cursor: help;
            padding-left: 2px;
        }
        */

        #tbhelp span {
            color: blue;
            font-weight: bold;
            font-size: medium;
            cursor: pointer;
            /*font-size: 12px;*/
        }

        #tbhelp table {
            color: red;
            font-weight: bold;
            font-size: medium;
            /*font-size: 12px;*/
        }

    </style>
</head>
<body class="body" onkeydown="if(event.keyCode==13) return false;">
    <form id="Form1" name="form" runat="server">
        <asp:HiddenField ID="hfGroupItemUid" runat="server" />
        <asp:HiddenField ID="hfGroupItemScrollTop" runat="server" />
        <asp:HiddenField ID="hfDonorScrollTop" runat="server" />
        <asp:HiddenField ID="hfHelp" runat="server" Value="0"/>
        <asp:HiddenField ID="hfHighLight" runat="server"/>
        <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
        <div id="menucontrol">
            <a href="#">
                <img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
            <a href="#">
                <img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        </div>
        <div id="container">
            <h1 style="padding-bottom: 0px;">
                <img src="../images/h1_arr.gif" alt="" />
                群組代表成員管理
            </h1>
            <table style="width: 100%;">
                <tr>
                    <td style="vertical-align: top;">
                            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
                                <tr>
                                    <th align="right">群組類別：</th>
                                    <td align="left" width="35%">
                                        <asp:DropDownList ID="ddlGroupClass" runat="server" AutoPostBack="False">
                                        </asp:DropDownList>
                                    </td>
                                    <th align="right">群組代表：</th>
                                    <td align="left" width="35%">
                                        <asp:TextBox runat="server" ID="txtGroupItemName" Width="95%" CssClass="font9"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <th align="right">備註：</th>
                                    <td align="left">
                                        <asp:TextBox runat="server" ID="txtSupplement" Width="95%" CssClass="font9"></asp:TextBox>
                                    </td>
                                    <th align="right">捐款人編號：</th>
                                    <td align="left">
                                        <asp:TextBox runat="server" ID="txtDonorID" Width="95%" CssClass="font9"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" colspan="4" valign="middle">
                                        <asp:Button ID="btnQuery" class="npoButton npoButton_Search" runat="server" Text="查詢" OnClick="btnQuery_Click" OnClientClick="return CheckField();" />
                                    </td>
                                </tr>
                                <!--<tr>
                                    <td colspan="4">
                                        <div id="divhelp"><span>如何新增或移出成員？</span></div>
                                    </td>
                                </tr>-->
                                <tr>
                                    <td width="100%" colspan="4">
                                        <asp:Label ID="lblGroupItemGridList" runat="server"></asp:Label>
                                    </td>
                                </tr>
                            </table>
                    </td>
                    <td style="vertical-align: top;">
                            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v">
                                <tr>
                                    <th align="right" colspan="1">捐款人：
                                    </th>
                                    <td align="left" colspan="1" width="35%">
                                        <asp:TextBox runat="server" ID="tbxDonor_Name" Width="95%" CssClass="font9"></asp:TextBox>
                                    </td>
                                    <th align="right" colspan="1">捐款人編號：
                                    </th>
                                    <td align="left" colspan="1" width="35%">
                                        <asp:TextBox runat="server" ID="tbxDonor_Id" Width="95%" CssClass="font9"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <th align="right" colspan="1">收據抬頭：
                                    </th>
                                    <td align="left" colspan="1">
                                        <asp:TextBox runat="server" ID="tbxInvoice_Title" Width="95%" CssClass="font9"></asp:TextBox>
                                    </td>
                                    <th align="right" colspan="1">連絡電話：
                                    </th>
                                    <td align="left" colspan="1">
                                        <asp:TextBox runat="server" ID="tbxTel_Office" Width="95%" CssClass="font9"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <th align="right" colspan="1">地址：
                                    </th>
                                    <td align="left" colspan="3">
                                        <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
                                            <ContentTemplate>
                                                <asp:CheckBox runat="server" ID="cbxIsAbroad" Text="海外地址" CssClass="font9"
                                                    AutoPostBack="True" OnCheckedChanged="cbxIsAbroad_CheckedChanged"></asp:CheckBox>
                                                <asp:Panel runat="server" ID="PanelAbroad">
                                                    <asp:TextBox runat="server" ID="tbxIsAbroad_Address" CssClass="font9" Width="500px" MaxLength="200"></asp:TextBox>
                                                </asp:Panel>
                                                <asp:Panel runat="server" ID="PanelLocal">
                                                    <asp:TextBox runat="server" ID="tbxZipCode" CssClass="font9" Width="60px"></asp:TextBox>
                                                    <asp:DropDownList runat="server" ID="ddlCity" CssClass="font9"
                                                        AutoPostBack="True" OnSelectedIndexChanged="ddlCity_SelectedIndexChanged">
                                                    </asp:DropDownList>
                                                    <asp:DropDownList runat="server" ID="ddlArea" CssClass="font9"
                                                        AutoPostBack="True" OnSelectedIndexChanged="ddlArea_SelectedIndexChanged">
                                                    </asp:DropDownList>
                                                    <asp:TextBox runat="server" ID="tbxStreet" CssClass="font9" Width="100px"></asp:TextBox>大道/路/街/部落
                    <asp:DropDownList runat="server" ID="ddlSection" CssClass="font9"></asp:DropDownList>段
                    <asp:TextBox runat="server" ID="tbxLane" CssClass="font9" Width="40px"></asp:TextBox>巷
                    <asp:TextBox runat="server" ID="tbxAlley" CssClass="font9" Width="40px"></asp:TextBox>
                                                    <asp:DropDownList runat="server" ID="ddlAlley" CssClass="font9"></asp:DropDownList>
                                                    <asp:TextBox runat="server" ID="tbxNo1" CssClass="font9" Width="55px"></asp:TextBox>號之
                    <asp:TextBox runat="server" ID="tbxNo2" CssClass="font9" Width="40px"></asp:TextBox>
                    <asp:TextBox runat="server" ID="tbxFloor1" CssClass="font9" Width="40px"></asp:TextBox>樓之
                    <asp:TextBox runat="server" ID="tbxFloor2" CssClass="font9" Width="40px"></asp:TextBox>&nbsp;
                    <asp:TextBox runat="server" ID="tbxRoom" CssClass="font9" Width="40px"></asp:TextBox>室
                                                </asp:Panel>
                                            </ContentTemplate>
                                        </asp:UpdatePanel>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                    </td>
                                    <td align="right">
                                        <asp:Button ID="btnDonorQuery" CssClass="npoButton npoButton_Search" runat="server" Width="20mm"
                                            Text="查詢" OnClientClick="return CheckFieldMustFillBasic();" OnClick="btnDonorQuery_Click" />
                                    </td>
                                </tr>
                                <tr>
                                    <td width="100%" colspan="4">
                                        <asp:Label ID="lblDonorGridList" runat="server"></asp:Label>
                                    </td>
                                </tr>
                            </table>
                    </td>
                </tr>
            </table>
        </div>

        <div id="dialog-message" title="群組代表成員管理" style="display: none">
            <p>
                <span class="ui-icon"></span>
                捐款人編號 欄位必須為數字！
            </p>
        </div>

        <div id="tbhelp" style="background-color: #FFFF00; padding: 2px; margin: 0px; position: absolute; display: none; left: 28px; top: 130px;">
            <span>新增或移出成員說明：</span>
            <table>
                <tr>
                    <td valign="top">1.</td>
                    <td>先在左邊查詢找出要管理的群組，查詢結果會列出所屬成員。</td>
                </tr>
                <tr>
                    <td valign="top">2.</td>
                    <td>再來在右邊查詢找出欲加入左邊關聯群組的捐款人。(移除成員不需查詢)</td>
                </tr>
                <tr>
                    <td valign="top">3.</td>
                    <td>新增成員：用滑鼠點選並持續按在右邊的任一位捐款人，然後拖拉到左邊要加入的群組位置，放開就會自動加到成員列表中。</td>
                </tr>
                <tr>
                    <td valign="top">4.</td>
                    <td>移除成員：使用滑鼠點選並持續按在左邊群組代表成員的捐款人不放，再拖移到任何地方，即可刪除該成員。</td>
                </tr>
            </table>
        </div>

    </form>
</body>
</html>
