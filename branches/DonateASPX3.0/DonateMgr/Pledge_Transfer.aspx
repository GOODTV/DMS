<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Pledge_Transfer.aspx.cs" Inherits="DonateMgr_Pledge_Transfer" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>批次固定轉帳作業</title>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">

        // menu控制
        $(document).ready(function () {

            InitMenu();

            // 2014/4/11 增加全選的功能
            $("#checkboxAll").click(function () {

                if ($("#checkboxAll").prop("checked")) {
                    $("input[id=checkbox]").prop("checked", true);
                }
                else {
                    $("input[id=checkbox]").prop("checked", false);
                }
            });

            initCalendar();
            Window_OnLoad();
        });

        //window.onload = initCalendar;

        function initCalendar() {
            Calendar.setup({
                inputField: "tbxDonateDate",   // id of the input field
                button: "imgDonateDate"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxNext_DonateDateB",   // id of the input field
                button: "imgNext_DonateDateB"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxNext_DonateDateE",   // id of the input field
                button: "imgNext_DonateDateE"     // 與觸發動作的物件ID相同
            });
        }

        function Print() {
            var tbxDonateDate = document.getElementById('tbxDonateDate');
            var ddlFormat = document.getElementById('ddlFormat');
            var strRet = "";
            if (tbxDonateDate.value == "") {
                strRet += "扣款日期 ";
            }
            if (ddlFormat.value == "") {
                strRet += "TXT格式 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet);
                return false;
            }
            if (confirm('您是否確定要將TXT授權資料匯出？') == false) {
                return false;
            }
            else {
                return true;
            }
        }

        function Confirm_Data(LastTransfer_Check) {

            var Total_Row = "";
            var Pledge_Id = "";

            $("input[id=checkbox]:checked").each(function () { //判斷checkbox勾選了幾個
                Pledge_Id += "1";
            });

            Total_Row = Pledge_Id.length;
            if (Total_Row != 0) {
            
                if (LastTransfer_Check == '1') {
                    if (confirm('本月份已執行過授權資料轉入捐款資料，\n\n您是否確定要『再次』執行？') == false) {
                        return false;
                    }
                    else {
                        window.open('Pledge_Check.aspx?', 'NewWindows',
                        'status=no,toolbar=no,location=no,menubar=no,width=800,height=600);');
                    }
                }
                else {
                        var DonateDate_D = document.getElementById('tbxDonateDate');
                        window.open('Pledge_Check.aspx?DonateDate=' + DonateDate_D.value + '&LastTransfer_Check=HFD_LastTransfer_Check' , 'NewWindows',
                        'status=no,toolbar=no,location=no,menubar=no,width=800,height=600);');
                }
            
            } else if (Total_Row == 0) {
                alert('查無相關授權資料無法轉入！');
                return false;
            }
            
        }

        function Confirm_Data1() {

            var Total_Row = "";
            var Pledge_Id = "";
            //20140518 修改by Ian
            //var LastTransfer_Check = document.getElementById('LastTransfer_Check');
            var HFD_LastTransfer_Check = document.getElementById('HFD_LastTransfer_Check');

            $("input[id=checkbox]:checked").each(function () { //判斷checkbox勾選了幾個
                Pledge_Id += "1";
            });

            Total_Row = Pledge_Id.length;
            if (Total_Row != 0) {

                //20140518 修改by Ian
                if (HFD_LastTransfer_Check.value == '1') {
                    if (confirm('本月份已執行過授權資料轉入捐款資料，\n\n您是否確定要『再次』執行？') == false) {
                        return false;
                    }
                }
                else
                return true;
            }
            else if (Total_Row == 0) {
                alert('查無相關授權資料無法轉入！');
                return false;
            }
        }

        //20140122新增
        function Window_OnLoad() {
            if (document.getElementById("HFD_Query_Flag").value == ''){
                $("#btnImport").attr("disabled", "disabled");
                $("#btnImport_CheckFalse").attr("disabled", "disabled");
                $("#btnImportCathayTXT").attr("disabled", "disabled");
                $("#btnExportErrorExcel").attr("disabled", "disabled");
            }
            else{
                if (document.getElementById("HFD_Flag").value == 'True') {
                    document.getElementById('btnImport').removeAttribute("disabled");
                    document.getElementById('btnImport_CheckFalse').removeAttribute("disabled");
                    document.getElementById('btnImportCathayTXT').removeAttribute("disabled");
                    document.getElementById('btnExportErrorExcel').removeAttribute("disabled");
                }
                else {
                    $("#btnImport").attr("disabled", "disabled");
                    $("#btnImport_CheckFalse").attr("disabled", "disabled");
                    $("#btnImportCathayTXT").attr("disabled", "disabled");
                    $("#btnExportErrorExcel").attr("disabled", "disabled");
                }
            }
        }

        function showFileImport() {
            $('#HFD_Pledge_Import').val('bot');
            $("#reg").fadeTo("fast", 0.6);
            $('#openFileImport').show();
        }

        //增加國泰世華
        function showFileImport2() {
            $('#HFD_Pledge_Import').val('cathay');
            $("#reg").fadeTo("fast", 0.6);
            $('#openFileImport').show();
        }

        function hideFileImport() {
            $("#reg").hide()
            $('#openFileImport').hide();
        }

        // 2014/9/15 修改匯入檔案功能 by Samson Hsu
        function FileImportCheck() {

            //檢查檔案欄位名稱
            var filename = document.getElementById("FileUpload");
            if (filename.value == "") {
                alert("檔案名稱不能有空白！");
                return false;
            }

            //檢查副檔名
            var validExtensions = ["txt", "01r", "02r", "03r", "04r", "05r", "06r", "07r", "08r", "09r"];
            var ext = filename.value.substring(filename.value.lastIndexOf('.') + 1).toLowerCase();
            var boolExtOk = false;
            var boolFileIsExist = false;

            for (var i = 0; i < validExtensions.length; i++) {
                if (ext == validExtensions[i]) {

                    //副檔名正確
                    boolExtOk = true;

                    //檢核匯入回覆檔檔名是否存在
                    var strFile = filename.value.substring(filename.value.lastIndexOf('\\') + 1);

                    $.ajax({
                        type: 'post',
                        url: "../common/ajax.aspx",
                        data: 'Type=10' + '&file=' + strFile,
                        async: false,
                        success: function (result) {
                            if (result == "Y") {
                                boolFileIsExist = true;
                            }
                            else {
                                boolFileIsExist = false;
                            }
                        },
                        error: function () { alert('ajax failed'); }
                    })

                }
            }
            if (boolExtOk) {
                if (boolFileIsExist) {
                    alert("匯入回覆檔的檔名已存在，不允許重複上傳！");
                    return false;
                }
                else {
                    document.body.style.cursor = 'wait';
                    return true;
                }
            }
            else {
                alert("此類副檔名不允許上傳！");
                return false;
            }

        }

        function btnImport_Check() {
            if ($('#HFD_Pledge_Import').val() == 'botok') {
                document.location.href = "PledgeReturnError_Print_Excel.aspx";
            }
            else {
                alert("請先匯入台銀回覆檔！");
            }
        }

        function btnImport_Check2() {
            if ($('#HFD_Pledge_Import').val() == 'cathayok') {
                document.location.href = "PledgeReturnError_Print_Excel2.aspx";
            }
            else {
                alert("請先匯入國泰世華回覆檔！");
            }
        }

    </script>
    <style type="text/css">
        .style1
        {
            height: 35px;
        }
    </style>
</head>
<body class="body" onkeydown="if(event.keyCode==13) return false;">
    <form id="Form1" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <asp:HiddenField runat="server" ID="HFD_Dept_Id" />
    <asp:HiddenField runat="server" ID="HFD_LastTransfer_Check" />
    <asp:HiddenField runat="server" ID="HFD_Pledge_Import" />
    <asp:HiddenField runat="server" ID="HFD_Flag" />
    <asp:HiddenField runat="server" ID="HFD_Query_Flag"/>
    <asp:Literal ID="GridList" runat="server" Visible="False"></asp:Literal>
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        批次固定轉帳作業 
    </h1>
    <div id="reg" style="position: fixed; z-index: 100; top: 0px;
        left: 0px; height: 100%; width: 100%; background: #000; display: none;">
    </div>
    <div id="openFileImport" style="background: #ffffff; border-radius: 5px 5px 5px 5px; color: blue;
        font-size: large; display: none; padding-bottom: 2px; width: 650px; height: 150px;
        z-index: 11001; left: 30%; position: fixed; text-align: center; top: 200px;">
        <br />
        <table style="width: 500px;" align="center" cellpadding="2" cellspacing="4">
            <tr>
                <th>信用卡批次授權(一般)回覆檔匯入：
                </th>
            </tr>
            <tr>
                <td>
                    <asp:FileUpload ID="FileUpload" runat="server" Font-Size="18px" Width="600px" />
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Button ID="btnFileImport" CssClass="npoButton npoButton_Export" runat="server" Width="20mm"
                        Text="匯入" OnClientClick="return FileImportCheck();"  OnClick="btnFileImport_Click" />
                    &nbsp;<asp:Button ID="Button2" class="npoButton npoButton_Cancel" runat="server" Width="20mm" 
                        Text="取消" OnClientClick="hideFileImport(); return false;" />
                </td>
            </tr>
        </table>
    </div>
    <table>
        <tr>
        <td valign="top">
            <table width="830" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
                <tr>
                    <th align="right"colspan="1" >
                        機構：
                    </th>
                    <td align="left" colspan="1">
                        <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
                    </td>
                    <th align="right" colspan="1">
                        捐款人：
                    </th>
                    <td align="left" colspan="1">
                        <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9"></asp:TextBox>
                    </td>
                    <th align="right"colspan="1" >
                        TXT格式：
                    </th>
                    <td align="left" colspan="1">
                        <asp:dropdownlist runat="server" ID="ddlFormat" CssClass="font9"></asp:dropdownlist>
                    </td>
                </tr>
                <tr>
                    <th align="right" colspan="1">
                        扣款日期：
                    </th>
                    <td align="left" colspan="1">
                        <asp:TextBox ID="tbxDonateDate" runat="server" 
                            onchange="CheckDateFormat(this, '扣款日期');" Width="90px"></asp:TextBox>
                        <img id="imgDonateDate" alt="" src="../images/date.gif" /> </td>
                    <th align="right" colspan="1">
                        轉帳週期：
                    </th>
                    <td align="left" colspan="1">
                        <asp:dropdownlist runat="server" ID="ddlDonate_Period" CssClass="font9"></asp:dropdownlist>
                    </td> 
                    <th align="right" colspan="1">
                        授權方式：
                    </th>
                    <td align="left" colspan="1">
                        <asp:dropdownlist runat="server" ID="ddlDonate_Payment" CssClass="font9"></asp:dropdownlist>
                    </td> 
                </tr>
                <tr>
                    <th align="right" colspan="1">
                        下次扣款日：
                    </th>
                    <td colspan="5"><asp:TextBox ID="tbxNext_DonateDateB" runat="server" 
                            onchange="CheckDateFormat(this, '扣款日期');" Width="90px"></asp:TextBox>
                        <img id="imgNext_DonateDateB" alt="" src="../images/date.gif" />～<asp:TextBox ID="tbxNext_DonateDateE" runat="server" 
                            onchange="CheckDateFormat(this, '扣款日期');" Width="90px"></asp:TextBox>
                        <img id="imgNext_DonateDateE" alt="" src="../images/date.gif" /><font color="blue"> [下次扣款日]條件可協助迅速找到特別日扣款授權資料。</font></td>
                </tr>
                <tr>
                    <td align="left" colspan="3" class="style1" Width="500px">
                        本月應執行<br />
                        帳戶轉帳
                        <asp:Label ID="lblAccount_Count" runat="server" Text=""></asp:Label>
                        筆  金額
                        <asp:Label ID="lblAccount_Amount" runat="server" Text=""></asp:Label>
                        元 ╱ 信用卡扣款
                        <asp:Label ID="lblCard_Count" runat="server" Text=""></asp:Label>
                        筆  金額
                        <asp:Label ID="lblCard_Amount" runat="server" Text=""></asp:Label>
                        元╱ ACH轉帳
                        <asp:Label ID="lblACH_Count" runat="server" Text=""></asp:Label>
                        筆  金額
                        <asp:Label ID="lblACH_Amount" runat="server" Text=""></asp:Label>
                        元<br />
                        <font color="blue">
                            <asp:Label ID="lblPledgeBatchFileTile" runat="server"></asp:Label>
                            <asp:Label ID="lblPledgeBatchFileName" runat="server"></asp:Label></font>
                    </td>
                    <td align="right" colspan="3" class="style1">
                        <asp:Button ID="btnQuery" CssClass="npoButton npoButton_Search" runat="server"  Width="17mm"
                            Text="查詢" onclick="btnQuery_Click" />
                        <asp:Button ID="btnPrint" CssClass="npoButton npoButton_Txt" runat="server"  Width="25mm"
                            Text="1.TXT匯出" OnClientClick="return Print();" onclick="btnPrint_Click"/>
                        <asp:Button ID="btnInput" runat="server"  Width="30mm"
                            Text="2.授權轉捐款" CssClass=" npoButton npoButton_Export" 
                           OnClientClick="return Confirm_Data1()" onclick="btnInput_Click"/>
                           <!-- Text="2.授權轉捐款" CssClass=" npoButton npoButton_Export" onclick="btnInput_Click"/>
                            <%--OnClientClick="return Confirm_Data()" onclick="btnInput_Click"/>--%> -->
                    </td>
                </tr>
                <tr>
                    <td align="right" colspan="6" class="style1">
                        <asp:Button ID="btnImport" CssClass="npoButton npoButton_Export" runat="server" Width="130px"
                            Text="匯入台銀回覆檔" OnClientClick="showFileImport(); return false;" />
                        <asp:Button ID="btnImport_CheckFalse" CssClass="npoButton npoButton_Excel" runat="server" Width="220px"
                            Text="匯出台銀回覆檔授權失敗Excel檔" OnClientClick="btnImport_Check(); return false;" />
                        <asp:Button ID="btnImportCathayTXT" CssClass="npoButton npoButton_Export" runat="server" Width="150px"
                            Text="匯入國泰世華回覆檔" OnClientClick="showFileImport2(); return false;" />
                        <asp:Button ID="btnExportErrorExcel" CssClass="npoButton npoButton_Excel" runat="server" Width="240px"
                            Text="匯出國泰世華回覆檔授權失敗Excel檔" OnClientClick="btnImport_Check2(); return false;" />
                    </td>
                </tr>
                <tr>
                    <td colspan="6">
                        <table>
                            <tr>
                                <td width="100">信用卡有效月年：</td>
                                <td width="20" height="15" bgcolor='#66FF99'></td>
                                <td width="40">2個月</td>
                                <td width="20" height="15" bgcolor='#FFFF99'></td>
                                <td width="40">1個月</td>
                                <td width="20" height="15" bgcolor='#FFCCFF'></td>
                                <td width="35">本月</td>
                                <td width="20" height="15" bgcolor='#FF6666'></td>
                                <td width="35">逾期</td>
                            </tr>
                        </table>
                    </td>
                    </tr>
                <tr >
                    <td align="center" width="100%" colspan="8">

                         <asp:Panel ID="Panel1" runat="server">
                         <asp:Label ID="lblGridList" runat="server" ></asp:Label>
                         </asp:Panel>
                    </td>
                </tr>
            </table>
        </td>
        <td valign="top">
           		    <table border="0" cellspacing="0" cellpadding="0" align="left">
           			    <tr style="height: 35px">
    							    <td style="background-color:#EEEEE3">
    								    <b><font size="4" color="brown">&nbsp;::信用卡批次授權(一般)說明-台銀及國泰世華::</font></b>
    							    </td>
    						    </tr>
    						    <tr>
    							    <td style="background-color:#EEEEE3">
    								    <font size="3" color="darkmagenta">
    									    &nbsp;&nbsp;&nbsp;1.1查詢預進行批次授權資料</font><br/>
    									    <font size="3">
    										    &nbsp;&nbsp;&nbsp;&nbsp;☆扣款日期：預計執行授權之日期<br/>
    										    <!--&nbsp;&nbsp;&nbsp;&nbsp;☆授權方式：信用卡授權書(一般)<br/>-->
    										    &nbsp;&nbsp;&nbsp;&nbsp;☆點選按鈕</font><font size="3" color="blue">【查詢】</font>
    									    <p/> 								      								
    							    </td>
    						    </tr>
    						    <tr>
    							    <td style="background-color:#EEEEE3">
    								    <font size="3" color="darkmagenta">
    									    &nbsp;&nbsp;&nbsp;1.2匯出批次授權文字檔</font><br>
    									    <font size="3">
    										    &nbsp;&nbsp;&nbsp;&nbsp;☆勾選預計進行批次授權資料<br/>
    										    &nbsp;&nbsp;&nbsp;&nbsp;☆TXT格式：台灣銀行或國泰世華<br/>
    										    &nbsp;&nbsp;&nbsp;&nbsp;☆點選按鈕</font><font size="3" color="green">【1.TXT匯出】</font>
    								    <p/>
    							    </td>
    						    </tr>
    						    <tr>
    							    <td style="background-color:#EEEEE3">
    								    <font size="3" color="darkmagenta">
    								    &nbsp;&nbsp;&nbsp;1.3將匯出文字檔壓縮加密後，寄送給台銀或透過國泰世華提供之後台界面上傳</font><br>
    								    <font size="3" color="red">
    								    &nbsp;&nbsp;&nbsp;※文字檔附檔名不得為TXT</font>
    								    <p />
    							    </td>
    						    </tr>
    						    <tr align="center">
    							    <td style="background-color:#EEEEE3" >
    								    <font size="2" color="chocolate">
    								    ===========待收到台銀或國泰世華回覆檔後===========<p /></font>
    							    </td>
    					      </tr>
    						    <tr>
    							    <td style="background-color:#EEEEE3">
    									    <font size="3" color="darkmagenta">
    									    &nbsp;&nbsp;&nbsp;1.4需重新以同樣(1.1)條件執行查詢</font><br/>
    									    <font size="3">
    										    &nbsp;&nbsp;&nbsp;&nbsp;☆輸入相同(1.1)查詢條件資料<br/>
    										    &nbsp;&nbsp;&nbsp;&nbsp;☆點選按鈕</font><font size="3" color="blue">【查詢】</font>
    									    <p/>
    							    </td>
    						    </tr>
    						    <tr>
    							    <td style="background-color:#EEEEE3">
    								    <font size="3" color="darkmagenta">
    								    &nbsp;&nbsp;&nbsp;1.5匯入「回覆檔」</font><br/>
    								    <font size="3">
    									    &nbsp;&nbsp;&nbsp;&nbsp;☆點選按鈕</font><font size="3" color="red">【匯入台銀回覆檔】或【匯入國泰世華回覆檔】</font><font size="3">進行匯入</font>    									
    								    <p />
    							    </td>
    						    </tr>
    						    <tr>
    							    <td style="background-color:#EEEEE3">
    								    <font size="3" color="darkmagenta">
    								    &nbsp;&nbsp;&nbsp;頁面將呈現在本次匯入回覆檔的二層結果:</font><br/>
    								    <font size="3">
    									    &nbsp;&nbsp;&nbsp;&nbsp;☆第一層:本次匯入授權成功(系統會自動勾選已授權成功資料)</font><br />
    									    <font size="3">&nbsp;&nbsp;&nbsp;&nbsp;☆第二層:本次匯入授權失敗</font><br />
    									    <!--<font size="3px">&nbsp;&nbsp;&nbsp;&nbsp;☆第三層:扣款日期前累計等待授權紀錄</font>-->
    								    <p />
    							    </td>
    						    </tr>
    						    <tr>
    							    <td style="background-color:#EEEEE3">
    							    <font size="3" color="darkmagenta">
    								    &nbsp;&nbsp;&nbsp;1.6根據系統自動勾選已授權成功資料進行授權轉捐款作業</font><br/>
    								    <font size="3">
    									    &nbsp;&nbsp;&nbsp;&nbsp;☆點選按鈕【2.授權轉捐款】進行收據開立</font>
    							    </td>
    						    </tr>
           		    </table>
               </td>
             </tr>
           </table>
    </div>
    </form>
</body>
</html>

