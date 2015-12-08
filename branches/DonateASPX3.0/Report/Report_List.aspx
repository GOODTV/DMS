<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Report_List.aspx.cs" Inherits="Report_Report_List" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>捐款報表(客製)</title>
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
        });
        window.onload = initCalendar;
        function initCalendar() {
            //week1
            //Calendar.setup({
            //    inputField: "tbxDonateDateS_week1",   // id of the input field
            //    button: "imgtbxDonateDateS_week1"     // 與觸發動作的物件ID相同
            //});
            //Calendar.setup({
            //    inputField: "tbxDonateDateE_week1",   // id of the input field
            //    button: "imgtbxDonateDateE_week1"     // 與觸發動作的物件ID相同
            //});
            //week2
            Calendar.setup({
                inputField: "tbxDonateDateS_week2",   // id of the input field
                button: "imgtbxDonateDateS_week2"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_week2",   // id of the input field
                button: "imgtbxDonateDateE_week2"     // 與觸發動作的物件ID相同
            });
            //week3
            //Calendar.setup({
            //    inputField: "tbxDonateDateS_week3",   // id of the input field
            //    button: "imgtbxDonateDateS_week3"     // 與觸發動作的物件ID相同
            //});
            //Calendar.setup({
            //    inputField: "tbxDonateDateE_week3",   // id of the input field
            //    button: "imgtbxDonateDateE_week3"     // 與觸發動作的物件ID相同
            //});
            //Calendar.setup({
            //    inputField: "tbxAccumulateDateS_week3",   // id of the input field
            //    button: "imgtbxAccumulateDateS_week3"     // 與觸發動作的物件ID相同
            //});
            //Calendar.setup({
            //    inputField: "tbxAccumulateDateE_week3",   // id of the input field
            //    button: "imgtbxAccumulateDateE_week3"     // 與觸發動作的物件ID相同
            //});
            //week4
            Calendar.setup({
                inputField: "tbxDonateDateS_week4",   // id of the input field
                button: "imgtbxDonateDateS_week4"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_week4",   // id of the input field
                button: "imgtbxDonateDateE_week4"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxAccumulateDateS_week4",   // id of the input field
                button: "imgtbxAccumulateDateS_week4"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxAccumulateDateE_week4",   // id of the input field
                button: "imgtbxAccumulateDateE_week4"     // 與觸發動作的物件ID相同
            });
            //month5
            Calendar.setup({
                inputField: "tbxDonateDateS_month5",   // id of the input field
                button: "imgtbxDonateDateS_month5"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_month5",   // id of the input field
                button: "imgtbxDonateDateE_month5"     // 與觸發動作的物件ID相同
            });
            //season2
            Calendar.setup({
                inputField: "tbxDonateDateS_season2",   // id of the input field
                button: "imgtbxDonateDateS_season2"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_season2",   // id of the input field
                button: "imgtbxDonateDateE_season2"     // 與觸發動作的物件ID相同
            });
            //season4
            Calendar.setup({
                inputField: "tbxDonateDateS_season4",   // id of the input field
                button: "imgtbxDonateDateS_season4"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_season4",   // id of the input field
                button: "imgtbxDonateDateE_season4"     // 與觸發動作的物件ID相同
            });
            //year1
            Calendar.setup({
                inputField: "tbxDonateDateS_year1",   // id of the input field
                button: "imgtbxDonateDateS_year1"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_year1",   // id of the input field
                button: "imgtbxDonateDateE_year1"     // 與觸發動作的物件ID相同
            });
            //year3
            Calendar.setup({
                inputField: "tbxDonateDateS_year3",   // id of the input field
                button: "imgtbxDonateDateS_year3"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_year3",   // id of the input field
                button: "imgtbxDonateDateE_year3"     // 與觸發動作的物件ID相同
            });
            //year4
            Calendar.setup({
                inputField: "tbxDonateDateS_year4",   // id of the input field
                button: "imgtbxDonateDateS_year4"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_year4",   // id of the input field
                button: "imgtbxDonateDateE_year4"     // 與觸發動作的物件ID相同
            });
            //year5
            Calendar.setup({
                inputField: "tbxDonateDateS_year5",   // id of the input field
                button: "imgtbxDonateDateS_year5"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_year5",   // id of the input field
                button: "imgtbxDonateDateE_year5"     // 與觸發動作的物件ID相同
            });
            //year6
            //Calendar.setup({
            //    inputField: "tbxDonateDateS_year6",   // id of the input field
            //    button: "imgtbxDonateDateS_year6"     // 與觸發動作的物件ID相同
            //});
            //Calendar.setup({
            //    inputField: "tbxDonateDateE_year6",   // id of the input field
            //    button: "imgtbxDonateDateE_year6"     // 與觸發動作的物件ID相同
            //});
            //year7
            Calendar.setup({
                inputField: "tbxLast_Donate_Date_year7",   // id of the input field
                button: "imgtbxLast_Donate_Date_year7"     // 與觸發動作的物件ID相同
            });
            //year8
            //Calendar.setup({
            //    inputField: "tbxDonateDateS_year8",   // id of the input field
            //    button: "imgtbxDonateDateS_year8"     // 與觸發動作的物件ID相同
            //});
            //Calendar.setup({
            //    inputField: "tbxDonateDateE_year8",   // id of the input field
            //    button: "imgtbxDonateDateE_year8"     // 與觸發動作的物件ID相同
            //});
            //year9
            Calendar.setup({
                inputField: "tbxDonateDateS_year9",   // id of the input field
                button: "imgtbxDonateDateS_year9"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_year9",   // id of the input field
                button: "imgtbxDonateDateE_year9"     // 與觸發動作的物件ID相同
            });
            //other1
            Calendar.setup({
                inputField: "tbxDonateDateS_other1",   // id of the input field
                button: "imgtbxDonateDateS_other1"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_other1",   // id of the input field
                button: "imgtbxDonateDateE_other1"     // 與觸發動作的物件ID相同
            });
            //other2
            //Calendar.setup({
            //    inputField: "tbxContributeDateS_other2",   // id of the input field
            //    button: "imgtbxContributeDateS_other2"     // 與觸發動作的物件ID相同
            //});
            //Calendar.setup({
            //    inputField: "tbxContributeDateE_other2",   // id of the input field
            //    button: "imgtbxContributeDateE_other2"     // 與觸發動作的物件ID相同
            //});
            //other4
            Calendar.setup({
                inputField: "tbxDonateDateS_other4",   // id of the input field
                button: "imgtbxDonateDateS_other4"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_other4",   // id of the input field
                button: "imgtbxDonateDateE_other4"     // 與觸發動作的物件ID相同
            });
            //other6
            Calendar.setup({
                inputField: "tbxDonateDateS_other6",   // id of the input field
                button: "imgtbxDonateDateS_other6"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_other6",   // id of the input field
                button: "imgtbxDonateDateE_other6"     // 與觸發動作的物件ID相同
            });
            //other7
            Calendar.setup({
                inputField: "tbxDonateDateS_other7",   // id of the input field
                button: "imgtbxDonateDateS_other7"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonateDateE_other7",   // id of the input field
                button: "imgtbxDonateDateE_other7"     // 與觸發動作的物件ID相同
            });
        }
        function CheckFieldMustFillBasic_week2() {
            var strRet = "";
            var tbxDonateDateS_week2 = document.getElementById('tbxDonateDateS_week2');
            var tbxDonateDateE_week2 = document.getElementById('tbxDonateDateE_week2');
            if (tbxDonateDateS_week2.value == "" || tbxDonateDateE_week2.value == "") {
                strRet += "捐款日期 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet)
                return false;
            }
            else {
                if (confirm('您是否確定要將查詢結果列表？') == false) {
                    return false;
                }
                else {
                    return true;
                }
            }
        }
        function CheckFieldMustFillBasic_other2() {
            var strRet = "";
            var tbxContributeDateS_other2 = document.getElementById('tbxContributeDateS_other2');
            var tbxContributeDateE_other2 = document.getElementById('tbxContributeDateE_other2');
            if (tbxContributeDateS_other2.value == "" || tbxContributeDateE_other2.value == "") {
                strRet += "贈送日期 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet)
                return false;
            }
            else {
                if (confirm('您是否確定要將查詢結果列表？') == false) {
                    return false;
                }
                else {
                    return true;
                }
            }
        }
        function Print(PrintType) {
            window.open('../Report/DonateAccounQry_Print.aspx', 'DonateAccounQry_Print', 'toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600', '');
        }
        function PanelShow(Panel_id) {
            PanelAllHide();
            document.getElementById(Panel_id).style.display = "block";
        }
        function PanelAllHide() {
            //$("#Panel_week1").css('display', 'none');
            $("#Panel_week2").css('display', 'none');
            //$("#Panel_week3").css('display', 'none');
            $("#Panel_week4").css('display', 'none');
            $("#Panel_month1").css('display', 'none');
            $("#Panel_month2").css('display', 'none');
            $("#Panel_month3").css('display', 'none');
            $("#Panel_month4").css('display', 'none');
            $("#Panel_month5").css('display', 'none');
            //$("#Panel_season1").css('display', 'none');
            $("#Panel_season2").css('display', 'none');
            $("#Panel_season4").css('display', 'none');
            $("#Panel_year1").css('display', 'none');
            //$("#Panel_year2").css('display', 'none');
            $("#Panel_year3").css('display', 'none');
            $("#Panel_year4").css('display', 'none');
            $("#Panel_year5").css('display', 'none');
            //$("#Panel_year6").css('display', 'none');
            $("#Panel_year7").css('display', 'none');
            //$("#Panel_year8").css('display', 'none');
            $("#Panel_year9").css('display', 'none');
            $("#Panel_year10").css('display', 'none');
            $("#Panel_other1").css('display', 'none');
            //$("#Panel_other2").css('display', 'none');
            $("#Panel_other3").css('display', 'none');
            $("#Panel_other4").css('display', 'none');
            $("#Panel_other5").css('display', 'none');
            $("#Panel_other6").css('display', 'none');
            $("#Panel_other7").css('display', 'none');
            $("#Panel_other8").css('display', 'none');
        }

    </script>
    <style type="text/css">
      .likecss {
        cursor: pointer;
      }
      .likecss:hover {
        color:blue;
      }
  </style>
    </head>
<body class="body" >
    <form id="Form1" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        捐款報表(客製)</h1>
    <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
	    <tr>
		    <td><b><font color="#8B0000" size="3">作業報表</font></b>
		    </td>
		    <td><b><font color="#8B0000" size="3">捐款服務</font></b>
		    </td>
		    <td><b><font color="#8B0000" size="3">海外捐款</font></b>
		    </td>
		    <td><b><font color="#8B0000" size="3">分析報表</font></b>
		    </td>
	    </tr>
        <tr>
		    <td height="20"><div id="week4" class="likecss" onclick="PanelShow('Panel_week4');"><font size="2">週報用大額捐款人明細</font></div>
		    </td>
		    <td><div id="week2" class="likecss" onclick="PanelShow('Panel_week2');"><font size="2">新捐款人明細</font></div>
		    </td>
		    <td><div id="year3" class="likecss" onclick="PanelShow('Panel_year3');"><font size="2">國外捐款總額明細</font></div>
		    </td>
		    <td><div id="other6" class="likecss" onclick="PanelShow('Panel_other6');"><font size="2">捐款人單項資料分析</font></div>
		    </td>
	    </tr>
	    <tr>
		    <td height="20"><div id="year4" class="likecss" onclick="PanelShow('Panel_year4');"><font size="2">台灣捐款總額明細</font></div>
		    </td>
		    <td><div id="month1" class="likecss" onclick="PanelShow('Panel_month1');"><font size="2">捐款人生日明細</font></div>
		    </td>
		    <td><div id="year9" class="likecss" onclick="PanelShow('Panel_year9');"><font size="2">中國捐款總額明細</font></div>
		    </td>
		    <td><div id="other1" class="likecss" onclick="PanelShow('Panel_other1');"><font size="2">人數統計分析</font></div>
		    </td>
	    </tr>
	    <tr>
		    <td height="20"><div id="year1" class="likecss" onclick="PanelShow('Panel_year1');"><font size="2">國內外捐款用途總額明細</font></div>
		    </td>
		    <td><div id="month2" class="likecss" onclick="PanelShow('Panel_month2');"><font size="2">定期奉獻授權到期明細</font></div>
		    </td>
		    <td>
		    </td>
		    <td><div id="month5" class="likecss" onclick="PanelShow('Panel_month5');"><font size="2">捐款方式分析</font></div>
		    </td>
	    </tr>
	    <tr>
		    <td height="20"><div id="year5" class="likecss" onclick="PanelShow('Panel_year5');"><font size="2">捐款人身份總和明細</font></div>
		    </td>
		    <td><div id="month3" class="likecss" onclick="PanelShow('Panel_month3');"><font size="2">月刊郵寄明細</font></div>
		    </td>
		    <td>
		    </td>
		    <td><div id="other5" class="likecss" onclick="PanelShow('Panel_other5');"><font size="2">讀者分析</font></div>
		    </td>
	    </tr>
	    <tr>
		    <td height="20"><div id="other3" class="likecss" onclick="PanelShow('Panel_other3');"><font size="2">雷同聯絡資料明細</font></div>
		    </td>
		    <td><div id="month4" class="likecss" onclick="PanelShow('Panel_month4');"><font size="2">捐款人EMAIL明細</font></div>
		    </td>
		    <td>
		    </td>
		    <td><div id="other7" class="likecss" onclick="PanelShow('Panel_other7');"><font size="2">奉獻動機及收視管道統計分析</font></div>
		    </td>
	    </tr>
	    <tr>
		    <td height="20"><div id="season4" class="likecss" onclick="PanelShow('Panel_season4');"><font size="2">單筆及總捐款額明細</font></div>
		    </td>
		    <td>
		    </td>
		    <td>
		    </td>
		    <td><div id="other8" class="likecss" onclick="PanelShow('Panel_other8');"><font size="2">個別奉獻動機及收視管道統計分析</font></div>
		    </td>
	    </tr>
	    <tr>
		    <td height="20"><div id="other4" class="likecss" onclick="PanelShow('Panel_other4');"><font size="2">捐款人/讀者資料明細</font></div>
		    </td>
		    <td>
		    </td>
		    <td>
		    </td>
		    <td>
		    </td>
	    </tr>
	    <tr>
		    <td height="20"><div id="year7" class="likecss" onclick="PanelShow('Panel_year7');"><font size="2">單年首捐人逐年捐款明細</font></div>
		    </td>
		    <td>
		    </td>
		    <td>
		    </td>
		    <td>
		    </td>
	    </tr>
	    <tr>
		    <td height="20"><div id="year10" class="likecss" onclick="PanelShow('Panel_year10');"><font size="2">年度收據列印明細</font></div>
		    </td>
		    <td>
		    </td>
		    <td>
		    </td>
		    <td>
		    </td>
	    </tr>
        <tr>
		    <td height="20"><div id="season2" class="likecss" onclick="PanelShow('Panel_season2');"><font size="2">單筆捐款金額明細</font></div>
		    </td>
		    <td>
		    </td>
		    <td>
		    </td>
		    <td>
		    </td>
	    </tr>
    </table>
    <%-- <asp:Panel ID="Panel_week1" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                建台/非建台奉獻級距表 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_week1" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款用途： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDonate_Purpose_week1" runat="server" CssClass="font9">
                        <asp:ListItem>建台</asp:ListItem>
                        <asp:ListItem>非建台</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款日期： 
                </th>
                <td align="left" colspan="5">
                    <asp:TextBox ID="tbxDonateDateS_week1" runat="server" 
                        onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_week1" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_week1" runat="server" 
                        onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_week1" alt="" src="../images/date.gif" />
                </td>
            </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_week1" CssClass="npoButton npoButton_Print" 
            runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');"  />
        <asp:Button ID="btnToxls_week1" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>--%>

    <asp:Panel ID="Panel_week2" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                新捐款人明細
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_week2" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1"><font color="red">*</font>
                    捐款日期： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_week2" runat="server" 
                        onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_week2" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_week2" runat="server" 
                        onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_week2" alt="" src="../images/date.gif" />
                </td>
                <th align="right" colspan="1">
                    捐款總金額起訖： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Total_Begin_week2" runat="server" Width="90px"></asp:TextBox> ~ 
                    <asp:TextBox ID="tbxDonate_Total_End_week2" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款用途：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBoxList ID="cblDonate_Purpose_week2" runat="server" 
                        RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                </td> 
           </tr>
           <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_week2" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_week2" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_week2" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_week2" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick="return CheckFieldMustFillBasic_week2()" />
        <asp:Button ID="btnToxls_week2" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <%-- <asp:Panel ID="Panel_week3" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                查詢單筆捐款金額 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_week3" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_week3" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款期間： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_week3" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_week3" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_week3" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_week3" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    累計期間： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxAccumulateDateS_week3" runat="server" 
                        onchange="CheckDateFormat(this, '累計期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxAccumulateDateS_week3" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxAccumulateDateE_week3" runat="server" 
                        onchange="CheckDateFormat(this, '累計期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxAccumulateDateE_week3" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_week3" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_week3" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_week3" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_week3" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_week3" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>--%>
    <asp:Panel ID="Panel_week4" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                週報用大額捐款人明細
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_week4" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_week4" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款期間： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_week4" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_week4" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_week4" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_week4" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    累計期間： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxAccumulateDateS_week4" runat="server" 
                        onchange="CheckDateFormat(this, '累計期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxAccumulateDateS_week4" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxAccumulateDateE_week4" runat="server" 
                        onchange="CheckDateFormat(this, '累計期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxAccumulateDateE_week4" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_week4" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_week4" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_week4" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_week4" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_week4" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>
    <asp:Panel ID="Panel_month1" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                捐款人生日明細 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_month1" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_month1" runat="server" Width="90px"></asp:TextBox>
                </td>
                <th align="right" colspan="1">
                    捐款總金額起訖： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Total_Begin_month1" runat="server" Width="90px"></asp:TextBox> 
                    &nbsp;~ 
                    <asp:TextBox ID="tbxDonate_Total_End_month1" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    身分別： 
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBoxList ID="cblDonor_Type_month1" runat="server" 
                        RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    生日月份： 
                </th>
                <td align="left" colspan="2">
                    <asp:DropDownList ID="ddlBirthday_Month_month1" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_month1" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_month1" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_month1" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_month1" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_month1" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_month2" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                定期奉獻授權到期明細 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_month2" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    授權終止年月： 
                </th>
                <td align="left" colspan="2">
                    <asp:DropDownList ID="ddlYear_month2" runat="server" CssClass="font9">
                    </asp:DropDownList>年
                    <asp:DropDownList ID="ddlMonth_month2" runat="server" CssClass="font9">
                    </asp:DropDownList>月
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    授權方式： 
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBoxList ID="cblDonate_Payment_month2" runat="server" 
                        RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_month2" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_month2" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_month2" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_month2" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_month2" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_month3" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                月刊郵寄明細 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_month3" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    身分別： 
                </th>
                <td align="left" colspan="5">
                    <asp:RadioButtonList ID="rblDonor_Type_month3" runat="server" 
                        RepeatDirection="Horizontal">
                        <asp:ListItem Selected="True" Value="N">捐款人</asp:ListItem>
                        <asp:ListItem Value="Y">讀者</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_month3" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_month3" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_month3" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_month3" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_month3" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_month4" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                捐款人EMAIL明細
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_month4" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_month4" runat="server" Width="90px"></asp:TextBox>
                </td>
                <th align="right" colspan="1">
                    捐款累計金額(大於)：
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Total_Amt_month4" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款用途：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBoxList ID="cblDonate_Purpose_month4" runat="server" 
                        RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                </td> 
           </tr>
           <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_month4" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_month4" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_month4" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_month4" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_month4" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_month5" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                捐款方式分析
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_month5" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    統計期間：
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_month5" runat="server" 
                        onchange="CheckDateFormat(this, '統計期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_month5" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_month5" runat="server" 
                        onchange="CheckDateFormat(this, '統計期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_month5" alt="" src="../images/date.gif" />
                </td>
            </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_month5" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_month5" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <%-- <asp:Panel ID="Panel_season1" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                非建台奉獻統計分析表
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_season1" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款期間： 
                </th>
                <td align="left" colspan="5">
                    <asp:TextBox ID="tbxDonateDateS_season1" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_season1" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_season1" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_season1" alt="" src="../images/date.gif" />
                </td>
            </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_season1" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_season1" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>--%>

    <asp:Panel ID="Panel_season2" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                單筆捐款金額明細 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_season2" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_season2" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款期間： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_season2" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_season2" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_season2" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_season2" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_season2" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_season2" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_season2" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_season2" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_season2" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_season3" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                捐款用途各項總額明細表
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_season3" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_season3" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款累計金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Total_Amt_season3" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款用途： 
                </th>
                <td align="left" colspan="2">
                    <asp:CheckBoxList ID="cblDonate_Purpose_season3" runat="server" 
                        RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                </td>
            </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_season3" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_season3" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_season4" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                單筆及總捐款額明細
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_season4" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_season4" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款累計金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Total_Amt_season4" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款期間： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_season4" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_season4" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_season4" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_season4" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_season4" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_season4" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_season4" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_season4" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_season4" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_year1" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                國內外捐款用途總額明細 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_year1" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_year1" runat="server" Width="90px"></asp:TextBox>
                </td>
                <th align="right" colspan="1">
                    捐款總金額起訖： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Total_Begin_year1" runat="server" Width="90px"></asp:TextBox> &nbsp;~ 
                    <asp:TextBox ID="tbxDonate_Total_End_year1" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款期間： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_year1" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_year1" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_year1" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_year1" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款用途：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBoxList ID="cblDonate_Purpose_year1" runat="server" 
                        RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                </td> 
           </tr>
           <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_year1" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_year1" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_year1" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_year1" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_year1" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <%-- <asp:Panel ID="Panel_year2" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                捐款總額與月刊索取明細表
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_year2" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款年：
                </th>
                <td align="left" colspan="2">
                    <asp:DropDownList ID="ddlYear_year2" runat="server" CssClass="font9">
                    </asp:DropDownList>年
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_year2" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_year2" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_year2" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_year2" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_year2" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>--%>

    <asp:Panel ID="Panel_year3" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                國外捐款總額明細
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_year3" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_year3" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款累積金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Total_Amt_year3" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款期間： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_year3" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_year3" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_year3" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_year3" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_ErrAddress_year3" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_year3" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_year3" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_year3" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_year4" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                台灣捐款總額明細 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_year4" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_year4" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款累積金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Total_Amt_year4" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款期間： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_year4" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_year4" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_year4" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_year4" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    縣市：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBoxList ID="cblCity_year4" runat="server" 
                        RepeatDirection="Horizontal" RepeatLayout="Flow" Width="450px">
                    </asp:CheckBoxList>
                </td> 
           </tr>
           <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_ErrAddress_year4" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_year4" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
           <tr>
               <th align="right" colspan="1">
                    備註說明：
                </th>
                <td align="left" colspan="5">
                    北部範圍為：台北市、新北市、基隆市、桃園縣、新竹市、新竹縣、宜蘭縣<br />
                    中部範圍為：苗栗縣、台中市、彰化縣、南投縣、雲林縣<br />
                    南部範圍為：嘉義市、嘉義縣、台南市、高雄市、屏東縣<br />
                    東部範圍為：花蓮縣、台東縣<br />
                    外島範圍為：澎湖縣、金門縣、連江縣
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_year4" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_year4" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_year5" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                捐款人身份總和明細 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_year5" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_year5" runat="server" Width="90px"></asp:TextBox>
                </td>
                <th align="right" colspan="1">
                    捐款累積金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Total_Amt_year5" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    身分別： 
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBoxList ID="cblDonor_Type_year5" runat="server" 
                        RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款期間：
                </th>
                <td align="left" colspan="5">
                    <asp:TextBox ID="tbxDonateDateS_year5" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_year5" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_year5" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_year5" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_year5" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_year5" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_year5" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_year5" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_year5" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <%-- <asp:Panel ID="Panel_year6" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                期間單筆捐款金額次數明細表 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_year6" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_year6" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款年：
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_year6" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_year6" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_year6" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_year6" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_year6" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_year6" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_year6" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_year6" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_year6" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>--%>

    <asp:Panel ID="Panel_year7" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                單年首捐人逐年捐款明細
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_year7" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_year7" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款用途：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBoxList ID="cblDonate_Purpose_year7" runat="server" 
                        RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                </td> 
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款年：
                </th>
                <td align="left" colspan="2">
                    <asp:DropDownList ID="ddlYear_year7" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
                <th align="right" colspan="1">
                    首捐年：
                </th>
                <td align="left" colspan="2">
                    <asp:DropDownList ID="ddlFirst_DonateYear_year7" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    末捐日期： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxLast_Donate_Date_year7" runat="server" 
                        onchange="CheckDateFormat(this, '末捐日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxLast_Donate_Date_year7" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_year7" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_year7" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_year7" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_year7" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_year7" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <%-- <asp:Panel ID="Panel_year8" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                經常奉獻總額明細表 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_year8" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Amt_year8" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款累積金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Total_Amt_year8" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款期間： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_year8" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_year8" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_year8" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_year8" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_year8" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_year8" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_year8" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_year8" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_year8" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>--%>

    <asp:Panel ID="Panel_year9" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                中國捐款總額明細 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_year9" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款累積金額(大於)： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonate_Total_Amt_year9" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款期間： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_year9" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_year9" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_year9" runat="server" 
                        onchange="CheckDateFormat(this, '捐款期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_year9" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_ErrAddress_year9" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_year9" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_year9" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_year9" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_year10" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                年度收據列印明細
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_year10" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款年：
                </th>
                <td align="left" colspan="2">
                    <asp:DropDownList ID="ddlYear_year10" runat="server" CssClass="font9">
                    </asp:DropDownList>年
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款人編號： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonorId_year10" runat="server" Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_year10" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_year10" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_year10" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_year10" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_year10" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_other1" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                人數統計分析
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_other1" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款日期： 
                </th>
                <td align="left" colspan="5">
                    <asp:TextBox ID="tbxDonateDateS_other1" runat="server" 
                        onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_other1" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_other1" runat="server" 
                        onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_other1" alt="" src="../images/date.gif" />
                </td>
            </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_other1" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick="return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_other1" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <%-- <asp:Panel ID="Panel_other2" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                DVD贈品索取人報表
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_other2" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1"><font color="red">*</font>
                    贈送日期： 
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxContributeDateS_other2" runat="server" 
                        onchange="CheckDateFormat(this, '贈送日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxContributeDateS_other2" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxContributeE_other2" runat="server" 
                        onchange="CheckDateFormat(this, '贈送日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxContributeDateE_other2" alt="" src="../images/date.gif" />
                </td>
            <tr>
                <th align="right" colspan="1">
                    物品名稱：
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlLinked2Name_other2" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td> 
           </tr>
           <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_other2" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_other2" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_other2" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_other2" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick="return CheckFieldMustFillBasic_other2()" />
        <asp:Button ID="btnToxls_other2" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>--%>
    <asp:Panel ID="Panel_other3" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                雷同聯絡資料明細
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_other3" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1"><font color="red">*</font>
                    查詢：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_SameName_other3" runat="server" Checked="True" Text="同姓名" />
                    <asp:CheckBox ID="cbxIs_SameAddress_other3" runat="server"  Text="同地址" AutoPostBack="True" oncheckedchanged="cbxIs_Address_CheckedChanged"/>
                    <asp:CheckBox ID="cbxIs_SameTel_other3" runat="server"  Text="同電話" AutoPostBack="True" oncheckedchanged="cbxIs_Tel_CheckedChanged"/>
                </td> 
            </tr>
            <asp:Panel runat="server" ID="IsAddress_other3" Visible="false">
            <tr>
                <th align="right" colspan="1">
                    地址：
                </th>
                <td align="left" colspan="5">
                    <asp:RadioButtonList ID="rblAddress_other3" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Value="1" Selected="True">通訊地址</asp:ListItem>
                    <asp:ListItem Value="2">收據地址 </asp:ListItem>
                </asp:RadioButtonList>
                </td>
            </tr>
            </asp:Panel>
            <asp:Panel runat="server" ID="IsTel_other3" Visible="false">
            <tr>
                <th align="right" colspan="1">
                    電話：
                </th>
                <td align="left" colspan="5">
                    <asp:RadioButtonList ID="rblTel_other3" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem Value="1" Selected="True">手機</asp:ListItem>
                    <asp:ListItem Value="2">室內電話 </asp:ListItem>
                </asp:RadioButtonList>
                </td>
            </tr>
            </asp:Panel>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_other3" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick="return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_other3" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_other4" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                捐款人/讀者資料明細 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_other4" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    身分別： 
                </th>
                <td align="left" colspan="5">
                    <asp:RadioButtonList ID="rblDonor_Type_other4" runat="server" 
                        RepeatDirection="Horizontal">
                        <asp:ListItem Selected="True" Value="N">捐款人</asp:ListItem>
                        <asp:ListItem Value="Y">讀者</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款日期： 
                </th>
                <td align="left" colspan="5">
                    <asp:TextBox ID="tbxDonateDateS_other4" runat="server" 
                        onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_other4" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_other4" runat="server" 
                        onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_other4" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_other4" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_other4" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_other4" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_other4" CssClass="npoButton npoButton_Print" 
            runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');"  />
        <asp:Button ID="btnToxls_other4" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_other5" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                讀者分析 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_other5" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    身分別： 
                </th>
                <td align="left" colspan="5">
                    <asp:RadioButtonList ID="rblDonor_Type_other5" runat="server" 
                        RepeatDirection="Horizontal">
                        <asp:ListItem Selected="True" Value="Y">讀者</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_other5" CssClass="npoButton npoButton_Print" 
            runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');"  />
        <asp:Button ID="btnToxls_other5" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_other6" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                捐款人單項資料分析 
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_other6" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    項目： 
                </th>
                <td align="left" colspan="5">
                    <asp:RadioButtonList ID="rbl_Type_other6" runat="server" 
                        RepeatDirection="Horizontal">
                        <asp:ListItem Selected="True" Value="1">性別</asp:ListItem>
                        <asp:ListItem Value="2">年齡</asp:ListItem>
                        <asp:ListItem Value="3">通訊縣市</asp:ListItem>
                        <asp:ListItem Value="4">身份別</asp:ListItem>
                        <asp:ListItem Value="5">通訊資料</asp:ListItem>
                        <asp:ListItem Value="6">訂閱月刊</asp:ListItem>
                        <asp:ListItem Value="7">奉獻金額</asp:ListItem>
                        <asp:ListItem Value="8">捐款次數</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款日期： 
                </th>
                <td align="left" colspan="5">
                    <asp:TextBox ID="tbxDonateDateS_other6" runat="server" 
                        onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_other6" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_other6" runat="server" 
                        onchange="CheckDateFormat(this, '捐款日期');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_other6" alt="" src="../images/date.gif" />
                </td>
            </tr>
            <tr>
               <th align="right" colspan="1">
                    不含：
                </th>
                <td align="left" colspan="5">
                    <asp:CheckBox ID="cbxIs_Abroad_other6" runat="server" Checked="True" Text="海外地址" />
                    <asp:CheckBox ID="cbxIs_ErrAddress_other6" runat="server" Checked="True" Text="錯址" />
                    <asp:CheckBox ID="cbxSex_other6" runat="server" Checked="True" Text="歿" />
                </td> 
           </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_other6" CssClass="npoButton npoButton_Print" 
            runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');"  />
        <asp:Button ID="btnToxls_other6" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_other7" runat="server" align="center">
    <table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                奉獻動機及收視管道統計分析
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_other7" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    統計期間：
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonateDateS_other7" runat="server" 
                        onchange="CheckDateFormat(this, '統計期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateS_other7" alt="" src="../images/date.gif" />
                    ~
                    <asp:TextBox ID="tbxDonateDateE_other7" runat="server" 
                        onchange="CheckDateFormat(this, '統計期間');" Width="90px"></asp:TextBox>
                    <img id="imgtbxDonateDateE_other7" alt="" src="../images/date.gif" />
                </td>
            </tr>
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_other7" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_other7" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>

    <asp:Panel ID="Panel_other8" runat="server" align="center">
    <table width="700" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <caption>
            <h1 style=" padding-bottom:0px;">
                <img src="../images/h1_arr.gif" alt="" />
                個別奉獻動機及收視管道統計分析
            </h1>
            <tr>
                <th align="right" colspan="1">
                    機構： 
                </th>
                <td align="left" colspan="1">
                    <asp:DropDownList ID="ddlDept_other8" runat="server" CssClass="font9">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <th align="right" colspan="1">
                    捐款人姓名：
                </th>
                <td align="left" colspan="2">
                    <asp:TextBox ID="tbxDonorName_other8" runat="server"  Width="90px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <th align="right" style="width:150px;">奉獻動機(可複選)：</th>
                <td align="left">
                    <asp:CheckBox ID="DonateMotive1" runat="server" Text="支持媒體宣教大平台，可廣傳福音" /><br />
                    <asp:CheckBox ID="DonateMotive2" runat="server" Text="個人靈命得造就" /><br />
                    <asp:CheckBox ID="DonateMotive3" runat="server" Text="支持優質節目製作" /><br />
                    <asp:CheckBox ID="DonateMotive4" runat="server" Text="支持GOOD TV家庭事工" /><br />
                    <%--<asp:CheckBox ID="DonateMotive5" runat="server" Text="抵扣稅額" /><br />--%>
                    <asp:CheckBox ID="DonateMotive5" runat="server" Text="感恩奉獻" /><br />
                    <asp:CheckBox ID="DonateMotive6" runat="server" Text="其他" />
                </td>
                </tr>
                <tr>
                <th align="right" style="width:150px;">收看管道(可複選)：</th>
                <td align="left">
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
        </caption>
    </table>
    <div class="function">
        <asp:Button ID="btnPrint_other8" CssClass="npoButton npoButton_Print" runat="server"  Width="20mm"
                    Text="報表" onclick="btnPrint_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果列表？');" />
        <asp:Button ID="btnToxls_other8" runat="server"  Width="20mm"
                    Text="匯出" CssClass="npoButton npoButton_Excel" onclick="btnToxls_Click"
                    OnClientClick=" return confirm('您是否確定要將查詢結果匯出？');" />
    </div>
    </asp:Panel>
    </div>
    </form>
</body>
</html>
