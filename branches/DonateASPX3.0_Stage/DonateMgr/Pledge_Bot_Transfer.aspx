<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Pledge_Bot_Transfer.aspx.cs" Inherits="DonateMgr_Pledge_Bot_Transfer" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>批次固定自動轉帳</title>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">

        // menu控制
        $(document).ready(function () {

            InitMenu();

        });

        function showFileImport() {
            $("#reg").fadeTo("fast", 0.6);
            $('#openFileImport').show();
        }

        function hideFileImport() {
            $("#reg").hide()
            $('#openFileImport').hide();
        }

        function FileImportCheck() {

            //檢查檔案欄位名稱
            var filename = document.getElementById("FileUpload");
            if (filename.value == "") {
                alert("檔案名稱不能有空白！");
                return false;
            }

            //檢查副檔名
            var validExtensions = ["TXT", "01R", "02R", "03R"];
            var ext = filename.value.substring(filename.value.lastIndexOf('.') + 1).toUpperCase();
            for (var i = 0; i < validExtensions.length; i++) {
                if (ext == validExtensions[i]) {

                    alert("後面程式未完成！");
                    return false;
                    //document.body.style.cursor = 'wait';
                    //return true;
                }
            }
            alert("此類副檔名不允許上傳！");
            return false;

        }

    </script>

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
        批次固定自動轉帳作業 <font color="red">[未完成]</font>
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
                <th>銀行回覆檔匯入：
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
                        Text="取消" OnClientClick="hideFileImport();return false;" />
                </td>
            </tr>
        </table>
    </div>
    <table>
        <td valign="top">
            <table width="830" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
                <tr>
                    <td align="left" colspan="8">
                        本匯入回覆檔
                        <asp:Label ID="lblAccount_Count" runat="server" Text=""></asp:Label>
                        筆  總計金額
                        <asp:Label ID="lblAccount_Amount" runat="server" Text=""></asp:Label>
                        元 ╱ 信用卡授權成功扣款
                        <asp:Label ID="lblCard_Count" runat="server" Text=""></asp:Label>
                        筆  小計金額
                        <asp:Label ID="lblCard_Amount" runat="server" Text=""></asp:Label>
                        元 ╱ 信用卡授權失敗
                        <asp:Label ID="lblError_Count" runat="server" Text=""></asp:Label>
                        筆  小計金額
                        <asp:Label ID="lblError_Amount" runat="server" Text=""></asp:Label>
                        元
                    </td>
                </tr>
                <tr>
                    <td colspan="4">
                    </td>
                    <td align="right" colspan="4" >
                        <asp:Button ID="btnImport" CssClass="npoButton npoButton_Export" runat="server"
                            Text="匯入銀行回覆檔轉收據" OnClientClick="showFileImport();return false;" Width="190px"/>
                        <asp:Button ID="btnImport_CheckFalse" CssClass="npoButton npoButton_Excel" runat="server"
                            Text="匯出回覆檔授權失敗Excel檔" onclick="btnImport_CheckFalse_Click" Width="190px"/>
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
           		    <table width="330" border="0" cellspacing="0" cellpadding="0" align="left" >
           			    <tr height="35px">
    							    <td colspan="2" style="background-color:#EEEEE3" align="center">
    								    <b><font size="4px" color="brown">::批次固定自動轉帳作業說明::</font></b>
    							    </td>
    						    </tr>
    						    <tr align="center">
    							    <td colspan="2" style="background-color:#EEEEE3" >
    								    <font size="2px" color="chocolate">
    								    =========待收到銀行(台銀)回覆檔後=========</font>
    							    </td>
    					      </tr>
    						    <tr style="background-color:#EEEEE3" valign="top">
                                    <td><font size="3px" color="darkmagenta">1.</font></td>
    							    <td >
    								    <font size="3px" color="darkmagenta">匯入「回覆檔」</font><br/>
    								    <font size="3px">☆點選按鈕</font><font size="3px" color="red">【匯入銀行回覆檔轉收據】</font><font size="3px">進行匯入</font>
    							    </td>
    						    </tr>
    						    <tr><td colspan="2">&nbsp;</td>
    						    </tr>
    						    <tr style="background-color:#EEEEE3" valign="top">
                                    <td><font size="3px" color="darkmagenta">2.</font></td>
    							    <td>
    								    <font size="3px" color="darkmagenta">頁面將呈現在本次匯入回覆檔的二層結果:</font><br/>
    								    <font size="3px">☆第一層:本次匯入授權成功轉收據明細(系統會自動將已授權成功資料轉收據)</font><br />
    						    	    <font size="3px">☆第二層:本次匯入授權失敗</font>
    							    </td>
    						    </tr>
           		    </table>
               </td>
           </table>
    </div>
    </form>
</body>
</html>
