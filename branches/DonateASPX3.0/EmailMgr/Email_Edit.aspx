<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Email_Edit.aspx.cs" Inherits="EmailMgr_Email_Edit" ValidateRequest="false"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title> 郵件內容【修改】</title>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" src="../ckeditor/ckeditor.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            InitMenu();
        });
    </script>
    <script type="text/javascript">
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var cnt = 0;
            var sName = "";
            var ddlEmailMgr_Type = document.getElementById('ddlEmailMgr_Type');
            var tbxEmailMgr_Subject = document.getElementById('tbxEmailMgr_Subject');
            if (ddlEmailMgr_Type.value == "") {
                strRet += "郵件類別 ";
            }
            if (tbxEmailMgr_Subject.value == "") {
                strRet += "郵件標題 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet)
                return false;
            }
            cnt = 0;
            sName = tbxEmailMgr_Subject.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 60) {
                alert('郵件標題 欄位長度超過限制！');
                return false;
            }
            else {
                return true;
            }
        }
    </script>
    </head>
<body class="body">
    <form id="Form1" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" /> 
        郵件內容【修改】
    </h1>
     <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
           <td width="20%" class="font11" align="center" valign="top" bgcolor="EEEEE3" rowspan = "1">
                 <asp:dropdownlist runat="server" ID="ddlEmailMgr_Type_Banner" CssClass="font9" 
                     AutoPostBack="True" 
                     onselectedindexchanged="ddlEmailMgr_Type_Banner_SelectedIndexChanged"></asp:dropdownlist>
                 <br/>
                 <asp:Label ID="lblGridList" runat="server"></asp:Label>
          </td>
	        <td width="80%" class="font11" valign="top" rowspan = "2">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">

                      <tr>
                        <th align="right" colspan="1" width="13%">
                            郵件類別：
                        </th>
                        <td width="22%" align="left" colspan="1">
                            <asp:dropdownlist runat="server" ID="ddlEmailMgr_Type" CssClass="font9"></asp:dropdownlist>
                        </td>
                        <th align="right" colspan="1" width="13%">
                            郵件標題：
                        </th>
                        <td width="52%" align="left" colspan="1">
                            <asp:TextBox ID="tbxEmailMgr_Subject" maxlength="100" size="52" runat="server"></asp:TextBox>
                        </td>
                      </tr>
                      <tr> 
                        <th align="right" colspan="1">
                            郵件內容：
                        </th>
                        <td colspan="3" height="200">
                           <asp:TextBox ID="tbxEmailMgr_Desc" runat="server" TextMode="MultiLine" Rows="10" 
                               Columns="80" ClientIDMode="Static"></asp:TextBox>
                           <script type="text/javascript">
                               CKEDITOR.replace('<%= tbxEmailMgr_Desc.ClientID %>', { skin: 'kama' });
                           </script>
                  </td>
                </tr>
              </table>
              <div class="function">
                  <asp:Button ID="btnEdit" class="npoButton npoButton_Modify" runat="server" 
                      Text="修改" OnClientClick= "return CheckFieldMustFillBasic();" onclick="btnEdit_Click"/>
                  <asp:Button ID="btnDel" class="npoButton npoButton_Del" runat="server" 
                      Text="刪除" OnClientClick="return confirm('您是否確定要刪除？')" onclick="btnDel_Click"/>
                  <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
                      Text="離開" onclick="btnExit_Click" />
              </div>
            </center></div>
          </td>
        </tr>
        <tr>
            <td width="20%" class="font11" align="center" valign="top" bgcolor="EEEEE3" rowspan = "1">
                 <asp:Label ID="lblGridList2" runat="server"></asp:Label>
          </td>
        </tr>
      </table>
    </div>
    
    </form>
</body>
</html>


