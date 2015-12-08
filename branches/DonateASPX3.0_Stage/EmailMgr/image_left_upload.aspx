<%@ Page Language="C#" AutoEventWireup="true" CodeFile="image_left_upload.aspx.cs" Inherits="Function_image_left_upload" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1">
    <title>上傳圖檔 </title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />    
    <%--<link rel="stylesheet" href="../css/validationEngine.jquery.css" type="text/css"/>
    <link rel="stylesheet" href="../css/template.css" type="text/css"/>--%>
    <script src="../include/jquery-1.7.js" type="text/javascript" charset="utf-8"></script>
    <%--<script src="../include/jquery.validationEngine-zh_TW.js" type="text/javascript" charset="utf-8"></script>
    <script src="../include/jquery.validationEngine.js" type="text/javascript" charset="utf-8"></script>--%>
    <script type="text/javascript">
    jQuery(document).ready(function () {
        // binds form submission and fields to the validation engine
        //jQuery("#form").validationEngine();
    });
    </script>
</head>
<body class="body">
    <form id="form" method="post" runat="server">
        <asp:HiddenField ID="HFD_item" runat="server" />
        <asp:HiddenField ID="HFD_ser_no" runat="server" />
        <asp:HiddenField ID="HFD_subject" runat="server" />
        <asp:HiddenField ID="HFD_img_width" runat="server" />
        <asp:HiddenField ID="HFD_img_height" runat="server" />
        <asp:HiddenField ID="HFD_code_id" runat="server" />
        <asp:HiddenField ID="HFD_id" runat="server" />        
        <div align="center">
            <center>
                <table border="0" width="100%" class="table_v">
                    <tr>
                        <td width="100%">
                            <div align="center">
                                <center>
                                    <table width="100%" border="1" class="table_v">
                                        <tr>
                                            <th align="right" width="17%" height="22">
                                                上傳<br />
                                                圖檔：
                                            </th>
                                            <td align="left" width="83%" height="18">
                                                &nbsp;<asp:FileUpload ID="FileUpload1" runat="server" Width="370px" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <th align="right" height="22">
                                                圖檔<br />
                                                縮圖：
                                            </th>
                                            <td align="left" height="18">&nbsp;
                                                <input type="radio" name="resize" value="N" id="resizeNo"  checked="checked" >原圖大小
		                                        <input type="radio" name="resize" value="Y" id="resizeYes" >縮圖： 		                  
		                                        寬度<asp:TextBox runat="server" ID="img_width" size="8" MaxLength="8" onchange="Resize_OnFocus();" class="validate[custom[onlyNumberSp]]"></asp:TextBox>
		                                        或高度<asp:TextBox runat="server" ID="img_height" size="8" MaxLength="8" onchange="Resize_OnFocus();" class="validate[custom[onlyNumberSp]]"></asp:TextBox>                                                
                                            </td>
                                        </tr>
                                    </table>
                                </center>
                            </div>
                        </td>
                    </tr>
                </table>
                <div class="function">
                    <asp:Button ID="btnSave" runat="server" Text="存檔" OnClick="btnSave_Click" CssClass="npoButton npoButton_New" />
                    <asp:Button ID="btnCancle" runat="server" Text="取消" OnClientClick="window.close();" CssClass="npoButton npoButton_Exit" />
                </div>
            </center>
        </div>
    </form>
</body>
</html>
<%--<script language="JavaScript"><!--
    function Resize_OnFocus() {
        document.form.resize[1].checked = true;
    }
--></script>--%>