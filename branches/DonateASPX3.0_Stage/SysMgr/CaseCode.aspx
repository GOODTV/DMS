<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CaseCode.aspx.cs" Inherits="SysMgr_CaseCode" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>代碼主檔維護</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript" language="javascript">
    </script>
    <style type="text/css">
        .style1
        {
            font-size: 12px;
        }
        </style>
</head>
<body class="body">
    <form id="Form1" name="form" runat="server">
        <asp:HiddenField ID="HFD_Uid" runat="server" />
        <asp:HiddenField ID="HFD_GroupName" runat="server" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10"/>
        代碼主檔維護
    </h1>
        <!--內容頁面_START-->
        <table width="100%">
            <tr>
                <td width="50%" colspan="2" valign="top">
                    <!--List頁面_START-->
                    <div>
                        <!--主目錄列表-->
                        <div>
                            <asp:Label ID="Grid_List1" runat="server" Text=""></asp:Label>
                        </div>
                        </div>
                        <!--List頁面_END-->
                 </td>
<%--              </tr>
           </table>
           <table width="100%">
            <tr>--%>
                <td width="100%" colspan="2" valign="top">
                    <!--List頁面_START-->
                    <div>
                        <table border="0" width="100%" id="table2" class="table_v" >
                            <tr>
                                <td width="50%" align="center" >
                                    <span class="style1">代碼：</span>
                                    <asp:TextBox ID="tbxCodeType" runat="server" CssClass="readonly"
                                        Width="150px" ReadOnly="True"></asp:TextBox>
                                    <span class="style1">&nbsp;&nbsp;&nbsp;&nbsp; 名稱：</span>
                                    <asp:TextBox ID="tbxCaseFolder" runat="server" CssClass="readonly" 
                                        Width="150px" ReadOnly="True"></asp:TextBox>
                                    <asp:Button runat="server" Text="刪除主檔" ID="btnDelCaseFolder"  CssClass="npoButton npoButton_Del" Width="100px" 
                                        onclick="btnDelCaseFolder_Click"  OnClientClick= "return confirm( '您是否確定要刪除主檔？'); " />
                                    <br />
                                    <hr />
                                    <span class="style1">選項：</span>
                                    <asp:TextBox ID="tbxNewCaseItems" runat="server" Width="250px" ></asp:TextBox>
                                    <asp:Button runat="server" Text=" 新增選項" ID="btnNewCaseItems" CssClass="npoButton npoButton_New" Width="100px" 
                                        onclick="btnNewCaseItems_Click"  />
                                    <span class="style1">
                                    <br />
                                    排序：</span>
                                    <asp:TextBox ID="tbxNewCaseID" runat="server" Width="100px" ></asp:TextBox>
                                </td>
                            </tr>
                        </table>
            <!--分目錄列表-->
                        <div>
                            <asp:Label ID="Grid_List2" runat="server" Text=""></asp:Label>
                            </div>
                        </div>
            <!--List頁面_END-->
                </td>
            </tr>
           </table>
        
        <!--內容頁面_END-->
    </form>
</body>
</html>
