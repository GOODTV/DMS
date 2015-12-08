<%@ Page Language="C#" AutoEventWireup="true" CodeFile="LinkedList.aspx.cs" Inherits="Contribute_LinkedList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>實物奉獻主檔維護</title>
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>

    <script type="text/javascript" language="javascript">
        function Linked_List_Add(name) {
            var Linked_Id = document.getElementById("HFD_Linked_Id").value;
            window.open("../ContributeMgr/LinkedList_Add.aspx?Linked_Id=" + Linked_Id, "LinkedList_Add", "scrollbars=no,status=no,toolbar=no,top=100,left=120,width=580,height=200");
        }
    </script>
    <style type="text/css">
        .style1
        {
            font-size: 12px;
        }
        .style2
        {
            height: 40px;
        }
    </style>
</head>
<body class="body">
    <form id="Form1" name="form" runat="server">
        <asp:HiddenField ID="HFD_Linked_Id" runat="server" />
        <asp:HiddenField ID="HFD_Ser_No" runat="server" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10"/>
        實物奉獻主檔維護
    </h1>
        <!--內容頁面_START-->
        <table width="100%">
            <tr>
                <td width="50%" colspan="3" valign="top">
                    <!--List頁面_START-->
                    <div>
                        <table border="0" width="100%" id="table1" class="table_v">
                            <tr>
                                <td width="100%" align="center">
                                    <asp:Button runat="server" Text=" 新增物品類別 " ID="btnNewLinkedFolder"  CssClass="npoButton npoButton_New" Width="140px" OnClientClick="Linked_List_Add('');return false;" />
                                </td>
                            </tr>
                        </table>
                        <!--主目錄列表-->
                        <div>
                            <asp:Label ID="Grid_List1" runat="server" Text=""></asp:Label>
                        </div>
                        </div>
                        <!--List頁面_END-->
                 </td>
              <%--</tr>
           </table>
           <table width="100%">
            <tr>--%>
                <td width="100%" colspan="3" valign="top">
                    <!--List頁面_START-->
                    <div>
                        <table border="0" width="100%" id="table2" class="table_v">
                            <tr>
                                <td width="100%" align="center">
                                    <span class="style1">公關贈品品項：</span>
                                    <asp:TextBox ID="tbxLinkedFolder" runat="server" BackColor="#FFE1AF" 
                                        Width="150px" ReadOnly="True"></asp:TextBox>
                                    <asp:Button runat="server" Text="刪除物品類別" ID="btnDelLinkedFolder"  CssClass="npoButton npoButton_Del" Width="140px" 
                                        onclick="btnDelLinkedFolder_Click"  OnClientClick= "return confirm( '您是否確定要刪除物品類別？'); " />
                                    <br />
                                </td>
                                <tr>
                                <td width="100%" align="center" class="style2">
                                    <span class="style1">次分類名稱：</span>
                                    <asp:TextBox ID="tbxNewLinkedItems" runat="server" ></asp:TextBox>
                                    <asp:Button runat="server" Text=" 新增次分類" ID="btnNewLinkedItems" CssClass="npoButton npoButton_New" Width="110px" 
                                        onclick="btnNewLinkedItems_Click"  />
                                    <span class="style1">
                                    <br />
                                    排序：</span>
                                    <asp:TextBox ID="tbxNewLinkedItems_Seq" runat="server" Width="50px" ></asp:TextBox>
                                </td>
                                </tr>
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
