﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ResetPassword.aspx.cs" Inherits="Online_ResetPassword" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Reset Your Password</title>
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
</head>
<body style="background-color: #EEE">
    <form id="form1" runat="server">
        <table width="860" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td width="436">
                    <img alt="aa" src="images/bar_logo.jpg" width="390" height="40" />
                </td>
                <td width="424" style="background-image: url('images/bar_item.jpg')">
                    <table width="95%" border="0" align="center" cellpadding="5" cellspacing="5">
                        <tr>
                            <td align="left" valign="bottom">
                                <strong><font color="#003399">Reset Your Password</font></strong>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <div style="text-align: center">
            <ul>
                <asp:Label ID="lblStep1" runat="server" Text=" Your Donation Items >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep2" runat="server" Text=" Donor Information >> " ForeColor="Red" Font-Size="Medium" Font-Bold="true" />
                <asp:Label ID="lblStep3" runat="server" Text=" Confirm Donation >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep4" runat="server" Text=" Pledge Information >> " ForeColor="MediumBlue" />
                <asp:Label ID="lblStep5" runat="server" Text=" Confirm Payment " ForeColor="MediumBlue" />
            </ul>
        </div>
        <div style="text-align: center">
        <strong><font color="#003399">Please enter your name and your email. A reset intruction will be sent to you.</font></strong>
        </div>
        <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td>
                    <table width="600" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v2">
                        <tr>
                            <th align="right">
                                <span class="necessary">Name：</span>

                            </th>
                            <td>
                                <asp:TextBox ID="txtDonorName" runat="server" Width="300px"></asp:TextBox>

                            </td>
                        </tr>
                        <tr>
                            <th align="right">
                                <span class="necessary">Email：</span>

                            </th>
                            <td>
                                <asp:TextBox ID="txtEmail" runat="server" Width="300px" TextMode="Email"></asp:TextBox>

                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td>&nbsp;
                            </td>
                            <td width="100">
                                <asp:Button ID="Button1" runat="server" class="Online Online_Button3" Text="Submit" OnClick="Button1_Click" />
                            </td>
                            <td width="10">&nbsp;
                            </td>
                            <td width="100">
                                <asp:Button ID="Button2" runat="server" class="Online Online_Button3" Text="Cancel" OnClick="Button2_Click" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

    </form>
</body>
</html>
