<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DashBoard.aspx.cs" Inherits="Report_DashBoard" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
    <title></title>
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
    </script>
</head>
<body class="body">
    <form id="form1" name="form" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        DashBoard 
    </h1>
    <table width="100%">
        <tr>
            <td valign="top">
			    <table border="1" width="230" cellspacing="0" cellpadding="0"  style="BACKGROUND-COLOR: #EEEEE3">
				    <tr height="30px" style="BACKGROUND-COLOR: #F7FE2E">
					    <td align="center"><font color="green" size="3"><b>:::捐款資訊 Dash board:::</b></font></td>    		
				    </tr>
				    <tr height="30px"> 
					    <td align="left">◎天使總人數：<font color="blue"><asp:Label ID="lblAngel_Count" runat="server"></asp:Label></font><br/>◎讀者總人數：<font color="blue"><asp:Label ID="lblReader_Count" runat="server"></asp:Label></font></td>
				    </tr>				
				    <tr height="25px">
					    <td align="left" style="BACKGROUND-COLOR: #82FA58" ><center>今日 <asp:Label ID="lblDate_Today" runat="server"></asp:Label></center></td>
				    </tr>		
				    <tr height="50px">
					    <td align="left">◎累計奉獻筆數：<font color="blue"><asp:Label ID="lblToday_Count" runat="server"></asp:Label></font><br/>◎累計奉獻筆數(線上)：<font color="blue"><asp:Label ID="lblToday_Online_Count" runat="server"></asp:Label></font><br/>◎累計奉獻金額：<font color="blue"><asp:Label ID="lblToday_SumAmt" runat="server"></asp:Label></font><br/>◎平均奉獻金額：<font color="blue"><asp:Label ID="lblToday_DonateAmt_Avg" runat="server"></asp:Label></font></td>
				    </tr>		
				    <tr height="25px">
					    <td align="left" style="BACKGROUND-COLOR: #82FA58" ><center>本月 <asp:Label ID="lblThis_YearMonth" runat="server"></asp:Label></center></td>
				    </tr>	
				    <tr height="50px">
					    <td align="left">◎累計奉獻筆數：<font color="blue"><asp:Label ID="lblMonth_Count" runat="server"></asp:Label></font><br/>◎累計奉獻筆數(線上)：<font color="blue"><asp:Label ID="lblMonth_Online_Count" runat="server"></asp:Label></font><br/>◎累計奉獻金額：<font color="blue"><asp:Label ID="lblMonth_SumAmt" runat="server"></asp:Label></font><br/>◎平均奉獻金額：<font color="blue"><asp:Label ID="lblMonth_DonateAmt_Avg" runat="server"></asp:Label></font></td>
				    </tr>		
				    <tr height="25px">
					    <td align="left" style="BACKGROUND-COLOR: #82FA58" ><center>本年度 <asp:Label ID="lblThis_Year" runat="server"></asp:Label></center></td>
				    </tr>			
				    <tr height="50px">
					    <td align="left">◎累計奉獻筆數：<font color="blue"><asp:Label ID="lblYear_Count" runat="server"></asp:Label></font><br/>◎累計奉獻筆數(線上)：<font color="blue"><asp:Label ID="lblYear_Online_Count" runat="server"></asp:Label></font><br/>◎累計奉獻金額：<font color="blue"><asp:Label ID="lblYear_SumAmt" runat="server"></asp:Label></font><br/>◎平均奉獻金額：<font color="blue"><asp:Label ID="lblYear_DonateAmt_Avg" runat="server"></asp:Label></font></td>
				    </tr>				
				
				    <tr height="25px">
					    <td align="left" style="BACKGROUND-COLOR: #F781F3" ><center>上一年度 <asp:Label ID="lblLast_Year" runat="server"></asp:Label></center></td>
				    </tr>	
				    <tr height="50px">
					    <td align="left">◎奉獻總筆數：<font color="blue"><asp:Label ID="lblLastYear_Count" runat="server"></asp:Label></font><br/>◎奉獻總金額：<font color="blue"><asp:Label ID="lblLastYear_SumAmt" runat="server"></asp:Label></font><br/>◎奉獻平均金額：<font color="blue"><asp:Label ID="lblLastYear_DonateAmt_Avg" runat="server"></asp:Label></font></td>
				    </tr>		
				    <tr height="25px">
					    <td align="left" style="BACKGROUND-COLOR: #F781F3" ><center>上一年度本月 <asp:Label ID="lblLast_YearMonth" runat="server"></asp:Label></></center></td>
				    </tr>		
				    <tr height="50px">
					    <td align="left">◎奉獻總筆數：<font color="blue"><asp:Label ID="lblLastYearMonth_Count" runat="server"></asp:Label></font><br/>◎奉獻總金額：<font color="blue"><asp:Label ID="lblLastYearMonth_SumAmt" runat="server"></asp:Label></font><br/>◎奉獻平均金額：<font color="blue"><asp:Label ID="lblLastYearMonth_DonateAmt_Avg" runat="server"></asp:Label></font></td>
				    </tr>				
				
				    <tr height="25px">
					    <td align="left" style="BACKGROUND-COLOR: #58ACFA" ><center>歷年統計</center></td>
				    </tr>	
				    <tr height="60px">
					    <td align="left" colspan="2">
						    ◎累計奉獻100萬以上之天使筆數：<font color="blue"><asp:Label ID="lblOver100_Count" runat="server"></asp:Label></font>
						    <br/>◎累計奉獻50萬以上之天使筆數：<font color="blue"><asp:Label ID="lblOver50_Count" runat="server"></asp:Label></font>
						    <br/>◎累計奉獻30萬以上之天使筆數：<font color="blue"><asp:Label ID="lblOver30_Count" runat="server"></asp:Label></font>
						    <br/>◎累計奉獻10萬以上之天使筆數：<font color="blue"><asp:Label ID="lblOver10_Count" runat="server"></asp:Label></font>
					    </td>    		
				    </tr>				
			    </table>
		    </td>
            <td align="left">
            <table border="0" cellspacing="0" cellpadding="0" style="background-color:#FFFFFF" width="100%">
			 <tr>
				<td align="center"><asp:Label ID="lblyear1" runat="server"></asp:Label>年依台灣地區(縣市)別捐款比例</td>
				<td align="center"><asp:Label ID="lblyear2" runat="server"></asp:Label>年依捐款方式別捐款比例</td>
			 </tr>
			 <tr>
				<td align="center"><iframe src="../ChartDirector/chart1.aspx" width="530" height="400" scrolling="no" align="center" frameborder="0"></iframe></td>
				<td align="center"><iframe src="../ChartDirector/chart2.aspx" width="530" height="400" scrolling="no" align="center" frameborder="0"></iframe></td>
			 </tr>
			 <tr><td colspan="2">　</td></tr>
			 <tr>
				<td align="center"><asp:Label ID="lblyear3" runat="server"></asp:Label>年7月起線上奉獻金額與筆數</td>
				<td align="center">各年度累計總奉獻金額</td>
			 </tr>
			 <tr>
				<td align="center"><iframe src="../ChartDirector/chart3.aspx" width="530" height="400" scrolling="no" align="center" frameborder="0"></iframe></td>
				<td align="center"><iframe src="../ChartDirector/chart4.aspx" width="530" height="400" scrolling="no" align="center" frameborder="0"></iframe></td>
			 </tr>
             <tr>
                <td align="center" colspan="2">依捐款方式統計捐款金額與總筆數</td>
             </tr>
             <tr>
                <td align="center" colspan="2"><iframe src="../ChartDirector/chart5.aspx" width="530" height="500" scrolling="no" align="center" frameborder="0"></td>
             </tr>
			</table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>

