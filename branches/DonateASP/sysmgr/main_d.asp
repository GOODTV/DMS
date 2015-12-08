<!--#include file="../include/dbfunctionJ.asp"-->
<!--#include file="../include/md5.asp"-->
<%
year_now=year(Date)
month_now=month(Date)
%>
<%
  SQL="Select Style_Name,Style_Value From WEB_STYLE Where Style_Type='MD5' And Style_Name In ('KEY_6','KEY_7','KEY_8') Order By Style_Seq"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then
    While Not RS.EOF
      If RS("Style_Name")="KEY_6" And RS("Style_Value")<>"" Then
        For I=2009 To 2099
		      If MD5(Cstr(I))=Cstr(RS("Style_Value")) Then 
		        ExpireDate=Cstr(I)
		        Exit For
		      End If
		    Next
		  End If
		  If RS("Style_Name")="KEY_7" And RS("Style_Value")<>"" Then
		    For I=1 To 12
          If MD5(Cstr(I))=Cstr(RS("Style_Value")) Then 
            ExpireDate=ExpireDate&"/"&Cstr(I)
		      End If
		    Next
		  End If
		  If RS("Style_Name")="KEY_8" And RS("Style_Value")<>"" Then
		    For I=1 To 31
          If MD5(Cstr(I))=Cstr(RS("Style_Value")) Then 
            ExpireDate=ExpireDate&"/"&Cstr(I)
            Exit For
          End If
		    Next
		  End If
		  RS.MoveNext
		Wend
		If ExpireDate<>"" Then
		  If CDate(Date())>CDate(DateAdd("D",-30,CDate(ExpireDate))) Then
		    Session("errnumber")=1
        Session("msg")="親愛的管理者您好！\n\n貴會網站代管期將於  "&ExpireDate&"  到期\n\n請您盡速與網軟公司聯絡辦理續約"
		  End If
		End If		  
  End If
  
  'Add by GoodTV Tanya:新增奉獻紀錄統計資訊
  'Modify by GoodTV Tanay:Dash board新增日期及奉獻平均金額
  '日期
  Date_Today = Cstr(Date())
  This_YearMonth =Cstr(Year(Date())) &"/"& Cstr(Month(Date()))
  This_Year =Cstr(Year(Date()))
  Last_Year =Cstr(Year(Date())-1)
  Last_YearMonth=Cstr(Year(Date())-1) &"/"& Cstr(Month(Date()))
  '天使總人數
  InfoSQL="SELECT Replace(Convert(Varchar,Convert(money,COUNT(*)),1),'.00','') Angel_Count FROM dbo.DONOR WHERE IsMember <> 'Y'"
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Angel_Count = InfoRS("Angel_Count")
  InfoRS.Close
  Set InfoRS=Nothing
  '讀者總人數
  InfoSQL="SELECT Replace(Convert(Varchar,Convert(money,COUNT(*)),1),'.00','') Reader_Count FROM dbo.DONOR WHERE IsMember = 'Y'"
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Reader_Count = InfoRS("Reader_Count")
  InfoRS.Close
  Set InfoRS=Nothing
  '今日累計奉獻金額,筆數,平均
  SelectSQL="SELECT Replace(Convert(Varchar,A.Count,1),'.00','') Count " & _
          "       ,Replace(Convert(Varchar,A.SumAmt,1),'.00','') SumAmt " & _   
          "				,Replace(Convert(Varchar,Case When A.Count > 2 Then Round((A.SumAmt - A.Max_Amt - A.Min_Amt) / (A.Count - 2),0) " & _				
          " 																		Else 0 End,1),'.00','') DonateAmt_Avg " & _
  				"FROM (SELECT Convert(money,COUNT(*)) Count ,IsNull(SUM(Donate_Amt),0) SumAmt " & _
          "			 				,IsNull(Max(Donate_Amt),0) Max_Amt,IsNull(Min(Donate_Amt),0) Min_Amt " & _
          "			 FROM dbo.DONATE " 
  WhereSQL=" WHERE Convert(Varchar(10),Donate_Date,111) = Convert(Varchar(10),GETDATE(),111) and Issue_Type <> 'D' " 
  InfoSQL=SelectSQL&WhereSQL&" ) A"
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Today_Count = InfoRS("Count")
  Today_SumAmt = InfoRS("SumAmt")
  Today_DonateAmt_Avg = InfoRS("DonateAmt_Avg")
  InfoRS.Close
  Set InfoRS=Nothing
  '今日線上累計奉獻筆數
  InfoSQL="select (select COUNT(*) as online_count from DONATE_IEPAY P left join DONATE_Web W on P.orderid=W.od_sob where status = '0' " & _
          "  and Convert(Varchar(10),Donate_CreateDate,111)= Convert(Varchar(10),GETDATE(),111))+ " & _   
          "	(select COUNT(*) as pledge_count from PLEDGE_Temp " & _				
          "	 where  Convert(Varchar(10),Create_Date,111) = Convert(Varchar(10),GETDATE(),111)) as count " 
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Today_Online_Count = InfoRS("count")
  InfoRS.Close
  Set InfoRS=Nothing
  '本月累計奉獻金額,筆數,平均
  WhereSQL=" WHERE Year(Donate_Date) = Year(GETDATE()) AND MONTH(Donate_Date) = MONTH(GETDATE()) and Issue_Type <> 'D' "
  InfoSQL=SelectSQL&WhereSQL&" ) A"         
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Month_Count = InfoRS("Count")
  Month_SumAmt = InfoRS("SumAmt")   
  Month_DonateAmt_Avg = InfoRS("DonateAmt_Avg")
  InfoRS.Close
  Set InfoRS=Nothing
  '本月線上累計奉獻筆數
  InfoSQL="select (select COUNT(*) as online_count from DONATE_IEPAY P left join DONATE_Web W on P.orderid=W.od_sob where status = '0' " & _
          "  and Year(Donate_CreateDate) = Year(GETDATE()) AND MONTH(Donate_CreateDate) = MONTH(GETDATE()))+ " & _   
          "	(select COUNT(*) as pledge_count from PLEDGE_Temp " & _				
          "	 where  Year(Create_Date) = Year(GETDATE()) AND MONTH(Create_Date) = MONTH(GETDATE())) as count " 
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Month_Online_Count = InfoRS("count")
  InfoRS.Close
  Set InfoRS=Nothing
  '本年度累計奉獻金額,筆數,平均
  WhereSQL=" WHERE Year(Donate_Date) = Year(GETDATE()) and Issue_Type <> 'D' "
  InfoSQL=SelectSQL&WhereSQL&" ) A"         
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Year_Count = InfoRS("Count")
  Year_SumAmt = InfoRS("SumAmt")
  Year_DonateAmt_Avg = InfoRS("DonateAmt_Avg")
  InfoRS.Close
  Set InfoRS=Nothing
  '本年度線上累計奉獻筆數
  InfoSQL="select (select COUNT(*) as online_count from DONATE_IEPAY P left join DONATE_Web W on P.orderid=W.od_sob where status = '0' " & _
          "  and Year(Donate_CreateDate) = Year(GETDATE()))+ " & _   
          "	(select COUNT(*) as pledge_count from PLEDGE_Temp " & _				
          "	 where  Year(Create_Date) = Year(GETDATE())) as count " 
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Year_Online_Count = InfoRS("count")
  InfoRS.Close
  Set InfoRS=Nothing
  '上一年度奉獻總金額,筆數
  WhereSQL=" WHERE Year(Donate_Date) = Year(GETDATE())-1 and Issue_Type <> 'D' "
  InfoSQL=SelectSQL&WhereSQL&" ) A"          
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  LastYear_Count = InfoRS("Count")
  LastYear_SumAmt = InfoRS("SumAmt")
  LastYear_DonateAmt_Avg = InfoRS("DonateAmt_Avg")
  InfoRS.Close
  Set InfoRS=Nothing
  '上一年度本月奉獻總金額,筆數
  WhereSQL=" WHERE Year(Donate_Date) = Year(GETDATE())-1 AND MONTH(Donate_Date) = MONTH(GETDATE()) and Issue_Type <> 'D' "  
  InfoSQL=SelectSQL&WhereSQL&" ) A"        
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  LastYearMonth_Count = InfoRS("Count")
  LastYearMonth_SumAmt = InfoRS("SumAmt")
  LastYearMonth_DonateAmt_Avg = InfoRS("DonateAmt_Avg")
  InfoRS.Close
  Set InfoRS=Nothing
  '歷年累計奉獻100萬以上天使筆數
  InfoSQL="SELECT Replace(Convert(Varchar,Convert(money,COUNT(Donor_Id)),1),'.00','') Over100_Count " & _
          "FROM dbo.DONOR WHERE Donor_Id IN " & _
          "(SELECT Donor_Id FROM dbo.DONATE WHERE Donate_Amt > 0 GROUP BY Donor_Id HAVING SUM(Donate_Amt) >= 1000000)"         
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Over100_Count = InfoRS("Over100_Count")
  InfoRS.Close
  Set InfoRS=Nothing
  '歷年累計奉獻50萬以上天使筆數
  InfoSQL="SELECT Replace(Convert(Varchar,Convert(money,COUNT(Donor_Id)),1),'.00','') Over50_Count " & _
          "FROM dbo.DONOR WHERE Donor_Id IN " & _
          "(SELECT Donor_Id FROM dbo.DONATE WHERE Donate_Amt > 0 GROUP BY Donor_Id HAVING SUM(Donate_Amt) >= 500000)"         
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Over50_Count = InfoRS("Over50_Count")
  InfoRS.Close
  Set InfoRS=Nothing
  '歷年累計奉獻30萬以上天使筆數
  InfoSQL="SELECT Replace(Convert(Varchar,Convert(money,COUNT(Donor_Id)),1),'.00','') Over30_Count " & _
          "FROM dbo.DONOR WHERE Donor_Id IN " & _
          "(SELECT Donor_Id FROM dbo.DONATE WHERE Donate_Amt > 0 GROUP BY Donor_Id HAVING SUM(Donate_Amt) >= 300000)"         
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Over30_Count = InfoRS("Over30_Count")
  InfoRS.Close
  Set InfoRS=Nothing
  '歷年累計奉獻10萬以上天使筆數
  InfoSQL="SELECT Replace(Convert(Varchar,Convert(money,COUNT(Donor_Id)),1),'.00','') Over10_Count " & _
          "FROM dbo.DONOR WHERE Donor_Id IN " & _
          "(SELECT Donor_Id FROM dbo.DONATE WHERE Donate_Amt > 0 GROUP BY Donor_Id HAVING SUM(Donate_Amt) >= 100000)"         
  Set InfoRS = Server.CreateObject("ADODB.RecordSet")
  InfoRS.Open InfoSQL,Conn,1,1
  Over10_Count = InfoRS("Over10_Count")
  InfoRS.Close
  Set InfoRS=Nothing
  Response.write "<script language='JavaScript'>parent.left.location.reload();</script>"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link href="../include/dms.css" type="text/css" rel="stylesheet">	
  <title><%=session("sys_name")%></title>
</head>
<body class="gray" onselectstart="return false;" ondragstart="return false;" oncontextmenu="return false;">
  <div align="center"><center><br>
    <table border="0" cellspacing="0" cellpadding="0" style="BACKGROUND-COLOR: #EEEEE3">
      <!--tr>
        <td align="center" colspan="2"><br><br>您已經以管理員身分登入，請點選左方之管理功能進行網站管理與維護。<br><br>在每一階段的管理作業結束後，也請您記得利用上方的“管理登出”選項，登出您的管理員帳號。</td>
      </tr-->
	  <tr>
		<td valign="top">
			<table border="1" width="210" cellspacing="0" cellpadding="0"  style="BACKGROUND-COLOR: #EEEEE3">
				<tr height="30px" style="BACKGROUND-COLOR: #F7FE2E">
					<td align="center"><font color="green" size="3"><b>:::捐款資訊 Dash board:::</b></font></td>    		
				</tr>
				<tr height="30px"> 
					<td align="left">◎天使總人數：<font color="blue"><%=Angel_Count%></font><br/>◎讀者總人數：<font color="blue"><%=Reader_Count%></font></td>
				</tr>				
				<tr height="25px">
					<td align="left" style="BACKGROUND-COLOR: #82FA58" ><center>今日 <%=Date_Today%></center></td>
				</tr>		
				<tr height="50px">
					<td align="left">◎累計奉獻筆數：<font color="blue"><%=Today_Count%></font><br/>◎累計奉獻筆數(線上)：<font color="blue"><%=Today_Online_Count%></font><br/>◎累計奉獻金額：<font color="blue"><%=Today_SumAmt%></font><br/>◎平均奉獻金額：<font color="blue"><%=Today_DonateAmt_Avg%></font></td>
				</tr>		
				<tr height="25px">
					<td align="left" style="BACKGROUND-COLOR: #82FA58" ><center>本月 <%=This_YearMonth%></center></td>
				</tr>	
				<tr height="50px">
					<td align="left">◎累計奉獻筆數：<font color="blue"><%=Month_Count%></font><br/>◎累計奉獻筆數(線上)：<font color="blue"><%=Month_Online_Count%></font><br/>◎累計奉獻金額：<font color="blue"><%=Month_SumAmt%></font><br/>◎平均奉獻金額：<font color="blue"><%=Month_DonateAmt_Avg%></font></td>
				</tr>		
				<tr height="25px">
					<td align="left" style="BACKGROUND-COLOR: #82FA58" ><center>本年度 <%=This_Year%></center></td>
				</tr>			
				<tr height="50px">
					<td align="left">◎累計奉獻筆數：<font color="blue"><%=Year_Count%></font><br/>◎累計奉獻筆數(線上)：<font color="blue"><%=Year_Online_Count%></font><br/>◎累計奉獻金額：<font color="blue"><%=Year_SumAmt%></font><br/>◎平均奉獻金額：<font color="blue"><%=Year_DonateAmt_Avg%></font></td>
				</tr>				
				
				<tr height="25px">
					<td align="left" style="BACKGROUND-COLOR: #F781F3" ><center>上一年度 <%=Last_Year%></center></td>
				</tr>	
				<tr height="50px">
					<td align="left">◎奉獻總筆數：<font color="blue"><%=LastYear_Count%></font><br/>◎奉獻總金額：<font color="blue"><%=LastYear_SumAmt%></font><br/>◎奉獻平均金額：<font color="blue"><%=LastYear_DonateAmt_Avg%></font></td>
				</tr>		
				<tr height="25px">
					<td align="left" style="BACKGROUND-COLOR: #F781F3" ><center>上一年度本月 <%=Last_YearMonth%></></center></td>
				</tr>		
				<tr height="50px">
					<td align="left">◎奉獻總筆數：<font color="blue"><%=LastYearMonth_Count%></font><br/>◎奉獻總金額：<font color="blue"><%=LastYearMonth_SumAmt%></font><br/>◎奉獻平均金額：<font color="blue"><%=LastYearMonth_DonateAmt_Avg%></font></td>
				</tr>				
				
				<tr height="25px">
					<td align="left" style="BACKGROUND-COLOR: #58ACFA" ><center>歷年統計</center></td>
				</tr>	
				<tr height="60px">
					<td align="left" colspan="2">
						◎累計奉獻100萬以上之天使筆數：<font color="blue"><%=Over100_Count%></font>
						<br/>◎累計奉獻50萬以上之天使筆數：<font color="blue"><%=Over50_Count%></font>
						<br/>◎累計奉獻30萬以上之天使筆數：<font color="blue"><%=Over30_Count%></font>
						<br/>◎累計奉獻10萬以上之天使筆數：<font color="blue"><%=Over10_Count%></font>
					</td>    		
				</tr>				
			</table>
		</td>
		<td>
			<table border="0" cellspacing="0" cellpadding="0" style="background-color:#FFFFFF">
			 <tr>
				<td align="center"><%=year_now%>年依台灣地區(縣市)別捐款比例</td>
				<td align="center"><%=year_now%>年依捐款方式別捐款比例</td>
			 </tr>
			 <tr>
				<td align="center"><img src="chart1.asp"></td>
				<td align="center"><img src="chart2.asp"></td>
			 </tr>
			 <tr><td colspan="2">　</td></tr>
			 <tr>
				<td align="center"><%=year_now%>年7月起線上奉獻金額與筆數</td>
				<td align="center">各年度累計總奉獻金額</td>
			 </tr>
			 <tr>
				<td align="center"><img src="chart7.asp"></td>
				<td align="center"><img src="chart4.asp"></td>
			 </tr>
			 <tr><td colspan="2">　</td></tr>
			 <tr>
				<td colspan="2" align="center">依捐款方式統計捐款金額與總筆數</td>
			 </tr>
			 <tr>
				<td colspan="2" align="center"><iframe src="chart5_data.asp" width="550" height="500" scrolling="no" align="center" frameborder="0"></iframe></td>
			 </tr>
			</table>
		</td>
	  </tr>
    </table>
    <p />
    
	
  </center></div>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->