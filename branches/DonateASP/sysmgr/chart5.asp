<!--#include file="../include/dbfunctionJ.asp"-->

<%
year_now=year(Date)
'year_now= 2013
'response.write year_beforelast

Set cd = CreateObject("ChartDirector.API")
'
' Displays the monthly revenue for the selected year. The selected year should be
' passed in as a query parameter called "year"
'
selectedYear = Request("year")
if selectedYear = "" Then selectedYear = year_now

SQL1= "SELECT CONVERT(varchar,Donate_Payment),COUNT(Donate_Id)/1000 AS '筆數',Sum(Donate_Amt)/10000 AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & selectedYear & " GROUP by Donate_Payment"

Set rs1 = Server.CreateObject("ADODB.RecordSet")

rs1.Open SQL1,Conn,1,1

' Create a XYChart object of size 300 x 180 pixels
Set c = cd.XYChart(550, 440)
'Call c.setDefaultFonts("mingliu.ttc", "mingliu.ttc Bold") 

If rs1.RecordCount = 0 then
	Call c.addTitle("No data now!", "timesbi.ttf", 15)
	
Else
	' Set the plot area at (50, 20) and of size 200 x 130 pixels
	Call c.setPlotArea(55, 10, 430, 300)
	
	Set dbTable = cd.DBTable(rs1)
	' The data for the line chart
	data1 = dbTable.getCol(1)
	data2 = dbTable.getCol(2)
	
	' The labels for the line chart
	labels = dbTable.getCol(0)
	'labels = Array("ATM", "上海銀行(轉帳)", "支票", "信用卡", "信用卡授權書", "信用卡授權書(一般)", "美國運通", "現金", "郵局轉帳", "匯款", "劃撥", "彰銀(轉帳)", "網路信用卡")
	
	' Add a title to the primary (left) y axis
	Call c.yAxis().setTitle("總捐款金額 (單位：萬)","kaiu.ttf Bold",12)
	' Add a title to the secondary (right) y axis
	Call c.yAxis2().setTitle("總筆數 (單位：千)","kaiu.ttf Bold",12)
End If

' Add a title to the chart using 8 pts Arial Bold font
'Call c.addTitle2(8,"奉獻通路奉獻金額及筆數總計", "kaiu.ttf Bold", 12)

' Set the labels on the x axis.
Call c.xAxis().setLabels(labels)
Call c.xAxis().setLabelStyle("kaiu.ttf Bold",9,&H000000,90)

' Set the axis, label and title colors for the primary y axis to red (0xc00000) to
' match the first data set
Call c.yAxis().setColors(&H696969, &H696969, &H696969)

' set the axis, label and title colors for the primary y axis to green (0x008000) to
' match the second data set
Call c.yAxis2().setColors(&HFF1493, &HFF1493, &HFF1493)

' Add a line layer to for the first data set using red (0xc00000) color with a line
' width to 3 pixels
Call c.addLineLayer(data1, &HFF1493).setUseYAxis2()

' Add a bar layer to for the second data set using green (0x00C000) color. Bind the
' second data set to the secondary (right) y axis
Call c.addBarLayer(data2, &H696969).setLineWidth(3)

' Output the chart
Response.ContentType = "image/png"
Response.BinaryWrite c.makeChart2(cd.PNG)
Response.End
%>
