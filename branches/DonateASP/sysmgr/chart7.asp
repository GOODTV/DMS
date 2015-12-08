<!--#include file="../include/dbfunctionJ.asp"-->

<%
year_now=year(Date)
'year_now= 2013

Set cd = CreateObject("ChartDirector.API")

SQL1= "select Convert(varchar,DATEPART(Month,Donate_CreateDate))+'月' as Month,COUNT(orderid) AS '筆數' ,Replace(Convert(Varchar,Sum(Donate_Amount),1),'.00','') AS Sum_Amt " & _
      "from DONATE_IEPAY P left join DONATE_Web W on P.orderid=W.od_sob " & _
	  "where status = '0' " & _
	  "AND DATEPART(YEAR,Donate_CreateDate) = " & year_now & " and DATEPART(Month,Donate_CreateDate) between 7 and DATEPART(Month,GETDATE()) GROUP by DATEPART(Month,Donate_CreateDate)"

Set rs1 = Server.CreateObject("ADODB.RecordSet")

rs1.Open SQL1,Conn,1,1

' Create a XYChart object of size 300 x 180 pixels
Set c = cd.XYChart(470, 300)
'Call c.setDefaultFonts("mingliu.ttc", "mingliu.ttc Bold") 

If rs1.RecordCount = 0 then
	Call c.addTitle("No data now!", "timesbi.ttf", 15)
	
Else
	' Set the plot area at (50, 20) and of size 200 x 130 pixels
	Call c.setPlotArea(45, 10, 350, 260)
	
	Set dbTable = cd.DBTable(rs1)
	' The data for the line chart
	data1 = dbTable.getCol(1)
	data2 = dbTable.getCol(2)
	
	' The labels for the line chart
	labels = dbTable.getCol(0)
	'labels = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
	'labels = Array("1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月")
	
	' Add a title to the primary (left) y axis
	Call c.yAxis().setTitle("總筆數","kaiu.ttf Bold",12)
	' Add a title to the secondary (right) y axis
	Call c.yAxis2().setTitle("總捐款金額","kaiu.ttf Bold",12)
End If

' Add a title to the chart using 8 pts Arial Bold font
'Call c.addTitle2(8,"奉獻通路奉獻金額及筆數總計", "kaiu.ttf Bold", 12)

' Set the labels on the x axis.
Call c.xAxis().setLabels(labels)
Call c.xAxis().setLabelStyle("kaiu.ttf Bold",9,&H000000,0)

' Set the axis, label and title colors for the primary y axis to red (0xc00000) to
' match the first data set
Call c.yAxis().setColors(&HFF0000, &HFF0000, &HFF0000)

' set the axis, label and title colors for the primary y axis to green (0x008000) to
' match the second data set
Call c.yAxis2().setColors(&H4169E1, &H4169E1, &H4169E1)

' Add a line layer to for the first data set using red (0xc00000) color with a line
' width to 3 pixels
Call c.addLineLayer(data1, &HFF0000).setLineWidth(3)

' Add a bar layer to for the second data set using green (0x00C000) color. Bind the
' second data set to the secondary (right) y axis
Call c.addBarLayer(data2, &H4169E1).setUseYAxis2()

' Output the chart
Response.ContentType = "image/png"
Response.BinaryWrite c.makeChart2(cd.PNG)
Response.End
%>
