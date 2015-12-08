<!--#include file="../include/dbfunctionJ.asp"-->
<%
year_now=year(Date)
year_last=year(DateAdd("yyyy",-1,Now()))
year_beforelast=year(DateAdd("yyyy",-2,Now()))
'year_now= 2012
'response.write year_beforelast

Set cd = CreateObject("ChartDirector.API")

SQL1= "SELECT '01' AS '月份',Case when DATEPART(Month,Donate_Date)= '1' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date) " & _
	 "Union " & _
	 "SELECT '02' AS '月份',Case when DATEPART(Month,Donate_Date)= '2' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date)" & _
	 "Union " & _
	 "SELECT '03' AS '月份',Case when DATEPART(Month,Donate_Date)= '3' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date)" & _
	 "Union " & _
	 "SELECT '04' AS '月份',Case when DATEPART(Month,Donate_Date)= '4' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date)" & _
	 "Union " & _
	 "SELECT '05' AS '月份',Case when DATEPART(Month,Donate_Date)= '5' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date)" & _
	 "Union " & _
	 "SELECT '06' AS '月份',Case when DATEPART(Month,Donate_Date)= '6' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date)" & _
	 "Union " & _
	 "SELECT '07' AS '月份',Case when DATEPART(Month,Donate_Date)= '7' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date)" & _
	 "Union " & _
	 "SELECT '08' AS '月份',Case when DATEPART(Month,Donate_Date)= '8' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date)" & _
	 "Union " & _
	 "SELECT '09' AS '月份',Case when DATEPART(Month,Donate_Date)= '9' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date)" & _
	 "Union " & _
	 "SELECT '10' AS '月份',Case when DATEPART(Month,Donate_Date)= '10' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date)" & _
	 "Union " & _
	 "SELECT '11' AS '月份',Case when DATEPART(Month,Donate_Date)= '11' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date)" & _
	 "Union " & _
	 "SELECT '12' AS '月份',Case when DATEPART(Month,Donate_Date)= '12' Then Sum(Donate_Amt)/10000 Else 0 End AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_now & " GROUP by DATEPART(month,Donate_Date)"
SQL2= "SELECT Convert(varchar,DATEPART(month,Donate_Date)) AS 'Month',Sum(Donate_Amt)/10000 AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_last & " GROUP by DATEPART(month,Donate_Date)"
SQL3= "SELECT Convert(varchar,DATEPART(month,Donate_Date)) AS 'Month',Sum(Donate_Amt)/10000 AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "AND DATEPART(YEAR,Donate_Date) = " & year_beforelast & " GROUP by DATEPART(month,Donate_Date)"
'response.write SQL1
'response.end
' Create an XYChart object of size 600 x 300 pixels, with a light blue (EEEEFF)
' background, black border, 1 pxiel 3D border effect and rounded corners
Set c = cd.XYChart(450, 240, &Heeeeff, &H000000, 1)

Set rs1 = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set rs3 = Server.CreateObject("ADODB.RecordSet")
rs1.Open SQL1,Conn,1,1
rs2.Open SQL2,Conn,1,1
rs3.Open SQL3,Conn,1,1

Set dbTable1 = cd.DBTable(rs1)
Set dbTable2 = cd.DBTable(rs2)
Set dbTable3 = cd.DBTable(rs3)

' The data for the line chart
data1 = dbTable1.getCol(1)
data2 = dbTable2.getCol(1)
data3 = dbTable3.getCol(1)

' The labels for the line chart
'labels = dbTable.getCol(0)
labels = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

Call c.setRoundedFrame()

' Set the plotarea at (55, 58) and of size 520 x 195 pixels, with white background.
' Turn on both horizontal and vertical grid lines with light grey color (0xcccccc)
Call c.setPlotArea(55, 28, 380, 180, &Hffffff, -1, -1, &Hcccccc, &Hcccccc)

' Add a legend box at (50, 30) (top of the chart) with horizontal layout. Use 9 pts
' Arial Bold font. Set the background and border color to Transparent.
Call c.addLegend(50, 0, False, "arialbd.ttf", 9).setBackground(cd.Transparent)

' Add a title to the y axis
Call c.yAxis().setTitle("NT$10,000 per Unit")

' Set the labels on the x axis.
Call c.xAxis().setLabels(labels)
' Add a line layer to the chart
Set layer = c.addLineLayer2()

' Set the default line width to 2 pixels
Call layer.setLineWidth(2)

' Add the three data sets to the line layer. For demo purpose, we use a dash line
' color for the last line
Call layer.addDataSet(data1, &Hff0000, ""&year_now&"")
Call layer.addDataSet(data2, &H008800, ""&year_last&"")
Call layer.addDataSet(data3, &H3333ff, ""&year_beforelast&"")

' Output the chart
Response.ContentType = "image/png"
Response.BinaryWrite c.makeChart2(cd.PNG)
Response.End
%>
