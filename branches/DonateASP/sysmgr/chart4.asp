<!--#include file="../include/dbfunctionJ.asp"-->
<%
Set cd = CreateObject("ChartDirector.API")

SQL= "SELECT top 3 DATEPART(YEAR,Donate_Date) AS '年度',ISNULL(Sum(Donate_Amt)/100000000,0) AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	 "group by DATEPART(YEAR,Donate_Date) " & _
	 "union " & _
	 "SELECT DATEPART(YEAR,Donate_Date) AS '年度',ISNULL(Sum(Donate_Amt)/100000000,0) AS Sum_Amt FROM dbo.DONATE " & _
     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 and DATEPART(YEAR,Donate_Date) between 2007 and DATEPART(YEAR,GETDATE()) " & _
	 "group by DATEPART(YEAR,Donate_Date)"
	 
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open SQL,Conn,1,1
Set dbTable = cd.DBTable(rs)

' The data for the pie chart
data = dbTable.getCol(1)
'data = Array(450, 560, 630, 800, 1100, 1350, 1600, 1950, 2300, 2700)

' The labels for the pie chart
labels = dbTable.getCol(0)

' Create a XYChart object of size 250 x 250 pixels
Set c = cd.XYChart(480, 270)
'Call c.setDefaultFonts("simhei.ttf", "kaiu.ttf") 
' Set the plotarea at (30, 20) and of size 200 x 200 pixels
Call c.setPlotArea(60, 30, 400, 200)

' Add a bar chart layer using the given data
'Call c.addBarLayer(data, &Hff00).set3D()

Set layer = c.addBarLayer3(data)
Call layer.set3D(10)
' Set bar shape to circular (cylinder)
Call layer.setBarShape(cd.CircleShape)

' Add a title to the y axis
Call c.yAxis().setTitle("總捐款金額(億/單位)","kaiu.ttf Bold",12)

' Set the labels on the x axis.
Call c.xAxis().setLabels(labels)

' Output the chart
Response.ContentType = "image/png"
Response.BinaryWrite c.makeChart2(cd.PNG)
Response.End
%>