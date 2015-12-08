<!--#include file="../include/dbfunctionJ.asp"-->
<%
year_now=year(Date)
'year_now= 2012

Set cd = CreateObject("ChartDirector.API")

SQL = "SELECT Donate_Payment ,Sum(Donate_Amt) AS Amt FROM dbo.DONATE " & _
       "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
	   "AND DATEPART(YEAR,Donate_Date) = '" & year_now & "' GROUP BY Donate_Payment ORDER by Sum(Donate_Amt) desc"
	   
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open SQL,Conn,1,1

' Create a PieChart object of size 600 x 320 pixels. Use a vertical gradient color
' from light blue (99ccff) to white (ffffff) spanning the top 100 pixels as
' background. Set border to grey (888888). Use rounded corners. Enable soft drop
' shadow.
Set c = cd.PieChart(450, 270)
Call c.setBackground(c.linearGradientColor(0, 0, 0, 100, &H99ccff, &Hffffff), _
    &H888888)
	
If rs.RecordCount = 0 then
	Call c.addTitle("No data now!", "timesbi.ttf", 15)
Else
	Set dbTable = cd.DBTable(rs)
	' The data for the pie chart
	data = dbTable.getCol(1)

	' The labels for the pie chart
	labels = dbTable.getCol(0)
	' The colors to use for the sectors
	colors = Array(&H66aaee, &Heebb22, &Hbbbbbb, &H8844ff, &Hdd2222, &H009900)

	Call c.setRoundedFrame()
	Call c.setDropShadow()
	Call c.setDefaultFonts("simhei.ttf") 
	' Add a title using 18 pts Times New Roman Bold Italic font. Add 16 pixels top margin
	' to the title.
	'Call c.addTitle("Pie Chart With Legend Demonstration", "timesbi.ttf", 18 _
		').setMargin2(0, 0, 16, 0)

	' Set the center of the pie at (160, 165) and the radius to 110 pixels
	Call c.setPieSize(120, 140, 80)

	' Draw the pie in 3D with a pie thickness of 25 pixels
	Call c.set3D(15)

	' Set the pie data and the pie labels
	Call c.setData(data, labels)

	' Set the sector colors
	Call c.setColors2(cd.DataColor, colors)

	' Use local gradient shading for the sectors
	Call c.setSectorStyle(cd.LocalGradientShading)

	' Use the side label layout method, with the labels positioned 16 pixels from the pie
	' bounding box
	Call c.setLabelLayout(cd.SideLayout, 16)

	' Show only the sector number as the sector label
	Call c.setLabelFormat("{={sector}+1}")

	' Set the sector label style to Arial Bold 10pt, with a dark grey (444444) border
	Call c.setLabelStyle("arialbd.ttf", 9).setBackground(cd.Transparent, &H444444)
		' Add a legend box, with the center of the left side anchored at (330, 175), and
	' using 10 pts Arial Bold Italic font
	Set b = c.addLegend(240, 135, True, "simhei.ttf", 9)
	Call b.setAlignment(cd.Left)

	' Set the legend box border to dark grey (444444), and with rounded conerns
	Call b.setBackground(cd.Transparent, &H444444)
	Call b.setRoundedCorners()

	' Set the legend box margin to 16 pixels, and the extra line spacing between the
	' legend entries as 5 pixels
	Call b.setMargin(14)
	Call b.setKeySpacing(0, 4)

	' Set the legend box icon to have no border (border color same as fill color)
	Call b.setKeyBorder(cd.SameAsMainColor)

	' Set the legend text to show the sector number, followed by a 120 pixels wide block
	' showing the sector label, and a 40 pixels wide block showing the percentage
	Call b.setText( _
		"<*block,valign=top*>{={sector}+1}.<*advanceTo=20*><*block,width=110*>" & _
		"{label}<*/*><*block,width=30,halign=right*>{percent}<*/*>%")
End If

' Output the chart
Response.ContentType = "image/png"
Response.BinaryWrite c.makeChart2(cd.PNG)
Response.End

%>