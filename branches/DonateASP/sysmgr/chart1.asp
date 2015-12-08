<!--#include file="../include/dbfunctionJ.asp"-->
<%
year_now=year(Date)
'year_now= 2012

Set cd = CreateObject("ChartDirector.API")

'SQL= "SELECT Donate_Purpose ,Sum(Donate_Amt) AS Amt FROM dbo.DONATE " & _
'     "WHERE Donate_Purpose <> '' AND Donate_Amt > 0 " & _
'	 "AND DATEPART(YEAR,Donate_Date) = '" & year_now & "' GROUP BY Donate_Purpose ORDER by Sum(Donate_Amt) desc"

SQL= "SELECT CASE City " & _
	 "when 'A' then '台北市' when 'B' then '新北市' when 'C' then '基隆市' when 'D' then '宜蘭縣' when 'E' then '桃園縣' " & _
	 "when 'F' then '新竹市' when 'G' then '新竹縣' when 'H' then '苗栗縣' when 'I' then '台中市' when 'K' then '彰化縣' " & _
	 "when 'L' then '南投縣' when 'M' then '雲林縣' when 'N' then '嘉義市' when 'O' then '嘉義縣' when 'P' then '台南市' " & _
	 "when 'R' then '高雄市' when 'T' then '屏東縣' when 'U' then '台東縣' when 'V' then '花蓮縣' when 'W' then '澎湖縣' " & _
	 "when 'X' then '金門縣' when 'Y' then '連江縣' when '' then '不具名' else '未填縣市' end as City ,Sum(Donate_Amt) AS Sum_Amt " & _
	 "From DONATE D left join DONOR N on D.Donor_Id=N.Donor_Id " & _
     "WHERE Donate_Amt > 0 and Issue_Type <> 'D' and IsAbroad = 'N' " & _
	 "AND DATEPART(YEAR,Donate_Date) = '" & year_now & "' GROUP BY City ORDER by Sum(Donate_Amt) desc"
	  
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open SQL,Conn,1,1

' Create a PieChart object of size 560 x 270 pixels, with a golden background and a 1
' pixel 3D border
Set c = cd.PieChart(440, 350, cd.goldColor(), -1, 0)

If rs.RecordCount = 0 then
	Call c.addTitle("No data now!", "timesbi.ttf", 15)
Else
	Set dbTable = cd.DBTable(rs)
	' The data for the pie chart
	data = dbTable.getCol(1)

	' The labels for the pie chart
	labels = dbTable.getCol(0)
End If

' Add a title box using 15 pts Times Bold Italic font and metallic pink background
' color
'Call c.addTitle("Project Cost Breakdown", "timesbi.ttf", 15).setBackground( _
    'cd.metalColor(&Hff9999))
Call c.setDefaultFonts("simhei.ttf") 

' Set the center of the pie at (280, 135) and the radius to 110 pixels
Call c.setPieSize(215, 140, 80)

' Draw the pie in 3D with 20 pixels 3D depth
Call c.set3D(15)

' Use the side label layout method
Call c.setLabelLayout(cd.SideLayout)

' Set the label box background color the same as the sector color, with glass effect,
' and with 5 pixels rounded corners
Set t = c.setLabelStyle()
Call t.setBackground(cd.SameAsMainColor, cd.Transparent, cd.glassEffect())
Call t.setRoundedCorners(5)

' Set the border color of the sector the same color as the fill color. Set the line
' color of the join line to black (0x0)
Call c.setLineColor(cd.SameAsMainColor, &H000000)

' Set the start angle to 135 degrees may improve layout when there are many small
' sectors at the end of the data array (that is, data sorted in descending order). It
' is because this makes the small sectors position near the horizontal axis, where
' the text label has the least tendency to overlap. For data sorted in ascending
' order, a start angle of 45 degrees can be used instead.
Call c.setStartAngle(135)

' Set the pie data and the pie labels
Call c.setData(data, labels)

' Output the chart
Response.ContentType = "image/png"
Response.BinaryWrite c.makeChart2(cd.PNG)
Response.End
%>
