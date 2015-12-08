<!--#include file="../include/dbfunctionJ.asp"-->
<%
month_now=month(Date)
'month_now= 2012

Set cd = CreateObject("ChartDirector.API")

SQL= "SELECT CASE paytype when '1' then '信用卡' when '2' then 'iePay儲值帳戶付款' when '4' then 'PayPal' " & _
	 "when '5' then '其他超商 電子帳單' when '8' then 'Web ATM' when '9' then '7-11 電子帳單' " & _
	 "when '12' then '玉山銀行eCoin' when '16' then '郵局電子帳單付款' " & _
	 "when '17' then '郵局ATM付款' when '25' then '24hr超商取貨付款' " & _
	 "when '30' then '7-11 ibon' when '35' then '條碼' when '39' then '全家FamilyPort' end as Donate_Payment " & _
	 ",Sum(Donate_Amount) AS Amt " & _
	 "From DONATE_WEB W " & _
	 "left join DONATE_IEPAY P on P.orderid=W.od_sob " & _
     "WHERE paytype <> '' AND status = '0' " & _
	 "AND DATEPART(Month,Donate_CreateDate) = '" & month_now & "' GROUP BY paytype ORDER by Sum(Donate_Amount) desc"
	  
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open SQL,Conn,1,1

' Create a PieChart object of size 560 x 270 pixels, with a golden background and a 1
' pixel 3D border
Set c = cd.PieChart(460, 240, cd.goldColor(), -1, 0)

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
Call c.setPieSize(210, 100, 80)

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
