using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using ChartDirector;

public partial class ChartDirector_CSharpASP_chart2 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string year_now = DateTime.Now.ToString("yyyy");

        // SQL statement.
        string SQL =
            @"SELECT Donate_Payment ,Sum(Donate_Amt) AS Amt FROM dbo.DONATE
                WHERE Donate_Purpose <> '' AND Donate_Amt > 0
                AND DATEPART(YEAR,Donate_Date) = " +
                year_now + " GROUP BY Donate_Payment ORDER by Sum(Donate_Amt) desc";

        //
        // Connect to database and read the query result into arrays
        //

        // In this example, we use OleDbConnection to connect to MS Access (Jet Engine).
        // If you are using MS SQL, you can use SqlConnection instead of OleConnection.

        System.Data.IDbConnection dbconn = new System.Data.SqlClient.SqlConnection(NpoDB.GetConnectionString());
        dbconn.Open();

        // Set up the SQL statement
        System.Data.IDbCommand sqlCmd = dbconn.CreateCommand();
        sqlCmd.CommandText = SQL;

        // Read the data into the DBTable object
        DBTable table = new DBTable(sqlCmd.ExecuteReader());
        dbconn.Close();

        // The data for the pie chart
        double[] data = table.getCol(1);

        // The labels for the pie chart
        string[] labels = table.getColAsString(0);

        // The colors to use for the sectors
        int[] colors = { 0x66aaee, 0xeebb22, 0xbbbbbb, 0x8844ff, 0xdd2222, 0x009900 };

        // Create a PieChart object of size 600 x 320 pixels. Use a vertical gradient
        // color from light blue (99ccff) to white (ffffff) spanning the top 100 pixels
        // as background. Set border to grey (888888). Use rounded corners. Enable soft
        // drop shadow.
        PieChart c = new PieChart(470, 270);
        c.setBackground(c.linearGradientColor(0, 0, 0, 100, 0x99ccff, 0xffffff), 0x888888
            );
        c.setRoundedFrame();
        c.setDropShadow();

        // Add a title using 18 pts Times New Roman Bold Italic font. Add 16 pixels top
        // margin to the title.
        //c.addTitle("Pie Chart With Legend Demonstration", "Times New Roman Bold Italic",
        //    18).setMargin2(0, 0, 16, 0);

        // Set the center of the pie at (160, 165) and the radius to 110 pixels
        c.setPieSize(120, 140, 80);

        // Draw the pie in 3D with a pie thickness of 25 pixels
        c.set3D(15);

        // Set the pie data and the pie labels
        c.setData(data, labels);

        // Set the sector colors
        c.setColors2(Chart.DataColor, colors);

        // Use local gradient shading for the sectors
        c.setSectorStyle(Chart.LocalGradientShading);

        // Use the side label layout method, with the labels positioned 16 pixels from
        // the pie bounding box
        c.setLabelLayout(Chart.SideLayout, 16);

        // Show only the sector number as the sector label
        c.setLabelFormat("{={sector}+1}");

        // Set the sector label style to Arial Bold 10pt, with a dark grey (444444)
        // border
        c.setLabelStyle("Arial Bold", 10).setBackground(Chart.Transparent, 0x444444);

        // Add a legend box, with the center of the left side anchored at (330, 175), and
        // using 10 pts Arial Bold Italic font
        LegendBox b = c.addLegend(235, 135, true, "simhei.ttf", 9);
        b.setAlignment(Chart.Left);

        // Set the legend box border to dark grey (444444), and with rounded conerns
        b.setBackground(Chart.Transparent, 0x444444);
        b.setRoundedCorners();

        // Set the legend box margin to 16 pixels, and the extra line spacing between the
        // legend entries as 5 pixels
        b.setMargin(14);
        b.setKeySpacing(0, 4);

        // Set the legend box icon to have no border (border color same as fill color)
        b.setKeyBorder(Chart.SameAsMainColor);

        // Set the legend text to show the sector number, followed by a 120 pixels wide
        // block showing the sector label, and a 40 pixels wide block showing the
        // percentage
        b.setText(
            "<*block,valign=top*>{={sector}+1}.<*advanceTo=20*><*block,width=120*>" +
            "{label}<*/*><*block,width=35,halign=right*>{percent}<*/*>%");

        // Output the chart
        WebChartViewer1.Image = c.makeWebImage(Chart.PNG);

        // Include tool tip for the chart
        WebChartViewer1.ImageMap = c.getHTMLImageMap("", "",
            "title='{label}: NT${value}元 ({percent}%)'");
    }
}