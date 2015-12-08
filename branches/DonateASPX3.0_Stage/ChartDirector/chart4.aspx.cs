using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using ChartDirector;

public partial class ChartDirector_chart4 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        // SQL statement.
        string SQL =
            @"SELECT top 3 DATEPART(YEAR,Donate_Date) AS '年度',ISNULL(Sum(Donate_Amt)/100000000,0) AS Sum_Amt FROM dbo.DONATE
                WHERE Donate_Purpose <> '' AND Donate_Amt > 0 
                group by DATEPART(YEAR,Donate_Date) 
                union 
                SELECT DATEPART(YEAR,Donate_Date) AS '年度',ISNULL(Sum(Donate_Amt)/100000000,0) AS Sum_Amt FROM dbo.DONATE 
                WHERE Donate_Purpose <> '' AND Donate_Amt > 0 and DATEPART(YEAR,Donate_Date) between 2007 and DATEPART(YEAR,GETDATE()) 
                group by DATEPART(YEAR,Donate_Date)";

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
        // The data for the chart
       
        double[] data = table.getCol(1);
        // The labels for the bar chart
        string[] labels = table.getColAsString(0);
        // Create a XYChart object of size 400 x 240 pixels.
        XYChart c = new XYChart(480, 270);

        // Add a title to the chart using 14 pts Times Bold Italic font
        //c.addTitle("Weekly Server Load", "Times New Roman Bold Italic", 14);

        // Set the plotarea at (45, 40) and of 300 x 160 pixels in size. Use alternating
        // light grey (f8f8f8) / white (ffffff) background.
        c.setPlotArea(60, 30, 400, 200, 0xf8f8f8, 0xffffff);

        // Add a multi-color bar chart layer
        BarLayer layer = c.addBarLayer3(data);

        // Set layer to 3D with 10 pixels 3D depth
        layer.set3D(10);

        // Set bar shape to circular (cylinder)
        layer.setBarShape(Chart.CircleShape);

        // Set the labels on the x axis.
        c.xAxis().setLabels(labels);

        // Add a title to the y axis
        c.yAxis().setTitle("總捐款金額(億/單位)", "kaiu.ttf Bold", 12);

        // Add a title to the x axis
        //c.xAxis().setTitle("Work Week 25");

        // Output the chart
        WebChartViewer1.Image = c.makeWebImage(Chart.PNG);

        // Include tool tip for the chart
        WebChartViewer1.ImageMap = c.getHTMLImageMap("", "",
            "title='{xLabel}: {value} 億元'");
    }
}