using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using ChartDirector;

public partial class ChartDirector_chart3 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string year_now = DateTime.Now.ToString("yyyy");

        // SQL statement.
        string SQL =
            @"select Convert(varchar,DATEPART(Month,Donate_CreateDate))+'月' as Month,COUNT(orderid) AS '筆數' ,
                Replace(Convert(Varchar,Sum(Donate_Amount),1),'.00','') AS Sum_Amt
                from DONATE_IEPAY P left join DONATE_Web W on P.orderid=W.od_sob 
                where status = '0'  
                AND DATEPART(YEAR,Donate_CreateDate) = " +
                year_now + " and DATEPART(Month,Donate_CreateDate) between 6 and DATEPART(Month,GETDATE()) GROUP by DATEPART(Month,Donate_CreateDate)";

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
        double[] data1 = table.getCol(1);
        double[] data2 = table.getCol(2);
        string[] labels = table.getColAsString(0);

        // Create a XYChart object of size 300 x 180 pixels
        XYChart c = new XYChart(480, 300);

        // Set the plot area at (50, 20) and of size 200 x 130 pixels
        c.setPlotArea(65, 10, 350, 260);

        // Add a title to the chart using 8 pts Arial Bold font
        //c.addTitle("Independent Y-Axis Demo", "Arial Bold", 8);

        // Set the labels on the x axis.
        c.xAxis().setLabels(labels);

        // Add a title to the primary (left) y axis
        c.yAxis().setTitle("總筆數", "kaiu.ttf Bold", 12);

        // Set the axis, label and title colors for the primary y axis to red (0xc00000)
        // to match the first data set
        c.yAxis().setColors(0xFF0000, 0xFF0000, 0xFF0000);

        // Add a title to the secondary (right) y axis
        c.yAxis2().setTitle("總捐款金額", "kaiu.ttf Bold", 12);

        // set the axis, label and title colors for the primary y axis to green
        // (0x008000) to match the second data set
        c.yAxis2().setColors(0x4169E1, 0x4169E1, 0x4169E1);

        // Add a line layer to for the first data set using red (0xc00000) color with a
        // line width to 3 pixels
        LineLayer lineLayer = c.addLineLayer(data1, 0xFF0000);
        lineLayer.setLineWidth(3);

        // tool tip for the line layer
        lineLayer.setHTMLImageMap("", "",
            "title='總筆數 on {xLabel}: {value} 筆'");

        // Add a bar layer to for the second data set using green (0x00C000) color. Bind
        // the second data set to the secondary (right) y axis
        BarLayer barLayer = c.addBarLayer(data2, 0x4169E1);
        barLayer.setUseYAxis2();

        // tool tip for the bar layer
        barLayer.setHTMLImageMap("", "", "title='總捐款金額 in {xLabel}: {value} 元'"
            );

        // Output the chart
        WebChartViewer1.Image = c.makeWebImage(Chart.PNG);

        // include tool tip for the chart
        WebChartViewer1.ImageMap = c.getHTMLImageMap("");
    }
}