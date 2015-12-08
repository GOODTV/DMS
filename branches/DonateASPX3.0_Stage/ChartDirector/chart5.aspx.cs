using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using ChartDirector;

public partial class ChartDirector_chart5 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        { 
            Util.FillDropDownList(yearSelect, DateTime.Now.Year - 15, DateTime.Now.Year, false, -1);
            yearSelect.Text = DateTime.Now.Year.ToString();
        }
        
        // The currently selected year
        string selectedYear = yearSelect.SelectedItem.Value;

        // SQL statement.
        string SQL =
            @"SELECT CONVERT(varchar,Donate_Payment),COUNT(Donate_Id) AS '筆數',Sum(Donate_Amt)/10000 AS Sum_Amt FROM dbo.DONATE 
                WHERE Donate_Purpose <> '' AND Donate_Amt > 0  
                AND DATEPART(YEAR,Donate_Date) = " +
                selectedYear + " GROUP by Donate_Payment";

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
        XYChart c = new XYChart(480, 410);

        // Set the plot area at (50, 20) and of size 200 x 130 pixels
        c.setPlotArea(65, 10, 350, 260);

        // Add a title to the chart using 8 pts Arial Bold font
        //c.addTitle("Independent Y-Axis Demo", "Arial Bold", 8);

        // Set the labels on the x axis.
        c.xAxis().setLabels(labels);
        c.xAxis().setLabelStyle("kaiu.ttf Bold", 10, 0x000000, 90);

        // Add a title to the primary (left) y axis
        c.yAxis().setTitle("總筆數", "kaiu.ttf Bold", 12);

        // Set the axis, label and title colors for the primary y axis to red (0xc00000)
        // to match the first data set
        c.yAxis().setColors(0xFF1493, 0xFF1493, 0xFF1493);

        // Add a title to the secondary (right) y axis
        c.yAxis2().setTitle("總捐款金額(單位：萬元)", "kaiu.ttf Bold", 12);

        // set the axis, label and title colors for the primary y axis to green
        // (0x008000) to match the second data set
        c.yAxis2().setColors(0x696969, 0x696969, 0x696969);

        // Add a line layer to for the first data set using red (0xc00000) color with a
        // line width to 3 pixels
        LineLayer lineLayer = c.addLineLayer(data1, 0xFF1493);
        lineLayer.setLineWidth(3);

        // tool tip for the line layer
        lineLayer.setHTMLImageMap("", "",
            "title='總筆數 on {xLabel}: {value} 筆'");

        // Add a bar layer to for the second data set using green (0x00C000) color. Bind
        // the second data set to the secondary (right) y axis
        BarLayer barLayer = c.addBarLayer(data2, 0x696969);
        barLayer.setUseYAxis2();

        // tool tip for the bar layer
        barLayer.setHTMLImageMap("", "", "title='總捐款金額 in {xLabel}: {value} 萬元'"
            );

        // Output the chart
        WebChartViewer1.Image = c.makeWebImage(Chart.PNG);

        // include tool tip for the chart
        WebChartViewer1.ImageMap = c.getHTMLImageMap("");
    }
}