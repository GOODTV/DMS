using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using ChartDirector;

public partial class ChartDirector_chart1 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string year_now = DateTime.Now.ToString("yyyy");
        
        // SQL statement.
        string SQL =
            @"SELECT CASE City
                when 'A' then '台北市' when 'B' then '新北市' when 'C' then '基隆市' when 'D' then '宜蘭縣' when 'E' then '桃園縣' 
                when 'F' then '新竹市' when 'G' then '新竹縣' when 'H' then '苗栗縣' when 'I' then '台中市' when 'K' then '彰化縣' 
                when 'L' then '南投縣' when 'M' then '雲林縣' when 'N' then '嘉義市' when 'O' then '嘉義縣' when 'P' then '台南市' 
                when 'R' then '高雄市' when 'T' then '屏東縣' when 'U' then '台東縣' when 'V' then '花蓮縣' when 'W' then '澎湖縣' 
                when 'X' then '金門縣' when 'Y' then '連江縣' when '' then '不具名' else '未填縣市' end as City ,Sum(Donate_Amt) AS Sum_Amt 
                From DONATE D left join DONOR N on D.Donor_Id=N.Donor_Id 
                WHERE Donate_Amt > 0 and Issue_Type <> 'D' and IsAbroad = 'N' AND DATEPART(YEAR,Donate_Date) = " +
                year_now + " GROUP BY City ORDER by Sum(Donate_Amt) desc";

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

        // Create a PieChart object of size 560 x 270 pixels, with a golden background
        // and a 1 pixel 3D border
        PieChart c = new PieChart(440, 370, Chart.goldColor(), -1,0);

        // Add a title box using 15 pts Times Bold Italic font and metallic pink
        // background color
        //c.addTitle("Project Cost Breakdown", "Times New Roman Bold Italic", 15
        //    ).setBackground(Chart.metalColor(0xff9999));

        // Set the center of the pie at (280, 135) and the radius to 110 pixels
        c.setPieSize(215, 140, 80);

        // Draw the pie in 3D with 20 pixels 3D depth
        c.set3D(15);

        // Use the side label layout method
        c.setLabelLayout(Chart.SideLayout);

        // Set the label box background color the same as the sector color, with glass
        // effect, and with 5 pixels rounded corners
        ChartDirector.TextBox t = c.setLabelStyle();
        t.setBackground(Chart.SameAsMainColor, Chart.Transparent, Chart.glassEffect());
        t.setRoundedCorners(5);

        // Set the border color of the sector the same color as the fill color. Set the
        // line color of the join line to black (0x0)
        c.setLineColor(Chart.SameAsMainColor, 0x000000);

        // Set the start angle to 135 degrees may improve layout when there are many
        // small sectors at the end of the data array (that is, data sorted in descending
        // order). It is because this makes the small sectors position near the
        // horizontal axis, where the text label has the least tendency to overlap. For
        // data sorted in ascending order, a start angle of 45 degrees can be used
        // instead.
        c.setStartAngle(135);

        // Set the pie data and the pie labels
        c.setData(data, labels);

        // Output the chart
        WebChartViewer1.Image = c.makeWebImage(Chart.PNG);

        // Include tool tip for the chart
        WebChartViewer1.ImageMap = c.getHTMLImageMap("", "",
            "title='{label}: NT${value}元 ({percent}%)'");
    }
}