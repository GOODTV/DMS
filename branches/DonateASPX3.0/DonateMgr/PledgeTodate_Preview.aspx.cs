using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonateMgr_PledgeTodate_Preview : BasePage
{
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        Print();
    }
    private void Print()
    {
        Dictionary<string, int> ReportData = new Dictionary<string, int>();
        GridList.Text = GetTable(ReportData);
    }
    private string GetTable(Dictionary<string, int> ReportData)
    {
        //撈出信件內容
        string strSql2 = "select * from EmailMgr where EmailMgr_Subject='" + Session["Email_Subject"].ToString() + "'";
        DataTable dt2 = NpoDB.QueryGetTable(strSql2);
        DataRow dr2;
        dr2 = dt2.Rows[0];

        //組 table
        string strTemp = "";
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;
        table.Style.Add("width", "194mm");
        table.Style.Add("font-size", "13px");
        css = table.Style;
        //--------------------------------------------
        //data

        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        strTemp = dr2["EmailMgr_Desc"].ToString();
        //代換成個資
        strTemp = strTemp.Replace("○○○", "加百列");
        strTemp = strTemp.Replace("YY1", DateTime.Now.Year.ToString());
        strTemp = strTemp.Replace("MM1", DateTime.Now.Month.ToString().PadLeft(2, '0'));
        strTemp = strTemp.Replace("YY2", DateTime.Now.Year.ToString());
        strTemp = strTemp.Replace("MM2", (int.Parse(DateTime.Now.Month.ToString()) + 1).ToString().PadLeft(2, '0'));
        strTemp = strTemp.Replace("年月日", DateTime.Now.ToString("yyyy/MM/dd"));
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        row.Cells.Add(cell);
        table.Rows.Add(row);
        //--------------------------------------------

        //轉成 html 碼
        StringWriter sw = new StringWriter();

        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    } 
    //end of GetTable()
}