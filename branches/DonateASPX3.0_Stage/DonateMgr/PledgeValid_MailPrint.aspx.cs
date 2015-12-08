using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonateMgr_PledgeValid_MailPrint : BasePage
{
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        Print();
        if (GridList.Text != "** 沒有符合條件的資料 **")
        {
            Response.Write("<script language=javascript>window.print();</script>");
        }
    }
    private void Print()
    {
        Dictionary<string, int> ReportData = new Dictionary<string, int>();
        GridList.Text = GetTable(ReportData);
    }
    private string GetTable(Dictionary<string, int> ReportData)
    {
        string strSql = Session["strSql"].ToString();
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr;

        if (dt.Rows.Count == 0)
        {
            return "** 沒有符合條件的資料 **";
        }
        else
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

            table.Border = 0;
            table.CellPadding = 0;
            table.CellSpacing = 0;
            css = table.Style;
            css.Add("font-size", "16px");
            css.Add("font-family", "標楷體");
            css.Add("width", "100%");
            css.Add("border-collapse", "collpase");
            string strFontSize = "18px";
            //--------------------------------------------
            //data
            int count = dt.Rows.Count;
            for (int i = 0; i < count; i++)
            {
                row = new HtmlTableRow();
                cell = new HtmlTableCell();
                dr = dt.Rows[i];
                strTemp = dr2["EmailMgr_Desc"].ToString();
                //代換成個資
                strTemp = strTemp.Replace("○○○", dr["捐款人"].ToString());
                strTemp = strTemp.Replace("YY1", "20" + dr["Year"].ToString());
                strTemp = strTemp.Replace("MM1", dr["Month"].ToString().PadLeft(2, '0'));
                strTemp = strTemp.Replace("YY2", "20" + dr["Year"].ToString());
                strTemp = strTemp.Replace("MM2", (int.Parse(dr["Month"].ToString()) + 1).ToString().PadLeft(2, '0'));
                strTemp = strTemp.Replace("年月日", DateTime.Now.Year.ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Day.ToString());
                if (i != count - 1)
                {
                    strTemp += "<br clear=all style='page-break-before:always'>";  //列印強制換頁,最後一頁不用換頁

                }
                cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                css = cell.Style;
                css.Add("text-align", "left");
                css.Add("font-size", strFontSize);
                css.Add("width", "100%");
                css.Add("line-height", "150%");
                css.Add("letter-spacing", "1px");
                row.Cells.Add(cell);
                table.Rows.Add(row);
            }
            //--------------------------------------------

            //轉成 html 碼
            StringWriter sw = new StringWriter();

            HtmlTextWriter htw = new HtmlTextWriter(sw);
            table.RenderControl(htw);

            return htw.InnerWriter.ToString();
        }
    }
    //end of GetTable()
}