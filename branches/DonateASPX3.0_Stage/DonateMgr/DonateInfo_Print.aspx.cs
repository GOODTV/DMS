using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonateMgr_DonateInfo_Print : BasePage
{
    string TableWidth = "800px";

    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["strSql"] != null)
        {
            PrintReport();
        }
        else
        {
            AjaxShowMessage("請先查詢後再做列表輸出！");
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >window.close();</script>");
        }
    }
    //---------------------------------------------------------------------------
    private void PrintReport()
    {
        string strSql = Session["strSql"].ToString();
        strSql = strSql.Replace("order by do.Donate_Date desc", "");
        string Amt = Session["Amt"].ToString();
        strSql += @"union 
                  select '9999999',null,'','','捐款合計：',+'" + Amt + "' ,'','','','','','','','','','' from Donate";
        Dictionary<string, object> dict = new Dictionary<string, object>();

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
        }

        DataTable dtRet = CaseUtil.DonateInfo_Print(dt);

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        table.Width = "100%";
        css = table.Style;
        css.Add("font-size", "12px");
        css.Add("font-family", "細明體");
        css.Add("line-height", "20px");

        row = new HtmlTableRow();

        int iCtrl = 0;
        foreach (DataColumn dc in dtRet.Columns)
        {
            cell = new HtmlTableCell();
            cell.Style.Add("background-color", " #FFE4C4 ");
            cell.Style.Add("border-left", ".5pt solid windowtext");
            cell.Style.Add("border-top", ".5pt solid windowtext");
            cell.Style.Add("border-right", ".5pt solid windowtext");
            cell.Style.Add("border-bottom", ".5pt solid windowtext");

            //if (iCtrl == 0 || iCtrl == 9 || iCtrl == 10 || iCtrl == 11)
            //{
            //    cell.Width = "50mm";
            //}
            //else if (iCtrl == 4)
            //{
            //    cell.Width = "70mm";
            //}
            //else if (iCtrl == 1 || iCtrl == 7 || iCtrl == 8)
            //{
            //    cell.Width = "90mm";
            //}
            //else if (iCtrl == 6)
            //{
            //    cell.Width = "100mm";
            //}
            //else
            //{
            //    cell.Width = "80mm";
            //}
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

        foreach (DataRow dr in dtRet.Rows)
        {
            row = new HtmlTableRow();
                foreach (object objItem in dr.ItemArray)
                {
                    cell = new HtmlTableCell();
                    cell.Style.Add("border-left", ".5pt solid windowtext");
                    cell.Style.Add("border-top", ".5pt solid windowtext");
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                    cell.Style.Add("border-bottom", ".5pt solid windowtext");
                    cell.InnerHtml = objItem.ToString() == "" ? "&nbsp" : objItem.ToString();
                    row.Cells.Add(cell);
                }
            table.Rows.Add(row);
        }
        //轉成 html 碼
        StringWriter sw = new StringWriter();

        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);
        GridList.Text = "捐款資料維護<br/>" + htw.InnerWriter.ToString();

    }
}