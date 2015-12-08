using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class Report_DonateAccounReport_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        PrintReport();
    }
    //---------------------------------------------------------------------------
    private void PrintReport()
    {
        string strSql = Session["strSql"].ToString();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "沒有符合條件的資料可以列印!!";
            return;
        }

        DataTable dtRet = CaseUtil.DonateAccounReport_Print(dt);

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "12px");
        css.Add("font-family", "細明體");
        css.Add("line-height", "20px");
        row = new HtmlTableRow();

        int iCtrl = 0;
        foreach (DataColumn dc in dtRet.Columns)
        {
            cell = new HtmlTableCell();
            //cell.Style.Add("border-left", ".5pt solid windowtext");
            cell.Style.Add("border-top", ".5pt solid windowtext");
            //cell.Style.Add("border-right", ".5pt solid windowtext");
            cell.Style.Add("border-bottom", ".5pt solid windowtext");
            
            if (iCtrl == 0)
            {
                cell.Style.Add("border-left", ".5pt solid windowtext");
                cell.Width = "250mm";
                cell.Style.Add("text-align", "left");
            }
            if (iCtrl == 1)
            {
                cell.Width = "90mm";
                cell.Style.Add("text-align", "right");
            }
            if (iCtrl == 2)
            {
                cell.Width = "90mm";
                cell.Style.Add("text-align", "right");
            }
            if (iCtrl == 3)
            {
                cell.Width = "120mm";
                cell.Style.Add("text-align", "right");
            }
            if (iCtrl == 4)
            {
                cell.Style.Add("border-right", ".5pt solid windowtext");
                cell.Width = "90mm";
                cell.Style.Add("text-align", "right");
            }

            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

        foreach (DataRow dr in dtRet.Rows)
        {
            int Col = 0;
            row = new HtmlTableRow();
            foreach (object objItem in dr.ItemArray)
            {
                cell = new HtmlTableCell();
                //cell.Style.Add("border-left", ".5pt solid windowtext");
                cell.Style.Add("border-top", ".5pt solid windowtext");
                //cell.Style.Add("border-right", ".5pt solid windowtext");
                cell.Style.Add("border-bottom", ".5pt solid windowtext");
                cell.Style.Add("text-align", "right");
                cell.InnerHtml = objItem.ToString() == "" ? "&nbsp" : objItem.ToString();
                //if()
                if (Col == 0)
                {
                    cell.Style.Add("text-align", "left");
                    cell.Style.Add("border-left", ".5pt solid windowtext");
                }
                if (Col == 4)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                }
                row.Cells.Add(cell);
                Col++;
            }
            table.Rows.Add(row);
        }
        //轉成 html 碼
        StringWriter sw = new StringWriter();

        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        GridList.Text = "會計科目報表<br/>" + "<div span style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + Session["Date"].ToString() + "</span>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
}