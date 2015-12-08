using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonorMgr_DonorQry_Print : BasePage
{
    string TableWidth = "800px";

    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {

        PrintReport();
        //Response.Write("<script language=javascript>window.print();</script>");
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
            GridList.Text = "** 沒有符合條件的資料 **";
            return;
        }

        DataTable dtRet = CaseUtil.DonorQry_Print(dt);

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
            cell.Style.Add("background-color", " #FFE4C4 ");
            cell.Style.Add("border-left", ".5pt solid windowtext");
            cell.Style.Add("border-top", ".5pt solid windowtext");
            cell.Style.Add("border-right", ".5pt solid windowtext");
            cell.Style.Add("border-bottom", ".5pt solid windowtext");

            //if (iCtrl == 0)
            //{
            //    cell.Width = "90mm";
            //}
            //if (iCtrl == 1)
            //{
            //    cell.Width = "20mm";
            //}
            //if (iCtrl == 2)
            //{
            //    cell.Width = "25mm";
            //}
            //if (iCtrl == 3)
            //{
            //    cell.Width = "35mm";
            //}
            //if (iCtrl == 4)
            //{
            //    cell.Width = "45mm";
            //}
            //if (iCtrl == 5)
            //{
            //    cell.Width = "45mm";
            //}
            //if (iCtrl == 6)
            //{
            //    cell.Width = "45mm";
            //}
            //if (iCtrl == 7)
            //{
            //    cell.Width = "70mm";
            //}
            //if (iCtrl == 8)
            //{
            //    cell.Width = "220mm";
            //}
            //if (iCtrl == 9)
            //{
            //    cell.Width = "60mm";
            //}
            //if (iCtrl == 10)
            //{
            //    cell.Width = "20mm";
            //}
            //else
            //{
            //    cell.Width = "120mm";
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

        GridList.Text = GetTitle() + GetTime() + htw.InnerWriter.ToString() + "<br/>";
    }
    private string GetTitle()
    {
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "12pt");
        css.Add("font-family", "標楷體");
        css.Add("width", "100%");

        string strFontSize = "18px";
        string strTitle = "捐款報表(捐款人)<br/> ";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        row.Cells.Add(cell);
        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();

    }
    private string GetTime()
    {
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "10pt");
        css.Add("font-family", "標楷體");
        css.Add("width", "100%");
        css.Add("line-height", "20px");

        string strTitle = "列印日期：" + DateTime.Now.ToLocalTime().ToString() + "<br/> ";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "right");
        row.Cells.Add(cell);
        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
}