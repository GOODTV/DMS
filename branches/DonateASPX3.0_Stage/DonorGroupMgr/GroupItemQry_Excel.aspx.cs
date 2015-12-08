using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonorGroupMgr_GroupItemQry_Excel : BasePage
{

    protected void Page_Load(object sender, EventArgs e)
    {
        string strGridList = Session["GridList"].ToString();
        Print_Excel(strGridList);
    }
    //---------------------------------------------------------------------------
    private void Print_Excel(string strGridList)
    {

        GridList.Text = GetTitle() + strGridList.Replace("class='table_h'", "border='1'");
        Util.OutputTxt(GridList.Text, "1", "GroupItemQry");

    }
    //---------------------------------------------------------------------------
    private string GetTitle()
    {

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        table.Width = "1200px";
        css = table.Style;
        css.Add("font-size", "18px");
        css.Add("font-family", "標楷體");

        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = "捐款人群組查詢<br/> ";
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-weight", "bold");
        //css.Add("font-size", strFontSize);
        cell.ColSpan = 5;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }

}
