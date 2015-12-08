using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class ContributeMgr_ContributeList_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        Print_Excel();
    }
    //---------------------------------------------------------------------------
    private void Print_Excel()
    {
        string strSql = Session["strSql"].ToString();
        strSql = strSql.Replace("from Contribute C left join Donor D on C.Donor_Id = D.Donor_Id", "");
        //strSql = strSql.Replace("where 1=1", "");
        strSql = strSql.Replace(" WhereClause", @" ,left(C.Item,len(C.Item)-1) as [捐贈內容] from    
                                                    (select CC.Dept_Id,CC.Donor_id,CC.Contribute_Id,CC.Contribute_Date,CC.Contribute_Amt,CC.Contribute_Payment,CC.Contribute_Purpose,CC.InvoiCe_Type,CC.Invoice_Pre,CC.Invoice_No,CC.Accoun_Date,CC.Invoice_Print,CC.Issue_Type,CC.Export,CC.Create_User,
														(SELECT cast(CD.Goods_Name+'(' + (CONVERT(VARCHAR,CD.Goods_Qty))+CD.Goods_Unit+')' AS NVARCHAR ) + ','
														from Contribute CCC left join ContributeData CD on CCC.Contribute_Id = CD.Contribute_Id 
                                                        right join Goods G on CD.Goods_Id = G.Goods_Id
														where CC.Contribute_Id  = CCC.Contribute_Id FOR XML PATH('')) as [Item]
                                                    from Contribute CC left join Donor D on CC.Donor_Id = D.Donor_Id )C 
                                                    left join Donor D on C.Donor_Id = D.Donor_Id ");
        string Amt = Session["Amt"].ToString();
        strSql += @"union 
                  select '9999999',null,'' ,'折合現金合計：','" + Amt + "','' ,'','','','','','',''";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            ShowSysMsg("查無資料!!!");
            return;
        }
        GridList.Text = GetTitle(dt.Rows[0]) + GetTable1(dt.Rows[0], strSql);
        //GetTableFooter();
        Util.OutputTxt(GridList.Text, "1", "Contribute");

    }
    //---------------------------------------------------------------------------
    private string GetTitle(DataRow dr)
    {
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "15px");
        css.Add("font-family", "標楷體");
        css.Add("width", TableWidth);

        string strFontSize = "18px";
        string strTitle = "實物奉獻捐贈維護<br/> ";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 11;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    //---------------------------------------------------------------------------
    private string GetTable1(DataRow dr, string strSql)
    {
        //組 table
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;

        DataTable dtRet = CaseUtil.ContributeList_Print(dt);

        table.Border = 1;
        table.CellPadding = 3;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "15px");
        css.Add("font-family", "新細明體");
        css.Add("width", TableWidth);
        css.Add("line-height", "25px");

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

            cell.Width = "100mm";
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

        foreach (DataRow drow in dtRet.Rows)
        {

            row = new HtmlTableRow();
            foreach (object objItem in drow.ItemArray)
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

        return htw.InnerWriter.ToString();
    }
}