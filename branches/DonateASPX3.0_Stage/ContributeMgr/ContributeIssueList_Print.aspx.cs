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

public partial class ContributeMgr_ContributeIssue_Print : BasePage
{
    string TableWidth = "800px";

    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        PrintReport();
    }
    //---------------------------------------------------------------------------
    private void PrintReport()
    {
        string strSql = Session["strSql"].ToString();
        strSql = strSql.Replace("from Contribute_Issue CI", "");
        //strSql = strSql.Replace("where 1=1", "");
        strSql = strSql.Replace(" WhereClause", @" ,left(C.Item,len(C.Item)-1) as [領用內容] from
                                                    (select Issue_Id,(SELECT cast(CI.Goods_Name+'(' + (CONVERT(VARCHAR,CI.Goods_Qty))+CI.Goods_Unit+')' AS NVARCHAR ) + ',' 
                                                    from Contribute_IssueData CI right join Goods G on CI.Goods_Id = G.Goods_Id
                                                    where Issue_Id = CID.Issue_Id
                                                    FOR XML PATH('')) as [Item]
                                                    from Contribute_IssueData CID
                                                    GROUP BY Issue_Id)C
                                                    left join Contribute_Issue CI on CI.Issue_Id = C.Issue_Id");
        Dictionary<string, object> dict = new Dictionary<string, object>();

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
        }

        DataTable dtRet = CaseUtil.ContributeIssueList_Print(dt);

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

            if (iCtrl == 0 || iCtrl == 7)
            {
                cell.Width = "50mm";
            }
            else if (iCtrl == 1)
            {
                cell.Width = "100mm";
            }
            else if (iCtrl == 2 || iCtrl == 3)
            {
                cell.Width = "65mm";
            }
            else if (iCtrl == 5 || iCtrl == 6)
            {
                cell.Width = "35mm";
            }
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
                //cell.Style.Add("font-family", "標楷體");
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
        GridList.Text = "物品領用明細表<br/>" + htw.InnerWriter.ToString();

    }
}