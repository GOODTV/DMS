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

public partial class ContributeMgr_ContributeList_Print : BasePage
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
        //if (strSql=="")
        //{
        //    GridList.Text = "** 沒有符合條件的資料 **";
        //    return;
        //}
        strSql = strSql.Replace("from Contribute C left join Donor D on C.Donor_Id = D.Donor_Id", "");
        //strSql = strSql.Replace("where 1=1", "");
        strSql = strSql.Replace(" WhereClause", @"  ,left(C.Item,len(C.Item)-1) as [捐贈內容] from    
                                                    (select CC.Dept_Id,CC.Donor_id,CC.Contribute_Id,CC.Contribute_Date,CC.Contribute_Amt,CC.Contribute_Payment,CC.Contribute_Purpose,CC.InvoiCe_Type,CC.Invoice_Pre,CC.Invoice_No,CC.Accoun_Date,CC.Invoice_Print,CC.Issue_Type,CC.Export,CC.Create_User,
														(SELECT cast(CD.Goods_Name+'(' + (CONVERT(VARCHAR,CD.Goods_Qty))+CD.Goods_Unit+')' AS NVARCHAR ) + ','
														from Contribute CCC left join ContributeData CD on CCC.Contribute_Id = CD.Contribute_Id 
                                                        right join Goods G on CD.Goods_Id = G.Goods_Id
														where CC.Contribute_Id  = CCC.Contribute_Id FOR XML PATH('')) as [Item]
                                                    from Contribute CC left join Donor D on CC.Donor_Id = D.Donor_Id )C 
                                                    left join Donor D on C.Donor_Id = D.Donor_Id ");
        string Amt = Session["Amt"].ToString();
        strSql += @"union 
                  select '9999999',null,'' ,'折合現金合計：','" + Amt + "' ,'','','','','','','',''";

        
        Dictionary<string, object> dict = new Dictionary<string, object>();

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
        }

        DataTable dtRet = CaseUtil.ContributeList_Print(dt);

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

            if (iCtrl == 10)
            {
                cell.Width = "40mm";
            }
            else if (iCtrl == 1)
            {
                cell.Width = "90mm";
            }
            else if (iCtrl == 2)
            {
                cell.Width = "55mm";
            }
            else if (iCtrl == 4 || iCtrl == 5)
            {
                cell.Width = "50mm";
            }
            else if (iCtrl == 6)
            {
                cell.Width = "70mm";
            }
            else if (iCtrl == 7)
            {
                cell.Width = "90mm";
            }
            else if (iCtrl == 9)
            {
                cell.Width = "25mm";
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
        GridList.Text = "實物奉獻捐贈維護<br/>" + htw.InnerWriter.ToString();

    }
}