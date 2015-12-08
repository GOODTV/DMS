using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonateMgr_Web_ePayList_Print : BasePage
{
    //string TableWidth = "800px";

    //---------------------------------------------------------------------------
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
            GridList.Text = "** 沒有符合條件的資料 **";
        }

        //DataTable dtRet = Web_ePayList_Print(dt);

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
        //foreach (DataColumn dc in dtRet.Columns)
        foreach (DataColumn dc in dt.Columns)
        {
            cell = new HtmlTableCell();
            cell.Style.Add("background-color", " #FFE4C4 ");
            cell.Style.Add("border-left", ".5pt solid windowtext");
            cell.Style.Add("border-top", ".5pt solid windowtext");
            cell.Style.Add("border-right", ".5pt solid windowtext");
            cell.Style.Add("border-bottom", ".5pt solid windowtext");

            //if (iCtrl == 0)
            //{
            //    cell.Width = "60";
            //}
            //else if (iCtrl == 1)
            //{
            //    cell.Width = "60";
            //}
            //else if (iCtrl == 3)
            //{
            //    cell.Width = "60";
            //}
            //else if (iCtrl == 6)
            //{
            //    cell.Width = "60";
            //}
            //else if (iCtrl == 8)
            //{
            //    cell.Width = "100";
            //}
            //else if (iCtrl == 9)
            //{
            //    cell.Width = "70";
            //}
            //else if (iCtrl == 10)
            //{
            //    cell.Width = "70";
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

        int cnt = 0;
        Double Donate_Amt = 0;
       //foreach (DataRow dr in dtRet.Rows)
        foreach (DataRow dr in dt.Rows)
        {
            row = new HtmlTableRow();
            int i = 0;
            foreach (object objItem in dr.ItemArray)
            {
                cell = new HtmlTableCell();
                cell.Style.Add("border-left", ".5pt solid windowtext");
                cell.Style.Add("border-top", ".5pt solid windowtext");
                cell.Style.Add("border-right", ".5pt solid windowtext");
                cell.Style.Add("border-bottom", ".5pt solid windowtext");

                cell.InnerHtml = objItem.ToString() == "" ? "&nbsp" : objItem.ToString();
                row.Cells.Add(cell);
                i++;
                if (i == 3) Donate_Amt += Convert.ToDouble(objItem.ToString().Replace(",",""));
            }
            table.Rows.Add(row);
            cnt++;

        }

        //轉成 html 碼
        StringWriter sw = new StringWriter();

        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);


        GridList.Text = "線上單筆奉獻查詢<br/><br/>" + "列印日期：" + DateTime.Now.ToLocalTime().ToString() + "<br/><br/>" + htw.InnerWriter.ToString();
        GridList.Text += String.Format("<p align='left'>捐款筆數：{0:#,0} 筆 / 捐款金額合計：{1:#,0} 元</p>", cnt, Donate_Amt);

    }

}
