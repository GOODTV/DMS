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

public partial class DonateMgr_Invoice_Print_Invoice_Print_Style1 : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Print();
    }
    //---------------------------------------------------------------------------
    private void Print()
    {
        string strSql = Session["strSql_Print"].ToString();
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr;

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
        css.Add("font-size", "13px");
        css.Add("font-family", "標楷體");
        css.Add("width", "194mm");
        css.Add("border-collapse", "collpase");
        string strFontSize = "12pt";
        //--------------------------------------------
        int count = dt.Rows.Count;
        if (count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
            //data
            int page_row = 0;
            for (int i = 0; i < count; i = i + 2)
            {

                row = new HtmlTableRow();
                cell = new HtmlTableCell();
                dr = dt.Rows[i];
                strTemp = dr["郵遞區號"].ToString() + "</br>" + dr["地址"].ToString() + "</br>" + dr["捐款人"].ToString() + "  " + dr["稱謂"].ToString() + "啟</br>";
                cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                css = cell.Style;
                css.Add("font-size", strFontSize);
                css.Add("padding-left", "5mm");
                css.Add("padding-right", " 10mm");
                css.Add("padding-top", "10pt");
                css.Add("padding-bottom", "10pt");
                css.Add("width", "96mm");
                css.Add("height", "28mm");
                css.Add("text-align", "left");
                row.Cells.Add(cell);

                //--------------------
                if (i == count - 1 && count % 2 == 1)
                {
                    table.Rows.Add(row);
                    break;
                }

                cell = new HtmlTableCell();
                dr = dt.Rows[i + 1];
                strTemp = dr["郵遞區號"].ToString() + "</br>" + dr["地址"].ToString() + "</br>" + dr["捐款人"].ToString() + "  " + dr["稱謂"].ToString() + "啟</br>";
                cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                css = cell.Style;
                css.Add("font-size", strFontSize);
                css.Add("padding-left", "5mm");
                css.Add("padding-right", " 10mm");
                css.Add("padding-top", "10pt");
                css.Add("padding-bottom", "10pt");
                css.Add("width", "96mm");
                css.Add("height", "28mm");
                css.Add("text-align", "left");
                row.Cells.Add(cell);

                table.Rows.Add(row);

                if (page_row != 0)
                {
                    row = new HtmlTableRow();
                    cell = new HtmlTableCell();
                    strTemp = " ";
                    cell.InnerHtml = strTemp;
                    css = cell.Style;
                    css.Add("height", "3mm");
                    row.Cells.Add(cell);
                    table.Rows.Add(row);
                }

                page_row++;
                if (page_row % 7 == 0)
                {
                    cell = new HtmlTableCell();
                    strTemp = "<br clear=all style='page-break-before:always'>";
                    cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                    css = cell.Style;
                    row.Cells.Add(cell);
                }
            }
            //--------------------------------------------

            //轉成 html 碼
            StringWriter sw = new StringWriter();

            HtmlTextWriter htw = new HtmlTextWriter(sw);
            table.RenderControl(htw);

            GridList.Text = htw.InnerWriter.ToString();

            //Edit_Donate
            string strSql_Edit = Session["strSql_Edit"].ToString();
            if (strSql_Edit.Contains("do.Invoice_Print_Add = '0' or do.Invoice_Print_Add is null") == true)
            {
                DataTable dt2 = NpoDB.QueryGetTable(strSql_Edit);
                DataRow dr2;
                int count2 = dt2.Rows.Count;
                for (int i = 0; i < count2; i++)
                {
                    dr2 = dt2.Rows[i];
                    //Edit
                    //****變數宣告****//
                    Dictionary<string, object> dict = new Dictionary<string, object>();

                    //****設定SQL指令****//
                    string strSql_Donate = " update Donate set ";

                    strSql_Donate += "  Invoice_Print_Add = @Invoice_Print_Add";
                    strSql_Donate += " where Donate_Id = @Donate_Id";

                    dict.Add("Invoice_Print_Add", "1");
                    dict.Add("Donate_Id", dr2["捐款編號"].ToString());
                    NpoDB.ExecuteSQLS(strSql_Donate, dict);
                }
            }
        }
    }
}