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

public partial class EmailMgr_ThankList_Print_ThankList_Print_Style6 : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Print();
        if (GridList.Text != "** 沒有符合條件的資料 **")
        {
            Response.Write("<script language=javascript>window.print();</script>");
        }
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
        //data
        int count = dt.Rows.Count;
        if (count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
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
                css.Add("padding-right", " 5mm");
                css.Add("padding-top", "15pt");
                css.Add("padding-bottom", "10pt");
                css.Add("width", "96mm");
                css.Add("height", "10mm");
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
                css.Add("padding-right", " 5mm");
                css.Add("padding-top", "10pt");
                css.Add("padding-bottom", "10pt");
                css.Add("width", "96mm");
                css.Add("height", "10mm");
                css.Add("text-align", "left");
                row.Cells.Add(cell);

                table.Rows.Add(row);

                if (page_row != 0)
                {
                    row = new HtmlTableRow();
                    cell = new HtmlTableCell();
                    cell.InnerHtml = " ";
                    css = cell.Style;
                    css.Add("height", "1mm");
                    row.Cells.Add(cell);
                    table.Rows.Add(row);
                }
                ////else
                ////{
                ////    row = new HtmlTableRow();
                ////    cell = new HtmlTableCell();
                ////    cell.InnerHtml = " ";
                ////    css = cell.Style;
                ////    css.Add("height", "5mm");
                ////    row.Cells.Add(cell);
                ////    table.Rows.Add(row);
                ////}

                page_row++;
                if (page_row % 11 == 0)
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

            //更新Donor的IsThanks_Add欄位
            if (strSql.Contains("dr.IsThanks_Add = '0' or dr.IsThanks_Add is null") == true)
            {
                for (int i = 0; i < count; i++)
                {
                    dr = dt.Rows[i];
                    //Edit
                    //****變數宣告****//
                    Dictionary<string, object> dict = new Dictionary<string, object>();

                    //****設定SQL指令****//
                    string strSql_Donor = " update Donor set ";

                    strSql_Donor += "  IsThanks_Add = @IsThanks_Add";
                    strSql_Donor += " where Donor_Id = @Donor_Id";

                    dict.Add("IsThanks_Add", "1");
                    dict.Add("Donor_Id", dr["捐款人編號"].ToString());
                    NpoDB.ExecuteSQLS(strSql_Donor, dict);
                }
            }
        }
    }
}