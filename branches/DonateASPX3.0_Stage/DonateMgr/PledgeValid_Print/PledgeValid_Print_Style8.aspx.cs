﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonateMgr_PledgeValid_Print_PledgeValid_Print_Style8 : BasePage
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
        css.Add("width", "201.54mm");
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
            for (int i = 0; i < count; i = i + 3)
            {

                row = new HtmlTableRow();
                cell = new HtmlTableCell();
                dr = dt.Rows[i];
                strTemp = dr["郵遞區號"].ToString() + dr["地址"].ToString() + "</br>" + dr["捐款人"].ToString() + "  " + dr["稱謂"].ToString() + "啟</br>";
                cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                css = cell.Style;
                css.Add("font-size", strFontSize);
                css.Add("padding-left", "5mm");
                css.Add("padding-right", "4mm");
                css.Add("padding-bottom", "10pt");
                css.Add("width", "70mm");
                css.Add("height", "20mm");
                css.Add("text-align", "left");
                row.Cells.Add(cell);

                //--------------------
                if (i == count - 1 && count % 3 == 1)
                {
                    table.Rows.Add(row);

                    break;
                }

                cell = new HtmlTableCell();
                dr = dt.Rows[i + 1];
                strTemp = dr["郵遞區號"].ToString() + dr["地址"].ToString() + "</br>" + dr["捐款人"].ToString() + "  " + dr["稱謂"].ToString() + "啟</br>";
                cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                css = cell.Style;
                css.Add("font-size", strFontSize);
                css.Add("padding-left", "4mm");
                css.Add("padding-right", "4mm");
                css.Add("padding-bottom", "10pt");
                css.Add("width", "70mm");
                css.Add("height", "20mm");
                css.Add("text-align", "left");
                row.Cells.Add(cell);

                //--------------------
                if (i == count - 2 && count % 3 == 2)
                {
                    table.Rows.Add(row);

                    break;
                }

                cell = new HtmlTableCell();
                dr = dt.Rows[i + 2];
                strTemp = dr["郵遞區號"].ToString() + dr["地址"].ToString() + "</br>" + dr["捐款人"].ToString() + "  " + dr["稱謂"].ToString() + "啟</br>";
                cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                css = cell.Style;
                css.Add("font-size", strFontSize);
                css.Add("padding-left", "4mm");
                css.Add("padding-right", "4mm");
                css.Add("padding-bottom", "10pt");
                css.Add("width", "70mm");
                css.Add("height", "20mm");
                css.Add("text-align", "left");
                row.Cells.Add(cell);

                table.Rows.Add(row);

                row = new HtmlTableRow();
                cell = new HtmlTableCell();
                strTemp = " ";
                cell.InnerHtml = strTemp;
                css = cell.Style;
                css.Add("height", "3mm");
                row.Cells.Add(cell);
                table.Rows.Add(row);

                page_row++;
                if (page_row % 10 == 0)
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
        }
    }
}