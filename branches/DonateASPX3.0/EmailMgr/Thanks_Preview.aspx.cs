﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class EmailMgr_Thanks_Preview : BasePage
{
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        Print();
    }
    private void Print()
    {
        Dictionary<string, int> ReportData = new Dictionary<string, int>();
        GridList.Text = GetTable(ReportData);
    }
    private string GetTable(Dictionary<string, int> ReportData)
    {
        //撈出信件內容
        string strSql2 = "select * from EmailMgr where EmailMgr_Subject='" + Session["Email_Subject"].ToString() + "'";
        DataTable dt2 = NpoDB.QueryGetTable(strSql2);
        DataRow dr2;
        dr2 = dt2.Rows[0];

        //組 table
        string strTemp = "";
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;
        table.Style.Add("width", "194mm");
        table.Style.Add("font-size", "13px");
        css = table.Style;
        //--------------------------------------------
        //data

        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        strTemp = dr2["EmailMgr_Desc"].ToString();
        //代換成個資
        if (Session["Email_Subject"].ToString() == "首捐感謝函")
        {
            strTemp = strTemp.Replace("○○○", "加百列");
        }
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        row.Cells.Add(cell);
        table.Rows.Add(row);
        //--------------------------------------------

        //轉成 html 碼
        StringWriter sw = new StringWriter();

        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    //end of GetTable()
}