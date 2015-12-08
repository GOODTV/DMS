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

public partial class MemberMgr_MemberStatQry_Print : BasePage
{
    string TableWidth = "800px";

    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            LoadFormData();
        }
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql = Session["strSql"].ToString();
        string Report_condition = Session["Condition"].ToString();

        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;

        DataTable dtRet = CaseUtil.MemberReport_Condition(dt, Report_condition);

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "16px");
        css.Add("font-family", "標楷體");
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
            if (iCtrl == 0)
            {
                cell.Width = "250mm";
            }
            else
            {
                cell.Width = "250mm";
            }
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);


        //put data
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

        string Condition_Chs = "";
        if (Report_condition == "1")
        {
            Condition_Chs = "類別";
        }
        else if (Report_condition == "2")
        {
            Condition_Chs = "性別";
        }
        else if (Report_condition == "3")
        {
            Condition_Chs = "年齡";
        }
        else if (Report_condition == "4")
        {
            Condition_Chs = "教育程度 ";
        }
        else if (Report_condition == "5")
        {
            Condition_Chs = "職業別";
        }
        else if (Report_condition == "6")
        {
            Condition_Chs = "婚姻狀況";
        }
        else if (Report_condition == "7")
        {
            Condition_Chs = "宗教信仰";
        }
        else if (Report_condition == "8")
        {
            Condition_Chs = "通訊縣市";
        }
        else if (Report_condition == "9")
        {
            Condition_Chs = "狀態 ";
        }
        GridList.Text = GetTitle(Condition_Chs) + GetCondition(Condition_Chs) + htw.InnerWriter.ToString();
    }
    private string GetTitle(string Condition_Chs)
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
        string strTitle = "讀者" + Condition_Chs + "統計<br/>";
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
    private string GetCondition(string Condition_Chs)
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
        css.Add("width", "80%");
        css.Add("line-height", "20px");

        string strTitle = "統計項目：" + Condition_Chs + "<br>";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "left");
        row.Cells.Add(cell);
        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
}