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

public partial class MemberMgr_MemberStatQry_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        string strSql = Session["strSql"].ToString();
        string Report_condition = Session["Condition"].ToString();

        Print_Excel(strSql, Report_condition);
    }
    //---------------------------------------------------------------------------
    private void Print_Excel(string strSql, String Report_condition)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            ShowSysMsg("查無資料!!!");
            return;
        }
        GridList.Text = GetTitle(dt.Rows[0], Report_condition) + GetTable1(dt.Rows[0], strSql, Report_condition);
        //GetTableFooter();
        Util.OutputTxt(GridList.Text, "1", "member_stat");
    }
    //---------------------------------------------------------------------------
    private string GetTitle(DataRow dr, String Report_condition)
    {
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

        string strFontSize = "24px";

        string strTitle = "<span style='font-family: 標楷體'>讀者" + Condition_Chs + "統計</span><br>" + "<span style='font-size: 9pt; font-family: 新細明體'>統計項目：" + Condition_Chs + "<br>";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 3;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    //---------------------------------------------------------------------------
    private string GetTable1(DataRow dr, string strSql, String Report_condition)
    {
        //組 table
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;

        DataTable dtRet = CaseUtil.MemberReport_Condition(dt, Report_condition);

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

            if (iCtrl == 0)
            {
                cell.Width = "30";
            }
            else
            {
                cell.Width = "90mm";
            }
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