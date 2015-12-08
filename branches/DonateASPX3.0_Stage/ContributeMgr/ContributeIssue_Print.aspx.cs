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
    string Issue_Id = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Issue_Id = Util.GetQueryString("Issue_Id");
        }
        PrintReport();
    }
    //---------------------------------------------------------------------------
    private void PrintReport()
    {
        string strSql = "select Goods_Id, Goods_Name, convert(nvarchar,Goods_Qty) + Goods_Unit,Goods_Comment from Contribute_IssueData where Issue_Id='" + Issue_Id + "'";
        Dictionary<string, object> dict = new Dictionary<string, object>();

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
        }

        DataTable dtRet = CaseUtil.ContributeIssue_Print(dt);

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

            if (iCtrl == 0 )
            {
                cell.Width = "80mm";
            }
            else if (iCtrl == 1)
            {
                cell.Width = "300mm";
            }
            else if (iCtrl == 2)
            {
                cell.Width = "120mm";
            }
            else if (iCtrl == 3)
            {
                cell.Width = "250mm";
            }
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
        Dictionary<string, int> ReportData = new Dictionary<string, int>();
        GridList.Text = GetTitle()+GetTable(ReportData) + htw.InnerWriter.ToString()+GetEnd();
    }
    //---------------------------------------------------------------------------
    private string GetTitle()
    {
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "14pt");
        css.Add("font-family", "標楷體");
        css.Add("width", "100%");

        string strFontSize = "18px";
        string strTitle = "加百列物品領用單<br/> ";
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
    private string GetTable(Dictionary<string, int> ReportData)
    {
        string strSql = "select * from Contribute_Issue where Issue_Id ='" + Issue_Id + "'";
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
        css.Add("font-size", "16px");
        css.Add("font-family", "標楷體");
        css.Add("width", "100%");
        css.Add("border-collapse", "collpase");
        string strFontSize = "12pt";
        //--------------------------------------------
        //data
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        dr = dt.Rows[0];
        strTemp = "領用編號：" + dr["Issue_Pre"].ToString() + dr["Issue_No"].ToString();
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", strFontSize);
        css.Add("width", "60%");
        //css.Add("margin-left", "80px");
        css.Add("height", "25px");
        row.Cells.Add(cell);
        //table.Rows.Add(row);

        //row = new HtmlTableRow();
        cell = new HtmlTableCell();
        dr = dt.Rows[0];
        strTemp = "領用日期：" + DateTime.Parse(dr["Issue_Date"].ToString()).ToString("yyyy年MM月dd日");
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", strFontSize);
        css.Add("width", "40%");
        //css.Add("margin-right", "150px");
        css.Add("height", "25px");
        row.Cells.Add(cell);
        table.Rows.Add(row);

        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        dr = dt.Rows[0];
        strTemp = "領用用途：" + dr["Issue_Purpose"].ToString();
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", strFontSize);
        css.Add("width", "60%");
        //css.Add("line-height", "120%");
        //css.Add("margin-right", "150px");
        css.Add("height", "25px");
        row.Cells.Add(cell);
        //table.Rows.Add(row);


        //row = new HtmlTableRow();
        cell = new HtmlTableCell();
        dr = dt.Rows[0];
        strTemp = "出貨單位：" + dr["Issue_Org"].ToString();
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", strFontSize);
        css.Add("width", "40%");
        //css.Add("margin-right", "150px");
        css.Add("height", "25px");
        row.Cells.Add(cell);
        table.Rows.Add(row);

        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        dr = dt.Rows[0];
        strTemp = "";
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", strFontSize);
        css.Add("width", "60%");
        //css.Add("line-height", "120%");
        //css.Add("margin-right", "150px");
        //css.Add("height", "420px");
        row.Cells.Add(cell);
        //table.Rows.Add(row);

        //row = new HtmlTableRow();
        cell = new HtmlTableCell();
        dr = dt.Rows[0];
        strTemp = "列印時間：" + DateTime.Now.ToLocalTime().ToString();   
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", strFontSize);
        css.Add("width", "40%");
        //css.Add("margin-right", "150px");
        css.Add("height", "20px");
        row.Cells.Add(cell);
        table.Rows.Add(row);
        //--------------------------------------------

        //轉成 html 碼
        StringWriter sw = new StringWriter();

        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    } //end of GetTable()
    private string GetEnd()
    {
        string strTemp = "";
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
        css.Add("width", "100%");
        css.Add("border-collapse", "collpase");
        string strFontSize = "12pt";
        //--------------------------------------------
        //data
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        strTemp = "取貨人：" ;
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", strFontSize);
        css.Add("width", "33%");
        css.Add("height", "40px");
        row.Cells.Add(cell);

        //row = new HtmlTableRow();
        cell = new HtmlTableCell();
        strTemp = "領取人：" ;
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", strFontSize);
        css.Add("width", "33%");
        //css.Add("margin-right", "150px");
        css.Add("height", "40px");
        row.Cells.Add(cell);

        cell = new HtmlTableCell();
        strTemp = "主管簽核：" ;
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", strFontSize);
        css.Add("width", "33%");
        //css.Add("margin-right", "150px");
        css.Add("height", "40px");
        row.Cells.Add(cell);
        table.Rows.Add(row);

        //--------------------------------------------

        //轉成 html 碼
        StringWriter sw = new StringWriter();

        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
}