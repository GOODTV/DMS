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

public partial class DonateMgr_VerifyList_Print_Excel_style2 : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";
    protected void Page_Load(object sender, EventArgs e)
    {
        Print_Excel();
    }
    //---------------------------------------------------------------------------
    private void Print_Excel()
    {
        string strSql = @"SELECT REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,m.Donate_Amt),1),'.00','') ,
                              left(m.Donor_Names,len(m.Donor_Names)-1) as Donor_NamesFinal from
                                  (SELECT Donate_Amt,(SELECT cast(Donor_Name AS NVARCHAR )+ '(' +cast(MONTH(Donate_Date)AS NVARCHAR )+'/'+cast(Day(Donate_Date)AS NVARCHAR ) +')' + '、' 
                                  from Donate left join Donor on Donate.Donor_Id = Donor.Donor_Id
                                  where Donate_Amt = Do.Donate_Amt " + Session["Donor_Name"];
        strSql += @"          FOR XML PATH('')) as Donor_Names
                              from Donate Do left join Donor Dr on Do.Donor_Id = Dr.Donor_Id
                              where dr.DeleteDate is null ";
        strSql += Session["Condition"];
        strSql += @"     GROUP BY Donate_Amt) M --這個M一定要加，不知道為啥
                     ORDER by M.Donate_Amt desc";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        if (dt.Rows.Count == 0)
        {
            ShowSysMsg("查無資料!!!");
            return;
        }
        GridList.Text = GetTitle() + GetTime() + GetTable1(dt.Rows[0], strSql) + "<br/>" + GetEnd();
        //GetTableFooter();
        Util.OutputTxt(GridList.Text, "1", "Verify");

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
        css.Add("font-size", "13pt");
        css.Add("font-family", "標楷體");
        css.Add("width", TableWidth);

        string strTitle = "徵信芳名錄<br/> ";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        cell.ColSpan = 2;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    //---------------------------------------------------------------------------
    private string GetTime()
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
        css.Add("width", TableWidth);

        string strTitle = "匯出日期：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "<br/> ";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "left");
        cell.ColSpan = 2;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    //---------------------------------------------------------------------------
    private string GetTable1(DataRow dr, string strSql)
    {
        //組 table
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;

        DataTable dtRet = CaseUtil.VerifyList_Print_style2(dt);

        table.Border = 1;
        table.CellPadding = 3;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "9pt");
        css.Add("font-family", "新細明體");
        css.Add("width", TableWidth);
        css.Add("line-height", "40px");

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

            //cell.Width = "100mm";
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
    //---------------------------------------------------------------------------
    private string GetEnd()
    {
        string strSql = @"SELECT REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(do.Donate_Amt)),1),'.00','') as [Donate_Amt]
                          FROM Donate do left join Donor dr on do.Donor_Id = dr.Donor_id where dr.DeleteDate is null ";
        strSql += Session["Condition"];
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];

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
        css.Add("width", TableWidth);

        string strTitle = "捐款總金額：" + dr["Donate_Amt"];
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "left");
        cell.ColSpan = 2;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
}