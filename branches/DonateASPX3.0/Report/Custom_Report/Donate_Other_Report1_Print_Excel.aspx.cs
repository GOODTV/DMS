using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;
using System.Web.UI.HtmlControls;

public partial class Report_Custom_Report_Donate_Other_Report1_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_DonateDateS_other1.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_other1.Value = Util.GetQueryString("DonateDateE");
            Print_Excel();
        }
    }
    //---------------------------------------------------------------------------
    private void Print_Excel()
    {
        string strSql = Sql();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            ShowSysMsg("查無資料!!!");
            Response.Write("<script language=javascript>window.close();</script>");
            return;
        }

        GridList.Text = GetTitle(dt.Rows[0]) + GetTable1(dt.Rows[0], strSql);
        //GetTableFooter();
        Util.OutputTxt(GridList.Text, "1", "donate_other_report1");
    }
    //---------------------------------------------------------------------------
    private string GetTitle(DataRow dr)
    {
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

        string strFontSize = "20px";
        string strTitle = "人數統計分析<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + HFD_DonateDateS_other1.Value + "~" + HFD_DonateDateE_other1.Value + "</span><br>";
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 7;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
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

        DataTable dtRet = CaseUtil.Donate_Other_Report1_Print(dt);

        table.Border = 0;
        table.CellPadding = 0;
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
            //cell.Style.Add("border-left", ".5pt solid windowtext");
            cell.Style.Add("border-top", ".5pt solid windowtext");
            //cell.Style.Add("border-right", ".5pt solid windowtext");
            cell.Style.Add("border-bottom", ".5pt solid windowtext");

            if (iCtrl == 0)
            {
                cell.Width = "70mm";
                cell.Height = "40mm";

                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-left", ".5pt solid windowtext");
            }
            else if (iCtrl == 1)
            {
                //cell.Width = "50mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 2)
            {
                //cell.Width = "80mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 3)
            {
                //cell.Width = "60mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 4)
            {
                //cell.Width = "60mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 5)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 6)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-right", ".5pt solid windowtext");
            }
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

        foreach (DataRow drow in dtRet.Rows)
        {
            int Col = 0;
            row = new HtmlTableRow();
            foreach (object objItem in drow.ItemArray)
            {
                cell = new HtmlTableCell();
                //cell.Style.Add("border-left", ".5pt solid windowtext");
                cell.Style.Add("border-top", ".5pt solid windowtext");
                //cell.Style.Add("border-right", ".5pt solid windowtext");
                cell.Style.Add("border-bottom", ".5pt solid windowtext");

                cell.InnerHtml = objItem.ToString() == "" ? "&nbsp" : objItem.ToString();
                if (Col == 0)
                {
                    cell.Style.Add("border-left", ".5pt solid windowtext");
                    cell.Style.Add("text-align", "center");
                }
                else if (Col == 6)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                    cell.Style.Add("text-align", "center");
                }
                else
                {
                    cell.Style.Add("text-align", "center");
                }
                row.Cells.Add(cell);
                Col++;
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
    private string Sql()
    {
        string strSql1 = @" select COUNT(DISTINCT D.Donor_Id) AS '所有奉獻天使人數' ,COUNT(D1.Donor_Id) AS '累計奉獻10萬以上總人數' 
                               ,Convert(VARCHAR , COUNT(D2.Donor_Id)) AS '累計奉獻30萬以上總人數' ,COUNT(D3.Donor_Id) AS '累計奉獻50萬以上總人數' 
                               ,COUNT(D4.Donor_Id) AS '累計奉獻100萬以上總人數' , COUNT(M1.Donor_Type) AS '友好教會及機構數量'
                         from Donor M 
                         LEFT JOIN 
                               (SELECT Donor_Id, COUNT(DISTINCT Donor_Id) AS '奉獻人數' ,COUNT(Donate_Id) AS '奉獻筆數' ,SUM(Donate_Amt) AS '奉獻總金額'
                               FROM dbo.DONATE 
                               WHERE ISNULL(Issue_Type,'') != 'D' ";
        string strSql2 = @" GROUP BY Donor_Id )D  
                             ON M.Donor_Id = D.Donor_Id 
                         LEFT JOIN (SELECT Donor_Id  ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times 
                         FROM dbo.DONATE  
                         WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Amt >0 ";
        string strSql3 = @" GROUP BY Donor_Id Having Sum(Donate_Amt) >= 100000 ) D1 
                            ON M.Donor_Id = D1.Donor_Id 
                         LEFT JOIN (SELECT Donor_Id  ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times 
                         FROM dbo.DONATE 
                         WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Amt >0 ";
        string strSql4 = @" GROUP BY Donor_Id Having Sum(Donate_Amt) >= 300000 ) D2 
                             ON M.Donor_Id = D2.Donor_Id
                         LEFT JOIN (SELECT Donor_Id  ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times 
                         FROM dbo.DONATE 
                         WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Amt >0 ";
        string strSql5 = @" GROUP BY Donor_Id Having Sum(Donate_Amt) >= 500000 ) D3 
                             ON M.Donor_Id = D3.Donor_Id 
                         LEFT JOIN (SELECT Donor_Id  ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times 
                         FROM dbo.DONATE 
                         WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Amt >0 ";
        string strSql6 = @" GROUP BY Donor_Id Having Sum(Donate_Amt) >= 1000000 ) D4 
                             ON M.Donor_Id = D4.Donor_Id 
                         LEFT JOIN (SELECT distinct Dt.Donor_Id,Dr.Donor_type  FROM dbo.DONATE Dt 
                                    inner join Donor Dr on Dr.Donor_Id = Dt.Donor_Id
                         WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Amt >0  And Dr.Donor_type IN ('教會','福音機構') ";
        string strSql7 = @" ) M1
                         ON M.Donor_Id = M1.Donor_Id ";

        //以下是串UNION的SQL語法
        /*string strSql7 = @" union
                               SELECT COUNT(DISTINCT Donor_Id) AS '奉獻人數' ,COUNT(Donate_Id) AS '奉獻筆數' , REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,SUM(Donate_Amt)),1),'.00','') AS '奉獻總金額','','',''
                               FROM dbo.DONATE 
                               WHERE Donate_Purpose <> '' ";*/
        //搜尋條件
        string Donate_Where = "";
        if (HFD_DonateDateS_other1.Value != "")
        {
            Donate_Where += " And Donate_Date>='" + HFD_DonateDateS_other1.Value + "' ";
        }
        if (HFD_DonateDateE_other1.Value != "")
        {
            Donate_Where += " And Donate_Date<='" + HFD_DonateDateE_other1.Value + "' ";
        }
        string strSql = "";
        if (Donate_Where != "")
        {
            strSql = strSql1 + Donate_Where + strSql2 + Donate_Where + strSql3 + Donate_Where + strSql4 + Donate_Where + strSql5 + Donate_Where + strSql6 + Donate_Where + strSql7;
            //strSql = strSql1 + Donate_Where + strSql2 + Donate_Where + strSql3 + Donate_Where + strSql4 + Donate_Where + strSql5 + Donate_Where + strSql6 + strSql7 + Donate_Where;
        }
        else
        {
            strSql = strSql1 + strSql2 + strSql3 + strSql4 + strSql5 + strSql6 + strSql7;
            //strSql = strSql1 + strSql2 + strSql3 + strSql4 + strSql5 + strSql6 + strSql7;
        }
        return strSql;
    }
}