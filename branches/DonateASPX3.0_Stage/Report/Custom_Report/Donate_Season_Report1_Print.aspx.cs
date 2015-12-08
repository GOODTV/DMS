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

public partial class Report_Custom_Report_Donate_Season_Report1_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_DonateDateS_season1.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_season1.Value = Util.GetQueryString("DonateDateE");
            PrintReport();
        }
    }
    //---------------------------------------------------------------------------
    private void PrintReport()
    {
        string strSql = Sql();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 1)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
            return;
        }

        DataTable dtRet = CaseUtil.Donate_Season_Report1_Print(dt);

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        //table.Width = "770";
        css = table.Style;
        css.Add("font-size", "12px");
        css.Add("font-family", "細明體");
        css.Add("line-height", "20px");
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
                //cell.Width = "55mm";
                cell.Height = "40mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-left", ".5pt solid windowtext");
            }
            else if (iCtrl == 1)
            {
                cell.Width = "80mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 2)
            {
                cell.Width = "80mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 3)
            {
                cell.Width = "80mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-right", ".5pt solid windowtext");
            }
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);
        int count = 0;
        foreach (DataRow dr in dtRet.Rows)
        {
            int Col = 0;
            row = new HtmlTableRow();

            foreach (object objItem in dr.ItemArray)
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
                else if (Col == 3)
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
            //印出大小不確定,先不強制分頁
            ////size 45 
            //count++;
            //if (count % 45 == 0)
            //{
            //    cell = new HtmlTableCell();
            //    string strTemp = "<br clear=all style='page-break-before:always'>";
            //    cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            //    css = cell.Style;
            //    row.Cells.Add(cell);
            //}
            table.Rows.Add(row);
        }
        //轉成 html 碼
        StringWriter sw = new StringWriter();

        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        GridList.Text = DateTime.Now.AddYears(-1911).Year + "年" + DateTime.Now.Month + "月份 非建台奉獻統計分析表<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + HFD_DonateDateS_season1.Value + "~" + HFD_DonateDateE_season1.Value + "</span><br>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql = @"  SELECT ROW_NUMBER() OVER(ORDER BY COUNT(DISTINCT Donor_Id)) AS ROWID, CONVERT(VARCHAR,COUNT(DISTINCT Donor_Id)) AS '人數' ,COUNT(Donate_Id) AS '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,SUM(Donate_Amt)),1),'.00','') AS '總金額' 
                            FROM dbo.DONATE 
                            WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose <> '建台' ";

        //搜尋條件
        if (HFD_DonateDateS_season1.Value != "" )
        {
            strSql += " And Donate_Date>='" + HFD_DonateDateS_season1.Value + "' ";
        }
        if (HFD_DonateDateE_season1.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_season1.Value + "' ";
        }
        //以下是串UNION的SQL語法
        strSql += @" Union 
					 select '99999','列印日期：','',CONVERT(VARCHAR, GETDATE(), 120 )";
        return strSql;
    }
}