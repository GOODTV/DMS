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

public partial class Report_Custom_Report_Donate_Week_Report1_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Donate_Purpose_week1.Value = Util.GetQueryString("Donate_Purpose");
            HFD_DonateDateS_week1.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_week1.Value = Util.GetQueryString("DonateDateE");
            PrintReport();
        }
    }
    //---------------------------------------------------------------------------
    private void PrintReport()
    {
        string strSql = Sql();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;
        if (dt.Rows.Count == 1)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
            return;
        }

        DataTable dtRet = CaseUtil.Donate_Week_Report1_Print(dt);

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
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
                cell.Width = "80mm";
                cell.Height = "40mm";
                cell.Style.Add("border-left", ".5pt solid windowtext");
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 1)
            {
                cell.Width = "100mm";
                cell.Style.Add("text-align", "right");
            }
            else if (iCtrl == 2)
            {
                cell.Width = "180mm";
                cell.Style.Add("text-align", "right");
            }
            else if (iCtrl == 3)
            {
                cell.Width = "180mm";
                cell.Style.Add("border-right", ".5pt solid windowtext");
                cell.Style.Add("text-align", "right");
            }
           
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

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
                    cell.Style.Add("text-align", "right");
                }
                else
                {
                    cell.Style.Add("text-align", "right");
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
        if (HFD_Donate_Purpose_week1.Value == "建台")
        {
            GridList.Text = "建台奉獻級距表<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + HFD_DonateDateS_week1.Value + "~" + HFD_DonateDateE_week1.Value + "</span><br>" + htw.InnerWriter.ToString();
        }
        else
        {
            GridList.Text = "非建台奉獻級距表<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + HFD_DonateDateS_week1.Value + "~" + HFD_DonateDateE_week1.Value + "</span><br>" + htw.InnerWriter.ToString();
        }
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql = "";
        //搜尋條件
        string Donate_Where = "";
        if (HFD_DonateDateS_week1.Value != "")
        {
            Donate_Where = " And Donate_Date>='" + HFD_DonateDateS_week1.Value + "'";
        }
        if (HFD_DonateDateS_week1.Value != "")
        {
            Donate_Where += " And Donate_Date<='" + HFD_DonateDateE_week1.Value + "'";
        }
        if (HFD_Donate_Purpose_week1.Value == "建台")
        {
            strSql = @" SELECT '1' as '序號', '1萬(不含)以下' as '級距',COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0)),1),'.00','') as '小計'   
                        FROM dbo.DONATE   
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose = '建台'" + Donate_Where + @"
                              AND Donate_Amt Between 1 and 9999  
                        UNION  
                        SELECT '2' as '序號', '1萬(含)以上' as '級距', COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0)),1),'.00','') as '小計'  
                        FROM dbo.DONATE  
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose = '建台'" + Donate_Where + @"
	                          AND Donate_Amt Between 10000 and 99999  
                        UNION  
                        SELECT '3' as '序號', '10萬(含)以上' as '級距', COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0)),1),'.00','') as '小計'  
                        FROM dbo.DONATE  
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose = '建台'" + Donate_Where + @"
	                          AND Donate_Amt Between 100000 and 999999  
                        UNION  
                        SELECT '4' as '序號', '100萬(含)以上' as '級距', COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0)),1),'.00','') as '小計'  
                        FROM dbo.DONATE  
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose = '建台'" + Donate_Where + @"
	                          AND Donate_Amt Between 1000000 and 9999999  
                        UNION  
                        SELECT '5' as '序號', '1000萬(含)以上' as '級距', COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0)),1),'.00','') as '小計'  
                        FROM dbo.DONATE  
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose = '建台'" + Donate_Where + @"
	                          AND Donate_Amt >= 10000000  
                        UNION  
                        SELECT '6' as '序號', '總計' as '級距', COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0)),1),'.00','') as '總金額'  
                        FROM dbo.DONATE  
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose = '建台'" + Donate_Where + @"
	                          AND Donate_Amt > 0 
                        UNION  
                        Select '列印日期：' as '序號' , CONVERT(VARCHAR, GETDATE(), 120 )  as '級距', '' as '筆數' ,'頁數：1/1' as '總金額' ";
        }
        else
        {
            strSql = @" SELECT '1' as '序號', '1萬(不含)以下' as '級距',COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0)),1),'.00','') as '小計'   
                        FROM dbo.DONATE   
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose <> '建台'" + Donate_Where + @"
                              AND Donate_Amt Between 1 and 9999  
                        UNION  
                        SELECT '2' as '序號', '1萬(含)以上' as '級距', COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0)),1),'.00','') as '小計'  
                        FROM dbo.DONATE  
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose <> '建台'" + Donate_Where + @"
	                          AND Donate_Amt Between 10000 and 99999  
                        UNION  
                        SELECT '3' as '序號', '10萬(含)以上' as '級距', COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0)),1),'.00','') as '小計'  
                        FROM dbo.DONATE  
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose <> '建台'" + Donate_Where + @"
	                          AND Donate_Amt Between 100000 and 999999  
                        UNION  
                        SELECT '4' as '序號', '100萬(含)以上' as '級距', COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0)),1),'.00','') as '小計'  
                        FROM dbo.DONATE  
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose <> '建台'" + Donate_Where + @"
	                          AND Donate_Amt Between 1000000 and 9999999  
                        UNION  
                        SELECT '5' as '序號', '1000萬(含)以上' as '級距', COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(int,SUM(Donate_Amt)),1),0)),1),'.00','') as '小計'  
                        FROM dbo.DONATE  
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose <> '建台'" + Donate_Where + @"
	                          AND Donate_Amt >= 10000000  
                        UNION  
                        SELECT '6' as '序號', '總計' as '級距', COUNT(*) as '筆數' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,ISNUll(Convert(varchar,Convert(varchar,SUM(Donate_Amt)),1),0)),1),'.00','') as '總金額'  
                        FROM dbo.DONATE  
                        WHERE ISNULL(Issue_Type,'') != 'D' and Donate_Purpose <> '建台'" + Donate_Where + @"
	                          AND Donate_Amt > 0 
                        UNION  
                        Select '列印日期：' as '序號' , CONVERT(VARCHAR, GETDATE(), 120 )  as '級距', '' as '筆數' ,'' as '總金額' ";
        }
        return strSql;
    }
}