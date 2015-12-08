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

public partial class Report_Custom_Report_Donate_Year_Report2_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Year_year2.Value = Util.GetQueryString("Year");
            HFD_Is_Abroad_year2.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_year2.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_year2.Value = Util.GetQueryString("Sex");
            Print_Excel();
        }
    }
    //---------------------------------------------------------------------------
    private void Print_Excel()
    {
        string strSql = Sql();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 1)
        {
            ShowSysMsg("查無資料!!!");
            Response.Write("<script language=javascript>window.close();</script>");
            return;
        }

        GridList.Text = GetTitle(dt.Rows[0]) + GetTable1(dt.Rows[0], strSql);
        //GetTableFooter();
        Util.OutputTxt(GridList.Text, "1", "donate_year_report2");
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
        string strTitle = HFD_Year_year2.Value + "年 捐款總額與月刊索取明細表<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + DateTime.Now.ToString() + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依累計奉獻金額排序</span>";
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 11;
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

        DataTable dtRet = CaseUtil.Donate_Year_Report2_Print(dt);

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
                //cell.Width = "70mm";
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
            }
            else if (iCtrl == 7)
            {
                //cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 8)
            {
                //cell.Width = "85mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 9)
            {
                //cell.Width = "85mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 10)
            {
                //cell.Width = "40mm";
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
                else if (Col == 10)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                    cell.Style.Add("text-align", "center");
                }
                else if (Col == 7)
                {
                    cell.Style.Add("text-align", "left");
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
        string strSql = @"  SELECT D.Donor_Id AS '編號'   
                                    ,M.Donor_Name AS '天使姓名' ,M.Title AS '稱謂'  
								    ,M.Donor_Type AS '身分別' ,ISNULL(M.IDNo,'') AS '身份證或統編' 
								    ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Sum_Amt),1),'.00','') AS '累計奉獻金額' ,D.Times AS '奉獻次數'  
								    ,IsNull(IsSendNewsNum,0) AS '月刊份數'  
								    ,Case When M.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'')  
										  else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
								    ,CONVERT(VARCHAR, Begin_DonateDate, 111) AS '首捐日期' ,CONVERT(VARCHAR, Last_DonateDate, 111) AS '末捐日期' ,D.Sum_Amt   
									FROM dbo.DONOR M  
									INNER JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
									            FROM DONATE 
                                                WHERE ISNULL(Issue_Type,'') != 'D' ";

        //搜尋條件
        if (HFD_Year_year2.Value != "")
        {
            strSql += " and DATEPART(YEAR,Donate_Date) = '" + HFD_Year_year2.Value + "' GROUP BY Donor_Id) D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id) D ";
        }
        strSql += @" ON M.Donor_Id = D.Donor_Id 
                     LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
                     LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode
                     Where 1=1 ";
        if (HFD_Is_Abroad_year2.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_year2.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_year2.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        //以下是串UNION的SQL語法
        strSql += @" Union select count(D.Donor_Id),REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(D.Sum_Amt)),1),'.00',''),'','','','0','','','','','',sum(D.Sum_Amt)
                                 FROM dbo.DONOR M  
                                 INNER JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
                                             FROM DONATE 
                                             WHERE ISNULL(Issue_Type,'') != 'D' ";
        //搜尋條件
        if (HFD_Year_year2.Value != "")
        {
            strSql += " and DATEPART(YEAR,Donate_Date) = '" + HFD_Year_year2.Value + "' GROUP BY Donor_Id) D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id) D ";
        }
        strSql += " ON M.Donor_Id = D.Donor_Id  Where 1=1 ";
        if (HFD_Is_Abroad_year2.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_year2.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_year2.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        strSql += " Order by D.Sum_Amt ";
        return strSql;
    }
}