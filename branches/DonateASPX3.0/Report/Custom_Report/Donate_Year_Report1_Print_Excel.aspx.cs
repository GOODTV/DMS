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

public partial class Report_Custom_Report_Donate_Year_Report1_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Donate_Amt_year1.Value = Util.GetQueryString("Donate_Amt");
            HFD_Donate_Total_Begin_year1.Value = Util.GetQueryString("Donate_Total_Begin");
            HFD_Donate_Total_End_year1.Value = Util.GetQueryString("Donate_Total_End");
            HFD_DonateDateS_year1.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_year1.Value = Util.GetQueryString("DonateDateE");
            HFD_Donate_Purpose_year1.Value = Util.GetQueryString("Donate_Purpose");
            HFD_Is_Abroad_year1.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_year1.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_year1.Value = Util.GetQueryString("Sex");
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
        Util.OutputTxt(GridList.Text, "1", "donate_year_report1");
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
        string strTitle = "國內外捐款用途總額明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計期間：" + HFD_DonateDateS_year1.Value + "~" + HFD_DonateDateE_year1.Value + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人編號排序</span>";
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 10;
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

        DataTable dtRet = CaseUtil.Donate_Year_Report1_Print(dt);

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
            else if (iCtrl == 9)
            {
                //cell.Width = "85mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-right", ".5pt solid windowtext");
            }
            else
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
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
                else if (Col == 9)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                    cell.Style.Add("text-align", "center");
                }
                else if (Col == 6)
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
								    ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Sum_Amt),1),'.00','') AS '累計奉獻金額' ,D.Times AS '奉獻次數'  
								    ,IsNull(M.ZipCode,'') AS '郵遞區號'  
								    ,Case When M.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') 
										  else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址',ISNULL(Email,'') AS 'Email'  
								    ,CONVERT(VARCHAR, Begin_DonateDate, 111) AS '首捐日期' ,CONVERT(VARCHAR, Last_DonateDate, 111) AS '末捐日期'  
							FROM DONOR M  
							RIGHT JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
							            FROM DONATE 
                                        WHERE ISNULL(Issue_Type,'') != 'D' ";

        //搜尋條件
        if (HFD_Donate_Amt_year1.Value != "")
        {
            strSql += " and Donate_Amt>'" + HFD_Donate_Amt_year1.Value + "' ";
        }
        else
        {
            strSql += " and Donate_Amt > 0 ";
        }
        string[] Donate_Purpose;
        int count = 0;
        if (HFD_Donate_Purpose_year1.Value != "")
        {
            char[] ch = new char[] { ',' };
            Donate_Purpose = HFD_Donate_Purpose_year1.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
            count = Donate_Purpose.Length;

            bool first = true; int cnt = 0;
            for (int i = 0; i < count; i++)
            {
                if (first)
                {
                    strSql += " AND (Donate_Purpose='" + Donate_Purpose[i] + "' "; first = false; cnt++;
                }
                else
                {
                    strSql += " or Donate_Purpose='" + Donate_Purpose[i] + "' "; cnt++;
                }
            }
            if (cnt != 0) strSql += ')';
        }
        else
        {
            strSql += " And Donate_Purpose <> '' ";
        }
        if (HFD_DonateDateS_year1.Value != "")
        {
            strSql += " And Donate_Date>='" + HFD_DonateDateS_year1.Value + "' ";
        }
        if (HFD_DonateDateE_year1.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_year1.Value + "' ";
        }
        if (HFD_Donate_Total_Begin_year1.Value != "" && HFD_Donate_Total_End_year1.Value != "")
        {
            strSql += " GROUP BY Donor_Id HAVING SUM(Donate_Amt) BETWEEN '" + HFD_Donate_Total_Begin_year1.Value + "' AND '" + HFD_Donate_Total_End_year1.Value + "' ) D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id ) D ";
        }
        strSql += @" ON M.Donor_Id = D.Donor_Id 
                     LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
                     LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode
                     Where 1=1 ";
        if (HFD_Is_Abroad_year1.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_year1.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_year1.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        //以下是串UNION的SQL語法
        strSql += @" Union select '99999',CONVERT(VARCHAR,count(D.Donor_Id)),REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(D.Sum_Amt)),1),'.00',''),'','','','','','',''
                           FROM DONOR M   
                           RIGHT JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
							           FROM DONATE 
                                       WHERE ISNULL(Issue_Type,'') != 'D' ";
        //搜尋條件
        if (HFD_Donate_Amt_year1.Value != "")
        {
            strSql += " and Donate_Amt>'" + HFD_Donate_Amt_year1.Value + "' ";
        }
        else
        {
            strSql += " and Donate_Amt > 0 ";
        }
        if (HFD_Donate_Purpose_year1.Value != "")
        {
            char[] ch = new char[] { ',' };
            Donate_Purpose = HFD_Donate_Purpose_year1.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
            count = Donate_Purpose.Length;

            bool first = true; int cnt = 0;
            for (int i = 0; i < count; i++)
            {
                if (first)
                {
                    strSql += " AND (Donate_Purpose='" + Donate_Purpose[i] + "' "; first = false; cnt++;
                }
                else
                {
                    strSql += " or Donate_Purpose='" + Donate_Purpose[i] + "' "; cnt++;
                }
            }
            if (cnt != 0) strSql += ')';
        }
        else
        {
            strSql += " And Donate_Purpose <> '' ";
        }
        if (HFD_DonateDateS_year1.Value != "")
        {
            strSql += " And Donate_Date>='" + HFD_DonateDateS_year1.Value + "' ";
        }
        if (HFD_DonateDateE_year1.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_year1.Value + "' ";
        }
        if (HFD_Donate_Total_Begin_year1.Value != "" && HFD_Donate_Total_End_year1.Value != "")
        {
            strSql += " GROUP BY Donor_Id HAVING SUM(Donate_Amt) BETWEEN '" + HFD_Donate_Total_Begin_year1.Value + "' AND '" + HFD_Donate_Total_End_year1.Value + "' ) D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id ) D ";
        }
        strSql += " ON M.Donor_Id = D.Donor_Id Where 1=1 ";
        if (HFD_Is_Abroad_year1.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_year1.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_year1.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        strSql += "Order by D.Donor_Id ";
        return strSql;
    }
}