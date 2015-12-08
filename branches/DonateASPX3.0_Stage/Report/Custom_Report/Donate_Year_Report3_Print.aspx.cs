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

public partial class Report_Custom_Report_Donate_Year_Report3_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Donate_Amt_year3.Value = Util.GetQueryString("Donate_Amt");
            HFD_Donate_Total_Amt_year3.Value = Util.GetQueryString("Donate_Total_Amt");
            HFD_DonateDateS_year3.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_year3.Value = Util.GetQueryString("DonateDateE");
            HFD_Is_ErrAddress_year3.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_year3.Value = Util.GetQueryString("Sex");
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

        DataTable dtRet = CaseUtil.Donate_Year_Report3_Print(dt);

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        table.Width = "770";
        css = table.Style;
        css.Add("font-size", "12px");
        css.Add("font-family", "細明體");
        css.Add("line-height", "20px");
        row = new HtmlTableRow();
        /////格式很怪，若需要顯示的話需要再修改
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
                //cell.Width = "40mm";
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
                else if (Col == 9)
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

        GridList.Text = "國外捐款總額明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計期間：" + HFD_DonateDateS_year3.Value + "~" + HFD_DonateDateE_year3.Value + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人編號排序</span><br>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql = @"  SELECT D.Donor_Id AS '編號'   
                                    ,M.Donor_Name AS '天使姓名' ,M.Title AS '稱謂'
                                    ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Max_Amt),1),'.00','') AS '最高奉獻金額'  
								    ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Sum_Amt),1),'.00','') AS '累計奉獻金額' 
								    ,D.Times AS '奉獻次數'  
                                    ,(SELECT DISTINCT Donate_Payment + '.' FROM dbo.DONATE as D1 WHERE M.Donor_Id=D1.Donor_Id ";
        if (HFD_DonateDateS_year3.Value != "")
        {
            strSql += " And Donate_Date>='" + HFD_DonateDateS_year3.Value + "' ";
        }
        if (HFD_DonateDateE_year3.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_year3.Value + "' ";
        }
        strSql += @"FOR XML PATH('') ) AS '捐款方式' 
								    ,IsNull(OverseasAddress,'') AS '地址' ,ISNULL(Email,'') AS 'Email',IsNull(OverseasCountry,'') AS '國別'  
							FROM DONOR M  
							INNER JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt, MAX(Donate_Amt) AS Max_Amt ,COUNT(Donor_Id) AS Times  
									    FROM DONATE 
                                        WHERE ISNULL(Issue_Type,'') != 'D' ";

        //搜尋條件
        if (HFD_Donate_Amt_year3.Value != "")
        {
            strSql += " and Donate_Amt>'" + HFD_Donate_Amt_year3.Value + "' ";
        }
        else
        {
            strSql += " and Donate_Amt > 0 ";
        }
        if (HFD_DonateDateS_year3.Value != "")
        {
            strSql += " And Donate_Date>='" + HFD_DonateDateS_year3.Value + "' ";
        }
        if (HFD_DonateDateE_year3.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_year3.Value + "' ";
        }
        if (HFD_Donate_Total_Amt_year3.Value != "")
        {
            strSql += " GROUP BY Donor_Id Having SUM(Donate_Amt) > " + HFD_Donate_Total_Amt_year3.Value + " ) D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id ) D ";
        }
        strSql += " ON M.Donor_Id = D.Donor_Id ";
        strSql += " WHERE IsAbroad = 'Y' ";
        if (HFD_Is_ErrAddress_year3.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_year3.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        //以下是串UNION的SQL語法
        strSql += @" Union  select '99999',CONVERT(VARCHAR,count(D.Donor_Id)),REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(D.Sum_Amt)),1),'.00',''),'','','','','','',''
                            FROM DONOR M  
                            INNER JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt, MAX(Donate_Amt) AS Max_Amt ,COUNT(Donor_Id) AS Times  
                                        FROM DONATE 
                                        WHERE ISNULL(Issue_Type,'') != 'D' ";
        //搜尋條件
        if (HFD_Donate_Amt_year3.Value != "")
        {
            strSql += " and Donate_Amt>'" + HFD_Donate_Amt_year3.Value + "' ";
        }
        else
        {
            strSql += " and Donate_Amt > 0 ";
        }
        if (HFD_DonateDateS_year3.Value != "")
        {
            strSql += " And Donate_Date>='" + HFD_DonateDateS_year3.Value + "' ";
        }
        if (HFD_DonateDateE_year3.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_year3.Value + "' ";
        }
        if (HFD_Donate_Total_Amt_year3.Value != "")
        {
            strSql += " GROUP BY Donor_Id Having SUM(Donate_Amt) > " + HFD_Donate_Total_Amt_year3.Value + " ) D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id ) D ";
        }
        strSql += " ON M.Donor_Id = D.Donor_Id ";
        strSql += " WHERE IsAbroad = 'Y' ";
        if (HFD_Is_ErrAddress_year3.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_year3.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        strSql += "Order by D.Donor_Id ";
        return strSql;
    }
}