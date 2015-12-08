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

public partial class Report_Custom_Report_Donate_Month_Report5_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_DonateDateS_month5.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_month5.Value = Util.GetQueryString("DonateDateE");
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

        DataTable dtRet = CaseUtil.Donate_Month_Report5_Print(dt);

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
                //cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 2)
            {
                //cell.Width = "55mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 3)
            {
                //cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 4)
            {
                //cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 5)
            {
                //cell.Width = "25mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 6)
            {
                //cell.Width = "25mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 7)
            {
                //cell.Width = "100mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 8)
            {
                //cell.Width = "100mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 9)
            {
                //cell.Width = "60mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 10)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 11)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 12)
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
                else if (Col == 12)
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

        //捐款期間
        //DateTime DonateDateS = Convert.ToDateTime(HFD_Year_month5.Value + "/" + HFD_Month_month5.Value + "/" + "1");
        //DateTime DonateDateE = Convert.ToDateTime(HFD_Year_month5.Value + "/" + HFD_Month_month5.Value + "/" + "1").AddMonths(1).AddDays(-1);
        GridList.Text = "捐款方式分析<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>捐款期間：" + HFD_DonateDateS_month5.Value + "~" + HFD_DonateDateE_month5.Value + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款日期排序</span><br>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql = @"Select convert(nvarchar(10),A.Donate_Date,111) AS 'Donate_Date'  
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment01)),1),'.00','') AS '現金'
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment02)),1),'.00','') AS '劃撥'
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment03)),1),'.00','') AS '信用卡授權書(一般)'  
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment04)),1),'.00','') AS '郵局轉帳'
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment05)),1),'.00','') AS '匯款'
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment06)),1),'.00','') AS '支票'  
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment07)),1),'.00','') AS '實物奉獻'
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment08)),1),'.00','') AS 'ATM'
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment09)),1),'.00','') AS '網路信用卡'  
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment10)),1),'.00','') AS 'ACH'
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment11)),1),'.00','') AS '美國運通'
                                 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment01)+Sum(Payment02)+Sum(Payment03)+Sum(Payment04)+Sum(Payment05)+Sum(Payment06)+Sum(Payment07)  
                                 +Sum(Payment08)+Sum(Payment09)+Sum(Payment10)+Sum(Payment11)),1),'.00','') AS '小計'  
                                 FROM  
	                                 (Select Donate_Date  
			                                 , Case When Donate_Payment='現金' Then Donate_Amt Else 0 End 'Payment01'  
			                                 , Case When Donate_Payment='劃撥' Then Donate_Amt Else 0 End 'Payment02'  
			                                 , Case When Donate_Payment='信用卡授權書(一般)' Then Donate_Amt Else 0 End 'Payment03'  
			                                 , Case When Donate_Payment='郵局帳戶轉帳授權書' Then Donate_Amt Else 0 End 'Payment04'  
			                                 , Case When Donate_Payment='匯款' Then Donate_Amt Else 0 End 'Payment05'  
			                                 , Case When Donate_Payment='支票' Then Donate_Amt Else 0 End 'Payment06'  
			                                 , Case When Donate_Payment='實物奉獻' Then Donate_Amt Else 0 End 'Payment07'  
			                                 , Case When Donate_Payment='ATM' Then Donate_Amt Else 0 End 'Payment08'  
			                                 , Case When Donate_Payment='網路信用卡' Then Donate_Amt Else 0 End 'Payment09'  
			                                 , Case When Donate_Payment like '%ACH%' Then Donate_Amt Else 0 End 'Payment10'  
			                                 , Case When Donate_Payment='信用卡授權書(聯信)' Then Donate_Amt Else 0 End 'Payment11' ,Issue_Type 
	                                 From dbo.DONATE ) A";

        //搜尋條件
        if (HFD_DonateDateS_month5.Value != "")
        {
            strSql += " where Donate_Date>='" + HFD_DonateDateS_month5.Value + "' ";
        }
        if (HFD_DonateDateE_month5.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_month5.Value + "' ";
        }
        strSql += "  and Issue_Type != 'D' ";
        strSql += " GROUP BY convert(nvarchar(10),Donate_Date,111) ";
        //以下是串UNION的SQL語法
        strSql += @" Union Select '總計'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment01)),1),'.00','') AS '現金'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment02)),1),'.00','') AS '劃撥'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment03)),1),'.00','') AS '信用卡授權書(一般)'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment04)),1),'.00','')  AS '郵局轉帳'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment05)),1),'.00','') AS '匯款'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment06)),1),'.00','') AS '支票'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment07)),1),'.00','') AS '實物奉獻'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment08)),1),'.00','') AS 'ATM'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment09)),1),'.00','')  AS '網路信用卡'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment10)),1),'.00','') AS 'ACH'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment11)),1),'.00','') AS '美國運通'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Sum(Payment01)+Sum(Payment02)+Sum(Payment03)+Sum(Payment04)+Sum(Payment05)+Sum(Payment06)+Sum(Payment07)  
                            +Sum(Payment08)+Sum(Payment09)+Sum(Payment10)+Sum(Payment11)),1),'.00','') AS '小計'  
                            FROM  
	                            (Select Donate_Date  
			                            , Case When Donate_Payment='現金' Then Donate_Amt Else 0 End 'Payment01'  
			                            , Case When Donate_Payment='劃撥' Then Donate_Amt Else 0 End 'Payment02'  
			                            , Case When Donate_Payment='信用卡授權書(一般)' Then Donate_Amt Else 0 End 'Payment03'  
			                            , Case When Donate_Payment='郵局帳戶轉帳授權書' Then Donate_Amt Else 0 End 'Payment04'  
			                            , Case When Donate_Payment='匯款' Then Donate_Amt Else 0 End 'Payment05'  
			                            , Case When Donate_Payment='支票' Then Donate_Amt Else 0 End 'Payment06'  
			                            , Case When Donate_Payment='實物奉獻' Then Donate_Amt Else 0 End 'Payment07'  
			                            , Case When Donate_Payment='ATM' Then Donate_Amt Else 0 End 'Payment08'  
			                            , Case When Donate_Payment='網路信用卡' Then Donate_Amt Else 0 End 'Payment09'  
			                            , Case When Donate_Payment like '%ACH%' Then Donate_Amt Else 0 End 'Payment10'  
			                            , Case When Donate_Payment='信用卡授權書(聯信)' Then Donate_Amt Else 0 End 'Payment11' ,Issue_Type 
	                            From dbo.DONATE ) A";
                                
        //搜尋條件
        if (HFD_DonateDateS_month5.Value != "")
        {
            strSql += " where Donate_Date>='" + HFD_DonateDateS_month5.Value + "' ";
        }
        if (HFD_DonateDateE_month5.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_month5.Value + "' ";
        }
        strSql += "  and Issue_Type != 'D' ";
        strSql += @"Union SELECT '筆數' AS '筆數',
	                        CONVERT(VARCHAR,ISNULL([現金],0)) As '現金', CONVERT(VARCHAR,ISNULL([劃撥],0)) As '劃撥', 
	                        CONVERT(VARCHAR,ISNULL([信用卡授權書(一般)],0)) As '信用卡授權書(一般)',
	                        CONVERT(VARCHAR,ISNULL([郵局帳戶轉帳授權書],0)) As '郵局轉帳',
	                        CONVERT(VARCHAR,ISNULL([匯款],0)) As '匯款', CONVERT(VARCHAR,ISNULL([支票],0)) As '支票', 
	                        CONVERT(VARCHAR,ISNULL([實物奉獻],0)) As '實物奉獻',
	                        CONVERT(VARCHAR,ISNULL([ATM],0)) As 'ATM', CONVERT(VARCHAR,ISNULL([網路信用卡],0)) As '網路信用卡',
	                        CONVERT(VARCHAR,ISNULL([ACH轉帳授權書],0)) As 'ACH', CONVERT(VARCHAR,ISNULL([信用卡授權書(聯信)],0)) As '美國運通',
	                        CONVERT(VARCHAR,ISNULL([現金],0)+ISNULL([劃撥],0)+ISNULL([信用卡授權書(一般)],0)+ISNULL([郵局帳戶轉帳授權書],0)
	                        +ISNULL([匯款],0)+ISNULL([支票],0)+ISNULL([實物奉獻],0)+ISNULL([ATM],0)+ISNULL([網路信用卡],0)
	                        +ISNULL([ACH轉帳授權書],0)+ISNULL([信用卡授權書(聯信)],0)) AS '總計'
                        FROM (
                            select Donate_Payment,Count(Donate_Id) AS 'Cnt' from donate";
        //搜尋條件
        if (HFD_DonateDateS_month5.Value != "")
        {
            strSql += " where Donate_Date>='" + HFD_DonateDateS_month5.Value + "' ";
        }
        if (HFD_DonateDateE_month5.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_month5.Value + "' ";
        }
        strSql += "  and Issue_Type != 'D' ";
        strSql += @" group by Donate_Payment
                        ) As S
                        PIVOT (
                            Sum(Cnt) FOR
                            Donate_Payment IN ([現金], [劃撥], [信用卡授權書(一般)], [郵局帳戶轉帳授權書], [匯款], [支票]
                            , [實物奉獻], [ATM], [網路信用卡], [ACH轉帳授權書], [信用卡授權書(聯信)])
                        ) As P 
                      order BY convert(nvarchar(10),Donate_Date,111) ";
        
        return strSql;
    }
}