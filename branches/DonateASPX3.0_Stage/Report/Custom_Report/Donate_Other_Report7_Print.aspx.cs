﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class Report_Custom_Report_Donate_Other_Report7_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_DonateDateS_other7.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_other7.Value = Util.GetQueryString("DonateDateE");
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

        DataTable dtRet = CaseUtil.Donate_Other_Report7_Print(dt);

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
            else if (iCtrl == 15)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-right", ".5pt solid windowtext");
            }
            else
            {
                //cell.Width = "40mm";
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
                else if (Col == 15)
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
        GridList.Text = "奉獻動機及收視管道統計分析<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計期間：" + HFD_DonateDateS_other7.Value + "~" + HFD_DonateDateE_other7.Value + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款日期排序</span><br>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql = @"Select convert(nvarchar(10),A.Create_Date,111) AS 'Create_Date'  
                                ,Sum(Survey01) AS '支持媒體宣教大平台，可廣傳福音'
                                ,Sum(Survey02) AS '個人靈命得造就'
                                ,Sum(Survey03) AS '支持優質節目製作'
                                ,Sum(Survey04) AS '支持GOOD TV家庭事工'
                                ,Sum(Survey05) AS '感恩奉獻'
                                ,Sum(Survey15) AS '其他'
                                ,Sum(Survey06) AS 'GOOD TV電視頻道'
                                ,Sum(Survey07) AS '官網'
                                ,Sum(Survey08) AS 'Facebook'
                                ,Sum(Survey09) AS 'Youtube'
                                ,Sum(Survey10) AS '好消息月刊'
                                ,Sum(Survey11) AS 'GOOD TV簡介刊物'
                                ,Sum(Survey12) AS '教會牧者'
                                ,Sum(Survey13) AS '親友'
                                ,Sum(Survey14) AS '報章雜誌'
                                FROM  
                                (Select Create_Date  
		                                , Case When DonateMotive1='' Then 0 Else 1 End 'Survey01'  
		                                , Case When DonateMotive2='' Then 0 Else 1 End 'Survey02'  
		                                , Case When DonateMotive3='' Then 0 Else 1 End 'Survey03'  
		                                , Case When DonateMotive4='' Then 0 Else 1 End 'Survey04'  
		                                , Case When DonateMotive5='' Then 0 Else 1 End 'Survey05' 
                                        , Case When ToGOODTV is NULL Then 0 Else 1 End 'Survey15'   
		                                , Case When WatchMode1='' Then 0 Else 1 End 'Survey06'  
		                                , Case When WatchMode2='' Then 0 Else 1 End 'Survey07'  
		                                , Case When WatchMode3='' Then 0 Else 1 End 'Survey08'  
		                                , Case When WatchMode4='' Then 0 Else 1 End 'Survey09'  
		                                , Case When WatchMode5='' Then 0 Else 1 End 'Survey10'  
		                                , Case When WatchMode6='' Then 0 Else 1 End 'Survey11'
		                                , Case When WatchMode7='' Then 0 Else 1 End 'Survey12'
		                                , Case When WatchMode8='' Then 0 Else 1 End 'Survey13'
		                                , Case When WatchMode9='' Then 0 Else 1 End 'Survey14'
                                From dbo.Donate_OnlineQuestion ) A";

        //搜尋條件
        if (HFD_DonateDateS_other7.Value != "")
        {
            strSql += " where Create_Date>='" + HFD_DonateDateS_other7.Value + "' ";
        }
        if (HFD_DonateDateE_other7.Value != "")
        {
            strSql += " And Create_Date<='" + HFD_DonateDateE_other7.Value + "' ";
        }
        strSql += " GROUP BY convert(nvarchar(10),Create_Date,111) ";
        //以下是串UNION的SQL語法
        strSql += @" Union Select '總計'
                                    ,Sum(Survey01) AS '支持媒體宣教大平台，可廣傳福音'
                                    ,Sum(Survey02) AS '個人靈命得造就'
                                    ,Sum(Survey03) AS '支持優質節目製作'
                                    ,Sum(Survey04) AS '支持GOOD TV家庭事工'
                                    ,Sum(Survey05) AS '感恩奉獻'
                                    ,Sum(Survey15) AS '其他'
                                    ,Sum(Survey06) AS 'GOOD TV電視頻道'
                                    ,Sum(Survey07) AS '官網'
                                    ,Sum(Survey08) AS 'Facebook'
                                    ,Sum(Survey09) AS 'Youtube'
                                    ,Sum(Survey10) AS '好消息月刊'
                                    ,Sum(Survey11) AS 'GOOD TV簡介刊物'
                                    ,Sum(Survey12) AS '教會牧者'
                                    ,Sum(Survey13) AS '親友'
                                    ,Sum(Survey14) AS '報章雜誌'
                                    FROM  
                                    (Select Create_Date
		                                    , Case When DonateMotive1='' Then 0 Else 1 End 'Survey01'  
		                                    , Case When DonateMotive2='' Then 0 Else 1 End 'Survey02'  
		                                    , Case When DonateMotive3='' Then 0 Else 1 End 'Survey03'  
		                                    , Case When DonateMotive4='' Then 0 Else 1 End 'Survey04'  
		                                    , Case When DonateMotive5='' Then 0 Else 1 End 'Survey05'  
                                            , Case When ToGOODTV is NULL Then 0 Else 1 End 'Survey15'  
		                                    , Case When WatchMode1='' Then 0 Else 1 End 'Survey06'  
		                                    , Case When WatchMode2='' Then 0 Else 1 End 'Survey07'  
		                                    , Case When WatchMode3='' Then 0 Else 1 End 'Survey08'  
		                                    , Case When WatchMode4='' Then 0 Else 1 End 'Survey09'  
		                                    , Case When WatchMode5='' Then 0 Else 1 End 'Survey10'  
		                                    , Case When WatchMode6='' Then 0 Else 1 End 'Survey11'
		                                    , Case When WatchMode7='' Then 0 Else 1 End 'Survey12'
		                                    , Case When WatchMode8='' Then 0 Else 1 End 'Survey13'
		                                    , Case When WatchMode9='' Then 0 Else 1 End 'Survey14'
                                    From dbo.Donate_OnlineQuestion ) A";

        //搜尋條件
        if (HFD_DonateDateS_other7.Value != "")
        {
            strSql += " where Create_Date>='" + HFD_DonateDateS_other7.Value + "' ";
        }
        if (HFD_DonateDateE_other7.Value != "")
        {
            strSql += " And Create_Date<='" + HFD_DonateDateE_other7.Value + "' ";
        }
        return strSql;
    }
}