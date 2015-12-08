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

public partial class Report_Custom_Report_Donate_Other_Report5_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_IsMember_other5.Value = Util.GetQueryString("IsMember");
            PrintReport();
        }
    }
    private void PrintReport()
    {
        string strSql = Sql();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
            return;
        }

        DataTable dtRet = CaseUtil.Donate_Other_Report5_Print(dt);

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
                //cell.Width = "40mm";
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
                else if (Col == 2)
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

        GridList.Text = "讀者分析<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + DateTime.Now.ToString() + "</span><br>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql;
        //if (HFD_IsMember_other5.Value == "Y")  //讀者
        //{
            strSql = @"select '有手機' as 項目,sum(case when Cellular_Phone != '' then 1 else 0 end) as 統計人數
                        ,Round(CAST(sum(case when Cellular_Phone != '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
                         from Donor where DeleteDate is NULL and Donor_Type like '%讀者%'
                        Union
                        select '有室內電話' as 項目,sum(case when (Tel_Home != '' or Tel_Office != '') then 1 else 0 end) as 統計人數
                        ,Round(CAST(sum(case when (Tel_Home != '' or Tel_Office != '') then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
                         from Donor where DeleteDate is NULL and Donor_Type like '%讀者%'
                        Union
                        select '有Email' as 項目,sum(case when Email != '' then 1 else 0 end) as 統計人數
                        ,Round(CAST(sum(case when Email != '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
                         from Donor where DeleteDate is NULL and Donor_Type like '%讀者%'
                        Union
                        select '總人數' as 項目,Count(*) as 統計人數,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
                        from Donor where DeleteDate is NULL and Donor_Type like '%讀者%'";
        //}
//        else  //捐款人
//        {
//            strSql = @"select '個人' as 項目,sum(case when Donor_Type like '%個人%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%個人%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比] 
//                        from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '董監事' as 項目,sum(case when Donor_Type like '%董監事%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%董監事%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比] 
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '顧問' as 項目,sum(case when Donor_Type like '%顧問%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%顧問%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比] 
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select 'VIP' as 項目,sum(case when Donor_Type like '%VIP%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%VIP%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '公司行號' as 項目,sum(case when Donor_Type like '%公司行號%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%公司行號%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '教會' as 項目,sum(case when Donor_Type like '%教會%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%教會%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '機構' as 項目,sum(case when Donor_Type like '%機構%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%機構%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '醫院' as 項目,sum(case when Donor_Type like '%醫院%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%醫院%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '學校' as 項目,sum(case when Donor_Type like '%學校%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%學校%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '團契' as 項目,sum(case when Donor_Type like '%團契%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%團契%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '專案' as 項目,sum(case when Donor_Type like '%專案%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%專案%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '主知名' as 項目,sum(case when Donor_Type like '%主知名%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%主知名%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '中國' as 項目,sum(case when Donor_Type like '%中國%' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Donor_Type like '%中國%' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '有手機' as 項目,sum(case when Cellular_Phone != '' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Cellular_Phone != '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '有室內電話' as 項目,sum(case when (Tel_Home != '' or Tel_Office != '') then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when (Tel_Home != '' or Tel_Office != '') then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '有Email' as 項目,sum(case when Email != '' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when Email != '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'
//                        Union
//                        select '有訂月刊' as 項目,sum(case when IsSendNews='Y' then 1 else 0 end) as 統計人數
//                        ,Round(CAST(sum(case when IsSendNews='Y' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y' 
//                        Union
//                        select '總人數' as 項目,Count(*) as 統計人數,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
//                         from Donor where DeleteDate is NULL and ISNULL(IsMember,'') <> 'Y'";

//        }

        return strSql;
    }
}