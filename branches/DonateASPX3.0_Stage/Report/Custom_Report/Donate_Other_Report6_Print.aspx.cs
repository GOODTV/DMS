using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class Report_Custom_Report_Donate_Other_Report6_Print : BasePage
{
    string TableWidth = "800px";

    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Type_other6.Value = Util.GetQueryString("Type");
            HFD_DonateDateS_other6.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_other6.Value = Util.GetQueryString("DonateDateE");
            HFD_Is_Abroad_other6.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_other6.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_other6.Value = Util.GetQueryString("Sex");
            LoadFormData();
        }
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql = Sql();

        string Report_condition = HFD_Type_other6.Value;

        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;

        DataTable dtRet = CaseUtil.Donate_Other_Report6_Print(dt, Report_condition);

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "16px");
        css.Add("font-family", "標楷體");
        css.Add("line-height", "20px");

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
            if (iCtrl == 0)
            {
                cell.Width = "250mm";
            }
            else
            {
                cell.Width = "250mm";
            }
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);


        //put data
        foreach (DataRow dr in dtRet.Rows)
        {

            row = new HtmlTableRow();
            foreach (object objItem in dr.ItemArray)
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

        string Condition_Chs = "";
        if (Report_condition == "1")
        {
            Condition_Chs = "性別";
        }
        else if (Report_condition == "2")
        {
            Condition_Chs = "年齡";
        }
        else if (Report_condition == "3")
        {
            Condition_Chs = "通訊縣市";
        }
        else if (Report_condition == "4")
        {
            Condition_Chs = "身份別";
        }
        else if (Report_condition == "5")
        {
            Condition_Chs = "通訊資料";
        }
        else if (Report_condition == "6")
        {
            Condition_Chs = "訂閱月刊";
        }
        else if (Report_condition == "7")
        {
            Condition_Chs = "奉獻金額";
        }
        else if (Report_condition == "8")
        {
            Condition_Chs = "捐款次數";
        }
        GridList.Text = GetTitle(Condition_Chs) + GetCondition(Condition_Chs) + htw.InnerWriter.ToString();
    }
    private string Sql()
    {
        string strSql;
        switch (HFD_Type_other6.Value)
        {
            case "1":  //性別
                strSql = @"select 
                        sum(case when ISNULL(Sex,'')='男' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')='男' then 1 else 0 end) AS FLOAT)/(count (ISNULL(Sex,''))) * 100,2) as [百分比]
                        ,sum(case when ISNULL(Sex,'')='女' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')='女' then 1 else 0 end) AS FLOAT)/(count (ISNULL(Sex,''))) * 100,2) as [百分比]
                        ,sum(case when ISNULL(Sex,'')='歿' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')='歿' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]
                        ,sum(case when ISNULL(Sex,'')<>'男' and ISNULL(Sex,'')<>'女' and ISNULL(Sex,'')<>'歿'  then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')<>'男' and ISNULL(Sex,'')<>'女' and ISNULL(Sex,'')<>'歿' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]
                        ,count (M.Donor_Id) as [人數]
                        ,Round(CAST(count (M.Donor_Id) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]
                        from Donor M
                        Inner JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
			                                                FROM DONATE 
                                                            WHERE ISNULL(Issue_Type,'') != 'D' ";
                break;
            case "2":  //年齡
                strSql = @"select 
                            sum(case when DATEDIFF(year ,Birthday,getdate())<21 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate())<21 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 21 and 25 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 21 and 25 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 26 and 30 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 26 and 30 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 31 and 35 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 31 and 35 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 36 and 40 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 36 and 40 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 41 and 45 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 41 and 45 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 46 and 50 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 46 and 50 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 51 and 55 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 51 and 55 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 56 and 60 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 56 and 60 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 61 and 65 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 61 and 65 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 66 and 70 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 66 and 70 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when DATEDIFF(year ,Birthday,getdate()) >70 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) >70 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Birthday is null then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Birthday is null then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,count (M.Donor_Id) as [人數]
                            ,Round(CAST(count (M.Donor_Id) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]
                            from Donor M
                            Inner JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
			                                                    FROM DONATE 
                                                                WHERE ISNULL(Issue_Type,'') != 'D'";
                break;
            case "3":  //通訊縣市
                strSql = @"select 
                            sum(case when A.mValue='基隆市' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='基隆市' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='台北市' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='台北市' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='新北市' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='新北市' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='桃園市' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='桃園市' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='新竹市' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='新竹市' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='新竹縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='新竹縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='苗栗縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='苗栗縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='台中市' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='台中市' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='彰化縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='彰化縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='南投縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='南投縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='雲林縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='雲林縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='嘉義市' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='嘉義市' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='嘉義縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='嘉義縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='台南市' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='台南市' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='高雄市' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='高雄市' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='屏東縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='屏東縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='宜蘭縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='宜蘭縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='花蓮縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='花蓮縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='台東縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='台東縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='澎湖縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='澎湖縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='金門縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='金門縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when A.mValue='連江縣' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when A.mValue='連江縣' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when City is null or City = '' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when City is null or City = '' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,count (M.Donor_Id) as [人數]
                            ,Round(CAST(count (M.Donor_Id) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]
                            from Donor M
                            Left Join CODECITY As A On M.City=A.mCode 
                            Inner JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
			                                                    FROM DONATE 
                                                                WHERE ISNULL(Issue_Type,'') != 'D'";
                break;
            case "4":  //身份別
                strSql = @"select 
                            sum(case when Donor_Type like '%個人%' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when Donor_Type like '%個人%' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比] 

                            ,sum(case when Donor_Type like '%公司行號%' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when Donor_Type like '%公司行號%' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,sum(case when Donor_Type like '%教會%' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when Donor_Type like '%教會%' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,sum(case when Donor_Type like '%機構%' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when Donor_Type like '%機構%' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,sum(case when Donor_Type like '%醫院%' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when Donor_Type like '%醫院%' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,sum(case when Donor_Type like '%學校%' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when Donor_Type like '%學校%' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,Count(*) as [人數],Round(CAST(count (M.Donor_Id) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]
                            from Donor M
                            Inner JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
			                            FROM DONATE 
                                        WHERE ISNULL(Issue_Type,'') != 'D'";
                break;
            case "5":  //通訊資料
                strSql = @"select 
                            sum(case when Cellular_Phone != '' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when Cellular_Phone != '' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,sum(case when (Tel_Home != '' or Tel_Office != '') then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when (Tel_Home != '' or Tel_Office != '') then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,sum(case when Email != '' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when Email != '' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,Count(*) as [人數],Round(CAST(count (M.Donor_Id) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            from Donor M
                            Inner JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
			                                                    FROM DONATE 
                                                                WHERE ISNULL(Issue_Type,'') != 'D'";
                break;
            case "6":  //訂閱月刊
                strSql = @"select 
                            sum(case when IsSendNews='Y' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when IsSendNews='Y' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,sum(case when IsSendNews='N' then 1 else 0 end) as [人數]
                            ,Round(CAST(sum(case when IsSendNews='N' then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,Count(*) as [人數],Round(CAST(count (M.Donor_Id) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            from Donor M
                            Inner JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
			                                                    FROM DONATE 
                                                                WHERE ISNULL(Issue_Type,'') != 'D'";
                break;
            case "7":  //累計奉獻金額
                strSql = @"select 
                            sum(case when Sum_Amt >= 100000000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 100000000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 10000000 and Sum_Amt <100000000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 10000000 and Sum_Amt <100000000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 8000000 and Sum_Amt < 10000000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 8000000 and Sum_Amt < 10000000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 7000000 and Sum_Amt < 8000000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 7000000 and Sum_Amt < 8000000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 6000000 and Sum_Amt < 7000000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 6000000 and Sum_Amt < 7000000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 5000000 and Sum_Amt < 6000000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 5000000 and Sum_Amt < 6000000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 4000000 and Sum_Amt < 5000000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 4000000 and Sum_Amt < 5000000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 3000000 and Sum_Amt < 4000000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 3000000 and Sum_Amt < 4000000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 2000000 and Sum_Amt < 3000000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 2000000 and Sum_Amt < 3000000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 1000000 and Sum_Amt < 2000000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 1000000 and Sum_Amt < 2000000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 900000 and Sum_Amt < 1000000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 900000 and Sum_Amt < 1000000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 800000 and Sum_Amt < 900000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 800000 and Sum_Amt < 900000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 700000 and Sum_Amt < 800000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 700000 and Sum_Amt < 800000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 600000 and Sum_Amt < 700000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 600000 and Sum_Amt < 700000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 500000 and Sum_Amt < 600000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 500000 and Sum_Amt < 600000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 400000 and Sum_Amt < 500000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 400000 and Sum_Amt < 500000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 300000 and Sum_Amt < 400000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 300000 and Sum_Amt < 400000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 200000 and Sum_Amt < 300000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 200000 and Sum_Amt < 300000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 100000 and Sum_Amt < 200000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 100000 and Sum_Amt < 200000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 90000 and Sum_Amt < 100000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 90000 and Sum_Amt < 100000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 80000 and Sum_Amt < 90000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 80000 and Sum_Amt < 90000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 70000 and Sum_Amt < 80000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 70000 and Sum_Amt < 80000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 60000 and Sum_Amt < 70000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 60000 and Sum_Amt < 70000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 50000 and Sum_Amt < 60000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 50000 and Sum_Amt < 60000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 40000 and Sum_Amt < 50000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 40000 and Sum_Amt < 50000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 30000 and Sum_Amt < 40000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 30000 and Sum_Amt < 40000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 20000 and Sum_Amt < 30000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt >= 20000 and Sum_Amt < 30000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt >= 10000 and Sum_Amt < 20000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt between 10000 and 20000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Sum_Amt<10000 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Sum_Amt<10000 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,count (M.Donor_Id) as [人數]
                            ,Round(CAST(count (M.Donor_Id) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]
                            from Donor M
                            Inner JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
			                                                    FROM DONATE 
                                                                WHERE ISNULL(Issue_Type,'') != 'D'";
                break;
            case "8":  //捐款次數
                strSql = @"select 
                            sum(case when Times >= 300 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times >= 300 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times between 200 and 299 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 200 and 299 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times between 150 and 199 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 150 and 199 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times between 100 and 149 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 100 and 149 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times between 90 and 99 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 90 and 99 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times between 80 and 89 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 80 and 89 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times between 70 and 79 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 70 and 79 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times between 60 and 69 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 60 and 69 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times between 50 and 59 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 50 and 59 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times between 40 and 49 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 40 and 49 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times between 30 and 39 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 30 and 39 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when  Times between 20 and 29 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 20 and 29 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when  Times between 15 and 19 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times between 15 and 19 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times between 10 and 14 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when  Times between 10 and 14 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times=9 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times=9 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100,2) as [百分比]

                            ,sum(case when Times=8 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times=8 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times=7 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times=7 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times=6 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times=6 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times=5 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times=5 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times=4 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times=4 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times=3 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times=3 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times=2 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times=2 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,sum(case when Times=1 then 1 else 0 end)  as [人數]
                            ,Round(CAST(sum(case when Times=1 then 1 else 0 end) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]

                            ,count (M.Donor_Id) as [人數]
                            ,Round(CAST(count (M.Donor_Id) AS FLOAT)/(count (M.Donor_Id)) * 100 ,2) as [百分比]
                            from Donor M
                            Inner JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
			                                                    FROM DONATE 
                                                                WHERE ISNULL(Issue_Type,'') != 'D'";
                break;
            default:
                strSql = "";
                break;

        }
        
        //搜尋條件
        if (HFD_DonateDateS_other6.Value != "")
        {
            strSql += " And Donate_Date>='" + HFD_DonateDateS_other6.Value + "'";
        }
        if (HFD_DonateDateE_other6.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_other6.Value + "'";
        }
        strSql += @"GROUP BY Donor_Id) D
			ON M.Donor_Id = D.Donor_Id 
            Where DeleteDate is NULL and Donor_Type not like '%讀者%' ";
        if (HFD_Is_Abroad_other6.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_other6.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_other6.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        
        return strSql;
    }
    private string GetTitle(string Condition_Chs)
    {
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
        css.Add("width", "100%");

        string strFontSize = "18px";
        string strTitle = "捐款人" + Condition_Chs + "分析<br/>";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        row.Cells.Add(cell);
        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();

    }
    private string GetCondition(string Condition_Chs)
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
        css.Add("width", "80%");
        css.Add("line-height", "20px");

        string strTitle = "統計項目：" + Condition_Chs + "<br>";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "left");
        row.Cells.Add(cell);
        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
}