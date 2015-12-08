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

public partial class DonateMgr_VerifyList_Print_Excel_style3 : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";
    protected void Page_Load(object sender, EventArgs e)
    {
        Print_Excel();
    }
    //---------------------------------------------------------------------------
    private void Print_Excel()
    {
        string strSql = "";
        bool first = true;

        string strData = "";
        DataTable dp;

        //20140417 Modify by GoodTV Tanya
        string strDonate_Purpose = "select Donate_Purpose from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id where dr.DeleteDate is null " + Session["Condition"] + "group by Donate_Purpose ";
        DataTable dp_item = NpoDB.QueryGetTable(strDonate_Purpose);
        for (int j = 0; j < dp_item.Rows.Count; j++)
        {
            strData = "select * from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id where dr.DeleteDate is null and Donate_Purpose = '" + dp_item.Rows[j]["Donate_Purpose"].ToString() + "'" + Session["Condition"];

            dp = NpoDB.QueryGetTable(strData);
            if (dp.Rows.Count != 0)
            {
                if (first == false)
                    strSql += " union ";

                strSql += @"SELECT '" + j.ToString() + @"' as [分類], ROW_NUMBER() OVER(ORDER BY Donate_Date) AS [序號],'" + dp_item.Rows[j]["Donate_Purpose"].ToString() + @"', CONVERT(NVARCHAR, Donate_Date, 111)  as [捐款日期],
                               Donor_Name as [捐款人], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','') as [捐款金額]
                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
                        where dr.DeleteDate is null and Donate_Purpose = '" + dp_item.Rows[j]["Donate_Purpose"].ToString() + @"' " + Session["Condition"] + @"
                        union
                        select '" + j.ToString() + @"1' as [分類],null,'','','','金額小計：'+REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when Donate_Purpose = '" + dp_item.Rows[j]["Donate_Purpose"].ToString() + @"' then Donate_Amt else '0' end)),1),'.00','')
                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
                        where dr.DeleteDate is null " + Session["Condition"];
                first = false;
            }
        }  
//        strData = "select * from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id where dr.DeleteDate is null and Donate_Purpose = '一般' " + Session["Condition"];
//        dp = NpoDB.QueryGetTable(strData);
//        if (dp.Rows.Count != 0 && first == true)
//        {
//            strSql += @"SELECT '1' as [分類], ROW_NUMBER() OVER(ORDER BY Donate_Date) AS [序號],'一般', CONVERT(NVARCHAR, Donate_Date, 111)  as [捐款日期],
//                               Donor_Name as [捐款人], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','') as [捐款金額]
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null and (Donate_Purpose = '一般' or Donate_Purpose = '一般捐款') " + Session["Condition"] + @"
//                        union
//                        select '11' as [分類],null,'','','','金額小計：'+REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when (Donate_Purpose = '一般' or Donate_Purpose = '一般捐款') then Donate_Amt else '0' end)),1),'.00','')
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null " + Session["Condition"];
//            first = false;
//        }

//        strData = "select * from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id where dr.DeleteDate is null and Donate_Purpose = '建台'" + Session["Condition"];
//        dp = NpoDB.QueryGetTable(strData);
//        if (dp.Rows.Count != 0)
//        {
//            if (first == false)
//            {
//                strSql += " union ";
//            }
//            else
//            {
//                first = false;
//            }
//            strSql += @"SELECT '2' as [分類], ROW_NUMBER() OVER(ORDER BY Donate_Date) AS [序號],'建台', CONVERT(NVARCHAR, Donate_Date, 111)  as [捐款日期],
//                               Donor_Name as [捐款人], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','') as [捐款金額]
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null and Donate_Purpose = '建台' " + Session["Condition"] + @"
//                        union
//                        select '21' as [分類],null,'','','','金額小計：'+REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when Donate_Purpose = '建台' then Donate_Amt else '0' end)),1),'.00','')
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null " + Session["Condition"];
//        }

//        strData = "select * from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id where dr.DeleteDate is null and Donate_Purpose = '定點教室'" + Session["Condition"];
//        dp = NpoDB.QueryGetTable(strData);
//        if (dp.Rows.Count != 0)
//        {
//            if (first == false)
//            {
//                strSql += " union ";
//            }
//            else
//            {
//                first = false;
//            }
//            strSql += @"SELECT '3' as [分類], ROW_NUMBER() OVER(ORDER BY Donate_Date) AS [序號],'定點教室', CONVERT(NVARCHAR, Donate_Date, 111)  as [捐款日期],
//                               Donor_Name as [捐款人], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','') as [捐款金額]
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null and Donate_Purpose = '定點教室' " + Session["Condition"] + @"
//                        union
//                        select '31' as [分類],null,'','','','金額小計：'+REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when Donate_Purpose = '定點教室' then Donate_Amt else '0' end)),1),'.00','')
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null " + Session["Condition"];
//        }

//        strData = "select * from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id where dr.DeleteDate is null and Donate_Purpose = '專款'" + Session["Condition"];
//        dp = NpoDB.QueryGetTable(strData);
//        if (dp.Rows.Count != 0)
//        {
//            if (first == false)
//            {
//                strSql += " union ";
//            }
//            else
//            {
//                first = false;
//            }
//            strSql += @"SELECT '4' as [分類], ROW_NUMBER() OVER(ORDER BY Donate_Date) AS [序號],'專款', CONVERT(NVARCHAR, Donate_Date, 111)  as [捐款日期],
//                               Donor_Name as [捐款人], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','') as [捐款金額]
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null and Donate_Purpose = '專款' " + Session["Condition"] + @"
//                        union
//                        select '41' as [分類],null,'','','','金額小計：'+REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when Donate_Purpose = '專款' then Donate_Amt else '0' end)),1),'.00','')
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null " + Session["Condition"];
//        }

//        strData = "select * from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id where dr.DeleteDate is null and Donate_Purpose = '以色列事工'" + Session["Condition"];
//        dp = NpoDB.QueryGetTable(strData);
//        if (dp.Rows.Count != 0)
//        {
//            if (first == false)
//            {
//                strSql += " union ";
//            }
//            else
//            {
//                first = false;
//            }
//            strSql += @"SELECT '5' as [分類], ROW_NUMBER() OVER(ORDER BY Donate_Date) AS [序號],'以色列事工', CONVERT(NVARCHAR, Donate_Date, 111)  as [捐款日期],
//                               Donor_Name as [捐款人], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','') as [捐款金額]
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null and Donate_Purpose = '以色列事工' " + Session["Condition"] + @"
//                        union
//                        select '51' as [分類],null,'','','','金額小計：'+REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when Donate_Purpose = '以色列事工' then Donate_Amt else '0' end)),1),'.00','')
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null " + Session["Condition"];
//        }

//        strData = "select * from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id where dr.DeleteDate is null and Donate_Purpose = '不永康事工'" + Session["Condition"];
//        dp = NpoDB.QueryGetTable(strData);
//        if (dp.Rows.Count != 0)
//        {
//            if (first == false)
//            {
//                strSql += " union ";
//            }
//            else
//            {
//                first = false;
//            }
//            strSql += @"SELECT '6' as [分類], ROW_NUMBER() OVER(ORDER BY Donate_Date) AS [序號],'不永康事工', CONVERT(NVARCHAR, Donate_Date, 111)  as [捐款日期],
//                               Donor_Name as [捐款人], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','') as [捐款金額]
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null and Donate_Purpose = '不永康事工' " + Session["Condition"] + @"
//                        union
//                        select '61' as [分類],null,'','','','金額小計：'+REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when Donate_Purpose = '不永康事工' then Donate_Amt else '0' end)),1),'.00','')
//                        from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id
//                        where dr.DeleteDate is null " + Session["Condition"];
//        }

        DataTable dt = NpoDB.QueryGetTable(strSql);
        if (dp_item.Rows.Count == 0)
        {
            ShowSysMsg("查無資料!!!");
            return;
        }
        GridList.Text = GetTitle() + GetTime() + GetTable1(dt.Rows[0], strSql) + "<br/>" + GetEnd();
        //GetTableFooter();
        Util.OutputTxt(GridList.Text, "1", "Verify");

    }
    //---------------------------------------------------------------------------
    private string GetTitle()
    {
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "13pt");
        css.Add("font-family", "標楷體");
        css.Add("width", TableWidth);

        string strTitle = "徵信芳名錄<br/> ";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        cell.ColSpan = 5;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    //---------------------------------------------------------------------------
    private string GetTime()
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
        css.Add("width", TableWidth);

        string strTitle = "匯出日期：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "<br/> ";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "left");
        cell.ColSpan = 5;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    //---------------------------------------------------------------------------
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

        DataTable dtRet = CaseUtil.VerifyList_Print_style3(dt);

        table.Border = 1;
        table.CellPadding = 3;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "9pt");
        css.Add("font-family", "新細明體");
        css.Add("width", TableWidth);
        css.Add("line-height", "40px");

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

            //cell.Width = "100mm";
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

        foreach (DataRow drow in dtRet.Rows)
        {

            row = new HtmlTableRow();
            foreach (object objItem in drow.ItemArray)
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

        return htw.InnerWriter.ToString();
    }
    //---------------------------------------------------------------------------
    private string GetEnd()
    {
        string strSql = @"SELECT REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(do.Donate_Amt)),1),'.00','') as [Donate_Amt]
                          FROM Donate do left join Donor dr on do.Donor_Id = dr.Donor_id where dr.DeleteDate is null ";
        strSql += Session["Condition"];
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];

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
        css.Add("width", TableWidth);

        string strTitle = "捐款總金額：" + dr["Donate_Amt"];
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "left");
        cell.ColSpan = 5;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
}