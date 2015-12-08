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

public partial class Report_Custom_Report_Donate_Other_Report8_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_DonorName_other8.Value = Util.GetQueryString("DonorName");
            HFD_DonateMotive1_other8.Value = Util.GetQueryString("DonateMotive1");
            HFD_DonateMotive2_other8.Value = Util.GetQueryString("DonateMotive2");
            HFD_DonateMotive3_other8.Value = Util.GetQueryString("DonateMotive3");
            HFD_DonateMotive4_other8.Value = Util.GetQueryString("DonateMotive4");
            HFD_DonateMotive5_other8.Value = Util.GetQueryString("DonateMotive5");
            HFD_DonateMotive6_other8.Value = Util.GetQueryString("DonateMotive6");
            HFD_WatchMode1_other8.Value = Util.GetQueryString("WatchMode1");
            HFD_WatchMode2_other8.Value = Util.GetQueryString("WatchMode2");
            HFD_WatchMode3_other8.Value = Util.GetQueryString("WatchMode3");
            HFD_WatchMode4_other8.Value = Util.GetQueryString("WatchMode4");
            HFD_WatchMode5_other8.Value = Util.GetQueryString("WatchMode5");
            HFD_WatchMode6_other8.Value = Util.GetQueryString("WatchMode6");
            HFD_WatchMode7_other8.Value = Util.GetQueryString("WatchMode7");
            HFD_WatchMode8_other8.Value = Util.GetQueryString("WatchMode8");
            HFD_WatchMode9_other8.Value = Util.GetQueryString("WatchMode9");
            Print_Excel();
        }
    }
    //---------------------------------------------------------------------------
    private void Print_Excel()
    {
        string strSql = Sql();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            ShowSysMsg("查無資料!!!");
            Response.Write("<script language=javascript>window.close();</script>");
            return;
        }

        GridList.Text = GetTitle(dt.Rows[0]) + GetTable1(dt.Rows[0], strSql);
        //GetTableFooter();
        Util.OutputTxt(GridList.Text, "1", "donate_other_report8");
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
        //捐款期間
        //DateTime DonateDateS = Convert.ToDateTime(HFD_Year_month5.Value + "/" + HFD_Month_month5.Value + "/" + "1");
        //DateTime DonateDateE = Convert.ToDateTime(HFD_Year_month5.Value + "/" + HFD_Month_month5.Value + "/" + "1").AddMonths(1).AddDays(-1);
        string strTitle = "個別奉獻動機及收視管道統計分析<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'></span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人編號排序</span>";
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 19;
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

        DataTable dtRet = CaseUtil.Donate_Other_Report8_Print(dt);

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
                //cell.Width = "55mm";
                cell.Height = "40mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-left", ".5pt solid windowtext");
            }
            else if (iCtrl == 18)
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
                else if (Col == 18)
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
        string strSql = @"select Q.Donor_Id as [捐款人編號] ,D.Donor_Name as [捐款人姓名]
		                        ,CONVERT(VarChar,Q.Create_Date,111) as [問卷日期]
		                        ,Case When DonateMotive1='' Then '' Else 'V' End [支持媒體<br>宣教大平台，<br>可廣傳福音]
		                        ,Case When DonateMotive2='' Then '' Else 'V' End [個人靈命<br>得造就]
		                        ,Case When DonateMotive3='' Then '' Else 'V' End [支持優質<br>節目製作]
		                        ,Case When DonateMotive4='' Then '' Else 'V' End [支持GOOD TV<br>家庭事工]
		                        ,Case When DonateMotive5='' Then '' Else 'V' End [感恩奉獻]
		                        ,Case When WatchMode1='' Then '' Else 'V' End [GOOD TV<br>電視頻道]
		                        ,Case When WatchMode2='' Then '' Else 'V' End [官網]
		                        ,Case When WatchMode3='' Then '' Else 'V' End [Facebook]
		                        ,Case When WatchMode4='' Then '' Else 'V' End [Youtube]
		                        ,Case When WatchMode5='' Then '' Else 'V' End [好消息<br>月刊]
		                        ,Case When WatchMode6='' Then '' Else 'V' End [GOOD TV<br>簡介刊物]
		                        ,Case When WatchMode7='' Then '' Else 'V' End [教會牧者]
		                        ,Case When WatchMode8='' Then '' Else 'V' End [親友]
		                        ,Case When WatchMode9='' Then '' Else 'V' End [報章雜誌]
		                        ,Q.ToGOODTV as [給GOODTV<br>的話]
		                        ,DonateWay as [問卷來源]
		                        from Donate_OnlineQuestion Q
		                        left join Donor D on Q.Donor_Id=D.Donor_Id
		                        where 1=1";

        //搜尋條件
        if (HFD_DonorName_other8.Value != "")
        {
            strSql += " and D.Donor_Name like '%" + HFD_DonorName_other8.Value + "%' ";
        }
        if (HFD_DonateMotive1_other8.Value == "True")
        {
            strSql += " and DonateMotive1<>'' ";
        }
        if (HFD_DonateMotive2_other8.Value == "True")
        {
            strSql += " and DonateMotive2<>'' ";
        }
        if (HFD_DonateMotive3_other8.Value == "True")
        {
            strSql += " and DonateMotive3<>'' ";
        }
        if (HFD_DonateMotive4_other8.Value == "True")
        {
            strSql += " and DonateMotive4<>'' ";
        }
        if (HFD_DonateMotive5_other8.Value == "True")
        {
            strSql += " and DonateMotive5<>'' ";
        }
        if (HFD_DonateMotive6_other8.Value == "True")
        {
            strSql += " and Q.ToGOODTV <>'' ";
        }
        if (HFD_WatchMode1_other8.Value == "True")
        {
            strSql += " and WatchMode1<>'' ";
        }
        if (HFD_WatchMode2_other8.Value == "True")
        {
            strSql += " and WatchMode2<>'' ";
        }
        if (HFD_WatchMode3_other8.Value == "True")
        {
            strSql += " and WatchMode3<>'' ";
        }
        if (HFD_WatchMode4_other8.Value == "True")
        {
            strSql += " and WatchMode4<>'' ";
        }
        if (HFD_WatchMode5_other8.Value == "True")
        {
            strSql += " and WatchMode5<>'' ";
        }
        if (HFD_WatchMode6_other8.Value == "True")
        {
            strSql += " and WatchMode6<>'' ";
        }
        if (HFD_WatchMode7_other8.Value == "True")
        {
            strSql += " and WatchMode7<>'' ";
        }
        if (HFD_WatchMode8_other8.Value == "True")
        {
            strSql += " and WatchMode8<>'' ";
        }
        if (HFD_WatchMode9_other8.Value == "True")
        {
            strSql += " and WatchMode9<>'' ";
        }
        strSql += " Order BY Q.Donor_Id,convert(nvarchar(10),Q.Create_Date,111) ";
        
        return strSql;
    }
}