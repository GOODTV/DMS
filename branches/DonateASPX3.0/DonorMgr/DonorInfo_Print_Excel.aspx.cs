using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonorMgr_DonorInfo_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            String strSql = Session["strSql"].ToString();
            Print_Excel(strSql);
        }
    }
    //---------------------------------------------------------------------------
    private void Print_Excel(string strSql)
    {
        strSql = strSql.Replace("order by Donor_Id desc", "");
        strSql = strSql.Replace(", IDNo as [身分證/統編]", "");
        strSql = strSql.Replace(", Create_User as 建檔人員, Donor_Id_Old as 舊編號", "");
        strSql += " and IsMember<>'Y' ";
        //20140425 新增 by Ian_Kao 調整輸出欄位
        strSql = strSql.Replace("CONVERT(VARCHAR(10) ,Begin_DonateDate, 111) as 首捐日,", "");
        strSql = strSql.Replace("末捐日", "最近捐款日期");
        strSql = strSql.Replace("Donate_No as 捐款次數, REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Total),1),'.00','') as 累計捐款金額", "(Case When IsSendNews='Y' Then 'V' Else '' End) as 月刊,(Case When IsDVD='Y' Then 'V' Else '' End) as DVD,(Case When IsSendEpaper='Y' Then 'V' Else '' End) as 電子文宣,(Case When IsBirthday='Y' Then 'V' Else '' End) as 生日卡");
        //-------------------------------------
        strSql += "order by Donor_Id desc";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (Session["Donor_Id"] != "")
        {
            dict.Add("Donor_Id", Session["Donor_Id"].ToString());
        }
        //20140425 新增 by Ian_Kao
        if (Session["Donor_Name"] != "")
        {
            dict.Add("Donor_Name", Session["Donor_Name"].ToString());
        }
        if (Session["Invoice_Title"] != "")
        {
            dict.Add("Invoice_Title", Session["Invoice_Title"].ToString());
        }
        //-----------------------------
        if (Session["Donor_Id_Old"] != "")
        {
            dict.Add("Donor_Id_Old", Session["Donor_Id_Old"].ToString());
        }
        if (Session["Tel_Office"] != "")
        {
            dict.Add("Tel_Office", Session["Tel_Office"].ToString());
        }
        if (Session["ZipCode"] != "")
        {
            dict.Add("ZipCode", Session["ZipCode"].ToString());
        }
        if (Session["Street"] != "")
        {
            dict.Add("Street", Session["Street"].ToString());
        }
        if (Session["Lane"] != "")
        {
            dict.Add("Lane", Session["Lane"].ToString());
        }
        if (Session["Alley"] != "")
        {
            dict.Add("Alley", Session["Alley"].ToString());
        }
        if (Session["HouseNo"] != "")
        {
            dict.Add("HouseNo", Session["HouseNo"].ToString());
        }
        if (Session["HouseNoSub"] != "")
        {
            dict.Add("HouseNoSub", Session["HouseNoSub"].ToString());
        }
        if (Session["Floor"] != "")
        {
            dict.Add("Floor", Session["Floor"].ToString());
        }
        if (Session["FloorSub"] != "")
        {
            dict.Add("FloorSub", Session["FloorSub"].ToString());
        }
        if (Session["Room"] != "")
        {
            dict.Add("Room", Session["Room"].ToString());
        }
        if (Session["IsSendNewsNum"] != "")
        {
            dict.Add("IsSendNewsNum", Session["IsSendNewsNum"].ToString());
        }
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            SetSysMsg("查無資料!!!");
            Response.Redirect(Util.RedirectByTime("DonorInfo.aspx"));
        }
        
        GridList.Text = GetTitle(dt.Rows[0]) + GetTable1(dt.Rows[0], strSql);
        //GetTableFooter();
        Util.OutputTxt(GridList.Text, "1", "donor");
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

        string strFontSize = "24px";

        string strTitle = "捐款人基本資料維護";
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
    //---------------------------------------------------------------------------
    private string GetTable1(DataRow dr, string strSql)
    {
        //組 table
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (Session["Donor_Id"] != "")
        {
            dict.Add("Donor_Id", Session["Donor_Id"].ToString());
        }
        //20140425 新增 by Ian_Kao
        if (Session["Donor_Name"] != "")
        {
            dict.Add("Donor_Name", Session["Donor_Name"].ToString());
        }
        if (Session["Invoice_Title"] != "")
        {
            dict.Add("Invoice_Title", Session["Invoice_Title"].ToString());
        }
        //-----------------------------
        if (Session["Donor_Id_Old"] != "")
        {
            dict.Add("Donor_Id_Old", Session["Donor_Id_Old"].ToString());
        }
        if (Session["Tel_Office"] != "")
        {
            dict.Add("Tel_Office", Session["Tel_Office"].ToString());
        }
        if (Session["ZipCode"] != "")
        {
            dict.Add("ZipCode", Session["ZipCode"].ToString());
        }
        if (Session["Street"] != "")
        {
            dict.Add("Street", Session["Street"].ToString());
        }
        if (Session["Lane"] != "")
        {
            dict.Add("Lane", Session["Lane"].ToString());
        }
        if (Session["Alley"] != "")
        {
            dict.Add("Alley", Session["Alley"].ToString());
        }
        if (Session["HouseNo"] != "")
        {
            dict.Add("HouseNo", Session["HouseNo"].ToString());
        }
        if (Session["HouseNoSub"] != "")
        {
            dict.Add("HouseNoSub", Session["HouseNoSub"].ToString());
        }
        if (Session["Floor"] != "")
        {
            dict.Add("Floor", Session["Floor"].ToString());
        }
        if (Session["FloorSub"] != "")
        {
            dict.Add("FloorSub", Session["FloorSub"].ToString());
        }
        if (Session["Room"] != "")
        {
            dict.Add("Room", Session["Room"].ToString());
        }
        if (Session["IsSendNewsNum"] != "")
        {
            dict.Add("IsSendNewsNum", Session["IsSendNewsNum"].ToString());
        }
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;

        DataTable dtRet = CaseUtil.DonorInfo_Print(dt);

        table.Border = 1;
        table.CellPadding = 3;
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
            cell.Style.Add("background-color", " #FFE4C4 ");
            cell.Style.Add("border-left", ".5pt solid windowtext");
            cell.Style.Add("border-top", ".5pt solid windowtext");
            cell.Style.Add("border-right", ".5pt solid windowtext");
            cell.Style.Add("border-bottom", ".5pt solid windowtext");

            if (iCtrl == 0)
            {
                cell.Width = "30";
            }
            else
            {
                cell.Width = "90mm";
            }
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
}
