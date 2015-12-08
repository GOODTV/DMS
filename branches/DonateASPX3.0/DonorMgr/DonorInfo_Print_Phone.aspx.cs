using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;
public partial class DonorMgr_DonorInfo_Print_Phone : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        if(!IsPostBack)
        {
        String strSql = Session["strSql"].ToString();
        Print_Phone(strSql);
        }
    }
    //---------------------------------------------------------------------------
    private void Print_Phone(string strSql)
    {
        string [] spilt = new string[] {"From"};
        string[] strSql_temp = strSql.Split(spilt, StringSplitOptions.RemoveEmptyEntries);
        strSql = strSql.Replace("order by Donor_Id desc", "");
        strSql = strSql.Replace(strSql_temp[0],
                                 @" select Distinct Donor_Id , Donor_Id as 編號, 
                                           (Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End) as 捐款人, 
                                            Cellular_Phone as 手機號碼 ");
        strSql += " and Cellular_Phone != '' and Cellular_Phone is not null and IsMember<>'Y' ";
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
        Util.OutputTxt(GridList.Text, "1", "mobile");
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
        css.Add("font-family", "新細明體");
        css.Add("width", TableWidth);

        string strFontSize = "24px";

        string strTitle = "捐款人基本資料維護"; 
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 3;
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

        DataTable dtRet = CaseUtil.DonorInfo_Excel_Phone(dt);

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
                cell.Width = "80cm";
            }
            else
            {
                cell.Width = "80cm";
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