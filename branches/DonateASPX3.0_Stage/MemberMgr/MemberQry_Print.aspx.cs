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

public partial class MemberMgr_MemberQry_Print : BasePage
{
    string TableWidth = "800px";
    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        PrintReport();
    }
    //---------------------------------------------------------------------------
    private void PrintReport()
    {
        string strSql = Session["strSql"].ToString();
        if (strSql == "")
        {
            ShowSysMsg("請先搜尋過再列表！");
            Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Close();", true);
            return;
        }
        string[] spilt = new string[] { "From" };
        string[] strSql_temp = strSql.Split(spilt, StringSplitOptions.RemoveEmptyEntries);
        strSql = strSql.Replace(strSql_temp[0],
                                @"select Donor_Id , Donor_Id as 編號, Donor_Name as [讀者姓名], Donor_Type as [身分別],
                                         (Case When Cellular_Phone<>'' Then Cellular_Phone Else Case When Tel_Office<>'' Then Tel_Office Else Tel_Home End End) as [連絡電話], 
                                         Cellular_Phone as [手機號碼], (Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) as [通訊地址],
                                         (Case When IsSendEpaper='Y' Then 'V' Else '' End) as [電子文宣]");
        DataTable dt = NpoDB.QueryGetTable(strSql);
        int count = dt.Rows.Count;
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
            return;
        }

        DataTable dtRet = CaseUtil.MemberQry_Print(dt);

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        table.Width = "100%";
        css = table.Style;
        css.Add("font-size", "12px");
        css.Add("font-family", "細明體");
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

            //if (iCtrl == 0)
            //{
            //    cell.Width = "30";
            //}
            //else
            //{
            //    cell.Width = "90mm";
            //}
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

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

        GridList.Text = "會員資料明細表<br/><br/>" + htw.InnerWriter.ToString();
    }
}