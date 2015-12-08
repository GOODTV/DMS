using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class Report_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "Print";

        if (!IsPostBack)
        {
            LoadDropDownListData();
        }
    }
    public void LoadDropDownListData()
    {
        //縣市
        Util.FillDropDownList(ddlCity, Util.GetDataTable("CodeCity", "ParentCityID", "0", "", ""), "Name", "ZipCode", false);
        ddlCity.Items.Insert(0, new ListItem("", ""));
        ddlCity.SelectedIndex = 0;

        ddlArea.Items.Insert(0, new ListItem("", ""));
        ddlArea.SelectedIndex = 0;
    }
    public void CityToArea(DropDownList ddlTo, DropDownList ddlFrom)
    {
        Util.FillDropDownList(ddlTo, Util.GetDataTable("CodeCity", "ParentCityID", ddlFrom.SelectedValue, "", ""), "Name", "ZipCode", false);
        ddlTo.Items.Insert(0, new ListItem("", ""));
        ddlTo.SelectedIndex = 0;
    }
    protected void ddlCity_SelectedIndexChanged(object sender, EventArgs e)
    {
        CityToArea(ddlArea, ddlCity);
    }
    protected void btnWord_Click(object sender, EventArgs e)
    {
        string s1 = @"<head>
                        <style>
                        @page Section2
                        {margin:1.65cm 1.5cm 0.4cm 0.8cm}
                        div.Section2
                        {page:Section2;}
                        </style>
                    </head>
                    <body>
                        <div class=Section2>
                    ";//邊界-上右下左
        Dictionary<string, int> ReportData = new Dictionary<string, int>();
        GridList.Text = GetTable(ReportData);
        Util.OutputTxt(s1+GridList.Text, "2", DateTime.Now.ToString("yyyyMMddHHmmss"));
    }
    private string GetTable(Dictionary<string, int> ReportData)
    {
        string strSql = stringSql();
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr;

        //組 table
        string strTemp = "";
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
        css.Add("width", "100%");
        css.Add("border-collapse", "collpase");
        string strFontSize = "18px";
        //--------------------------------------------
        //data
        int count = dt.Rows.Count;
        for (int i = 0; i < count; i++)
        {
            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            dr = dt.Rows[i];
            strTemp = dr["Donor_Name"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", strFontSize);
            css.Add("width", "100%");
            css.Add("margin-left", "80px");
            css.Add("height", "0px");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            dr = dt.Rows[i];
            strTemp = dr["ZipCode"].ToString() + "　" + dr["City"].ToString() + dr["Area"].ToString() + "<br>" + dr["Address"].ToString();
            if (dr["Invoice_Title"].ToString().Trim() != "")
            {
                strTemp += "<br>" + dr["Invoice_Title"].ToString() + " &nbsp " + dr["Title2"].ToString() + "收";
            }
            else
            {
                strTemp += "<br>" + dr["Donor_Name"].ToString() + " &nbsp " + dr["Title"].ToString() + "收";
            }
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "right");
            css.Add("font-size", strFontSize);
            css.Add("width", "100%");
            css.Add("margin-right", "150px");
            css.Add("height", "230px");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            dr = dt.Rows[i];
            strTemp = dr["Invoice_No"].ToString() + "&nbsp&nbsp&nbsp <br>" + dr["year"].ToString() + "&nbsp &nbsp&nbsp&nbsp " + dr["month"].ToString() + " &nbsp&nbsp&nbsp " + dr["day"].ToString() + "<br>" + dr["Donor_Name"].ToString() + "&nbsp&nbsp&nbsp &nbsp&nbsp&nbsp<br>" + dr["money"].ToString() + "&nbsp &nbsp &nbsp &nbsp &nbsp<br>" + dr["Remark"].ToString() + "&nbsp &nbsp &nbsp &nbsp &nbsp";
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "right");
            css.Add("font-size", "20");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            css.Add("margin-right", "150px");
            css.Add("height", "420px");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            dr = dt.Rows[i];
            strTemp = "&nbsp&nbsp&nbsp <br>";
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "right");
            css.Add("font-size", "22");
            css.Add("width", "100%");
            css.Add("margin-right", "150px");
            css.Add("height", "350px");
            row.Cells.Add(cell);
            table.Rows.Add(row);
        }
        //--------------------------------------------

        //轉成 html 碼
        StringWriter sw = new StringWriter();

        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    } //end of GetTable()
    //*****查詢語法******//
    public string stringSql()
    {
   
        string s = @"
                     select D.Donor_Name,D.Title,D.Invoice_Title,D.Title2,D.ZipCode,D.City,D.Area,D.Address,D_D.Invoice_No
                           ,D_D.Invoice_No,Year(D_D.Donate_Date) as year,Month(D_D.Donate_Date) as month,Day(D_D.Donate_Date) as day,D_D.Donate_Amt as money,D.Remark
                     from Donate D_D 
                     left join Donor D 
                     on D.Donor_Id=D_D.Donor_Id
                     where 1=1";
//D_D.Invoice_Type='單次收據' ";
        if (txtDateFrom.Text.Trim() != "" && txtDateTo.Text.Trim() != "")
        {
            s += " and D_D.Donate_Date between'" + txtDateFrom.Text.Trim() + "' and '" + txtDateTo.Text.Trim() + "'";
        }
        if (ddlCity.SelectedItem.Text != "")
        {
            s += " and D.City='" + ddlCity.SelectedItem.Text + "'";
        }
        if (ddlArea.SelectedItem.Text != "")
        {
            s += " and D.Area='" + ddlArea.SelectedItem.Text + "'";
        }
        if (txtName.Text.Trim() != "")
        {
            s += " and D.Donor_Name LIKE'%" + txtName.Text.Trim() + "%' ";
        }
        s += ";";
        return s;
    }
}