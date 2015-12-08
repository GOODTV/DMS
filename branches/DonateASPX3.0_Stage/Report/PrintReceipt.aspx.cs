using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class Report_PrintReceipt : Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //Session["ProgID"] = "Print";

        if (!IsPostBack)
        {
            

            //取得request form提供之sql string
            string strSql = HttpUtility.UrlDecode(Request.Form["SQL"].ToString(), System.Text.Encoding.UTF8);
            string strUpdateSql = HttpUtility.UrlDecode(Request.Form["UpdateSQL"].ToString(), System.Text.Encoding.UTF8);
            //string strSql = stringSql();
            if (strSql != "")
            {
                //Response.Write(strSql);

                //因Server啟動時間差，需再導至另一頁處理
                Session["SQLTmp"] = strSql;
                //20131120 Add by GoodTV Tanya:更新收據列印紀錄  
                Session["SQLUpdateTmp"] = "";
                if(strUpdateSql != "")
                    Session["SQLUpdateTmp"] = strUpdateSql;

                Response.Redirect("PrintReceiptTemp.aspx");
            }
            else
            {
                Response.Write("empty string");
            }
        }
    }
    protected void RrintReceipt(string strSql)
    {
        string s1 = @"<head>
                        <style>
                        @page Section2
                        {margin:5mm 5mm 5mm 5mm}
                        div.Section2
                        {page:Section2;}
                        </style>
                    </head>
                    <body>
                        <div class=Section2>
                    ";//邊界-上右下左

        GridList.Text = GetTable(strSql);

        Response.Write("empty string");

        Response.Redirect("PrintReceiptTemp.aspx");

        //Util.OutputTxt(s1 + GridList.Text, "2", DateTime.Now.ToString("yyyyMMddHHmmss"));
    }
    private string GetTable(string strSql)
    {
        //string strSql = stringSql();

        Dictionary<string, object> ReportData = new Dictionary<string, object>();
        //ReportData.Add("DateFrom", txtDateFrom.Text);
        //ReportData.Add("DateTo", txtDateTo.Text);
        //ReportData.Add("Name", "%" + txtName.Text + "%");
        //ReportData.Add("City", ddlCity.SelectedValue);
        //ReportData.Add("Area", ddlArea.SelectedValue);

        DataTable dt = NpoDB.GetDataTableS(strSql, ReportData);
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
            css.Add("margin-top", "11mm");  
            css.Add("margin-left", "21mm");
            css.Add("height", "6mm");       
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            //strTemp = dr["ZipCode"].ToString() + "　" + dr["City"].ToString() + dr["Area"].ToString() + "<br>" + dr["Address"].ToString();
            if (dr["Invoice_ZipCode"].ToString() != "")
            {
                strTemp = dr["Invoice_ZipCode"].ToString();
            }
            else
            {
                strTemp = dr["ZipCode"].ToString();
            }
            if (dr["Invoice_Address"].ToString() != "")
            {
                strTemp += dr["Invoice_Address"].ToString();
            }
            else
            {
                strTemp += dr["Address"].ToString();
            }
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", strFontSize);
            css.Add("width", "100%");
            css.Add("margin-top", "18mm");      
            css.Add("margin-left", "105mm");
            css.Add("margin-right", "40mm");
            css.Add("height", "18mm");          
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            if (dr["Title2"].ToString() != "")
            {
                strTemp = dr["Donor_Name"].ToString() + " &nbsp " + dr["Title2"].ToString() + "收";
            }
            else
            {
                strTemp = dr["Donor_Name"].ToString() + " &nbsp " + dr["Title"].ToString() + "收";
            }
            //if (dr["Invoice_Title"].ToString().Trim() != "")
            //{
            //    strTemp += "<br>" + dr["Invoice_Title"].ToString() + " &nbsp " + dr["Title2"].ToString() + "收";
            //}
            //else
            //{
            //    strTemp += "<br>" + dr["Donor_Name"].ToString() + " &nbsp " + dr["Title"].ToString() + "收";
            //}
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "right");
            css.Add("font-size", strFontSize);
            css.Add("width", "100%");
            css.Add("margin-left", "105mm");
            css.Add("margin-right", "40mm");
            css.Add("height", "6mm");       
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            strTemp = dr["Invoice_No"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "20");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            css.Add("margin-top", "52mm");
            css.Add("margin-left", "89mm");
            css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            DateTime dtDB = Util.GetDBDateTime();
            strTemp = dtDB.Year.ToString() + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + dtDB.Month.ToString("00") + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + dtDB.Day.ToString("00");
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "20");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            css.Add("margin-left", "106mm");
            css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            strTemp = dr["Donor_Name"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "20");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            css.Add("margin-left", "89mm");
            css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            string strChineseAmt = Util.ChineseMoney(Convert.ToInt64(Convert.ToDouble(dr["Donate_Amt"].ToString())));
            string strAmt = Convert.ToDouble(dr["Donate_Amt"].ToString()).ToString("C0");
            strTemp = strChineseAmt + "&nbsp;" + strAmt;
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "20");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            css.Add("margin-left", "100mm");
            css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            //row = new HtmlTableRow();
            //cell = new HtmlTableCell();
            //dr = dt.Rows[i];
            //string strChineseAmt=Util.ChineseMoney(Convert.ToInt64(Convert.ToDouble(dr["Donate_Amt"].ToString())));
            //string strAmt = Convert.ToDouble(dr["Donate_Amt"].ToString()).ToString("C0");
            //DateTime dtDB = Util.GetDBDateTime();
            //strTemp = dr["Invoice_No"].ToString() + "&nbsp&nbsp&nbsp <br>" + dtDB.Year.ToString() + "&nbsp &nbsp&nbsp&nbsp " + dtDB.Month.ToString() + " &nbsp&nbsp&nbsp " + dtDB.Day.ToString() + "<br>" + dr["Donor_Name"].ToString() + "&nbsp&nbsp&nbsp &nbsp&nbsp&nbsp<br>" + strChineseAmt + strAmt + "&nbsp &nbsp &nbsp &nbsp &nbsp<br>&nbsp &nbsp &nbsp &nbsp &nbsp";
            //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            //css = cell.Style;
            //css.Add("text-align", "right");
            //css.Add("font-size", "20");
            //css.Add("width", "100%");
            //css.Add("line-height", "120%");
            //css.Add("margin-right", "150px");
            //css.Add("height", "420px");
            //row.Cells.Add(cell);
            //table.Rows.Add(row);

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

    //------------------------------------------------------------------------------------------
    //*****查詢語法******//
    public string stringSql()
    {
   
//        string s = @"
//                     select D.Donor_Name,D.Title,D.Invoice_Title,D.Title2,D.ZipCode,D.City,D.Area,D.Address,D_D.Invoice_No
//                           ,D_D.Invoice_No,Year(D_D.Donate_Date) as year,Month(D_D.Donate_Date) as month,Day(D_D.Donate_Date) as day,D_D.Donate_Amt as money,D.Remark
//                     from Donate D_D 
//                     left join Donor D 
//                     on D.Donor_Id=D_D.Donor_Id
//                     where 1=1 
//                     --and D_D.Invoice_Type='單次收據' 
//            ";

//        string s = @"
//Select Donor.Attn
//	,Donate_Id
//	,Invoice_No=(Case When Donate.Invoice_No_Old<>'' Then Donate.Invoice_No_Old Else Invoice_Pre+Invoice_No End),Donate_Date=CONVERT(VarChar,Donate_Date,111),Donor_Name=(Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End)
//	,Title
//	,Title2
//	,ZipCode
//	,Address=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Address Else B.mValue+Address End End),Address2=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+DONOR.ZipCode+C.mValue+Address Else B.mValue+DONOR.ZipCode+Address End End)
//	,Invoice_ZipCode,Invoice_Address=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+E.mValue+Invoice_Address Else D.mValue+Invoice_Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then DONOR.Invoice_ZipCode+D.mValue+E.mValue+Invoice_Address Else DONOR.Invoice_ZipCode+D.mValue+Invoice_Address End End),Donate_Amt,Donate_Amt2,Donate_Desc=CONVERT(nvarchar(4000),Donate_Desc),Donate.Dept_Id,Invoice_IDNo,DONATE.Donor_Id,Donate_Purpose,Donate_Payment,Invoice_Print,Invoice_Print_Add,Post_Donor_Name=Donor_Name,TEL=(Case When DONOR.Cellular_Phone<>'' Then DONOR.Cellular_Phone Else Case When DONOR.Tel_Office<>'' Then DONOR.Tel_Office Else DONOR.Tel_Home End End), Invoice_PrintComment=CONVERT(nvarchar(4000),Invoice_PrintComment),Comment=CONVERT(nvarchar(4000),Comment),Invoice_No_Old,Donate.Invoice_Type,Donate.Create_User,Donate_Forign From DONATE Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id Left Join ACT On DONATE.Act_Id=ACT.Act_Id Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode Where Donate_Purpose_Type In ('D') And Invoice_No<>'' And Donate.Invoice_Type In ('年度證明及收據') And Donate.Dept_Id = 'C001' And Invoice_Print='1' And (Issue_Type='M' Or Issue_Type='' Or Issue_Type Is null) And Donate_Payment Not In ('物資捐贈') Order By Donate.Dept_Id,Invoice_No 
//             ";

        string s = @"
Select Donor.Attn,Donate_Id,Invoice_No=(Case When Donate.Invoice_No_Old<>'' Then Donate.Invoice_No_Old Else Invoice_Pre+Invoice_No End),Donate_Date=CONVERT(VarChar,Donate_Date,111),Donor_Name=(Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End),Title,Title2,ZipCode,Address=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+C.mValue+Address Else B.mValue+Address End End),Address2=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+DONOR.ZipCode+C.mValue+Address Else B.mValue+DONOR.ZipCode+Address End End),Invoice_ZipCode,Invoice_Address=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then D.mValue+E.mValue+Invoice_Address Else D.mValue+Invoice_Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When D.mValue<>E.mValue Then DONOR.Invoice_ZipCode+D.mValue+E.mValue+Invoice_Address Else DONOR.Invoice_ZipCode+D.mValue+Invoice_Address End End),Donate_Amt,Donate_Amt2,Donate_Desc=CONVERT(nvarchar(4000),Donate_Desc),Donate.Dept_Id,Invoice_IDNo,DONATE.Donor_Id,Donate_Purpose,Donate_Payment,Invoice_Print,Invoice_Print_Add,Post_Donor_Name=Donor_Name,TEL=(Case When DONOR.Cellular_Phone<>'' Then DONOR.Cellular_Phone Else Case When DONOR.Tel_Office<>'' Then DONOR.Tel_Office Else DONOR.Tel_Home End End), Invoice_PrintComment=CONVERT(nvarchar(4000),Invoice_PrintComment),Comment=CONVERT(nvarchar(4000),Comment),Invoice_No_Old,Donate.Invoice_Type,Donate.Create_User,Donate_Forign From DONATE Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id Left Join ACT On DONATE.Act_Id=ACT.Act_Id Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode Where Donate_Purpose_Type In ('D') And Invoice_No<>'' And Donate.Invoice_Type In ('年度證明及收據') And Donate.Dept_Id = 'C001' And (Invoice_Print='0' Or Invoice_Print Is Null) And (Issue_Type='M' Or Issue_Type='' Or Issue_Type Is null) And Donate_Payment Not In ('物資捐贈') Order By Donate.Dept_Id,Invoice_No 
  
                    ";
        return s;
    }
        
}