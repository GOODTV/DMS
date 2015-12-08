using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class Report_PrintReceiptTemp : Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //Session["ProgID"] = "Print";

        if (!IsPostBack)
        {
            string strSql = Session["SQLTmp"].ToString();
            //20131120 Add by GoodTV Tanya:增加更新收據列印紀錄
            string strUpdateSql = Session["SQLUpdateTmp"].ToString();
            if (strSql != "")
            {
                //Response.Write(strSql);                
                RrintReceipt(strSql,strUpdateSql);
                
            }
            else
            {
                Response.Write("empty string");
            }
        }
    }
    protected void RrintReceipt(string strSql, string strUpdateSql)
    {
        string s1 = @"<head>
                        <style>
                        @page Section2
                        {margin:5mm 5mm 5mm 5mm}
                        div.Section2
                        {page:Section2;}
                    </head>
                    <body>
                        <div class=Section2>
                    ";//邊界-上右下左

        GridList.Text = GetTable(strSql);

        //20131120 Add by GoodTV Tanya:增加更新收據列印紀錄
        if (strUpdateSql != "")
            UpdatePrintRecord(strUpdateSql);

        Util.OutputTxt(s1 + GridList.Text+"</div></body>", "2", DateTime.Now.ToString("yyyyMMddHHmmss"));
    }
    private string GetTable(string strSql)
    {
        string strRet = "";

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
        //HtmlTable table = new HtmlTable();
        //HtmlTableRow row;
        //HtmlTableCell cell;
        //CssStyleCollection css;

        //table.Border = 0;
        //table.CellPadding = 0;
        //table.CellSpacing = 0;
        //css = table.Style;
        //css.Add("font-size", "16px");
        //css.Add("font-family", "標楷體");
        //css.Add("width", "100%");
        //css.Add("border-collapse", "collpase");
        //string strFontSize = "18px";
        //--------------------------------------------
        //data
        int count = dt.Rows.Count;
        for (int i = 0; i < count; i++)
        {
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

            
            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            dr = dt.Rows[i];

            //Add by GoodTV Tanya:收據抬頭中文長度大於20，僅列印至20
            string Donor_Name = dr["Donor_Name"].ToString();
            if (isChinese(Donor_Name))
            {   if (Donor_Name.Length > 20)
                Donor_Name = Donor_Name.Substring(0, 20);                                    
            }
            else
            {
                if (Donor_Name.Length > 40)                
                    Donor_Name = Donor_Name.Substring(0, 40);
            }

            //Add by GoodTV Tanya:收件人中文長度大於27時，僅取到27
            //Modify by GoodTV Tanya:收件資料上移並將收件人設定為可2列字數拉長到中文長度27
            string Donor_Name1 = dr["Donor_Name1"].ToString();
            if (isChinese(Donor_Name1))
            {
                if (Donor_Name1.Length > 27)
                    Donor_Name1 = Donor_Name1.Substring(0, 11);
            }
            else
            {
                if (Donor_Name1.Length > 54)
                    Donor_Name1 = Donor_Name1.Substring(0, 22);
            }

            strTemp = dr["Donor_Name"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", strFontSize);
            css.Add("width", "100%");
            css.Add("margin-top", "11mm");  
            css.Add("margin-left", "21mm");
            css.Add("height", "17mm");       
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            //strTemp = dr["ZipCode"].ToString() + "　" + dr["City"].ToString() + dr["Area"].ToString() + "<br>" + dr["Address"].ToString();
            //若是海外地址，不帶Zipcode
            if (dr["IsAbroad_Invoice"].ToString() != "Y")
            {
                if (dr["Invoice_ZipCode"].ToString() != "")
                {
                    strTemp = dr["Invoice_ZipCode"].ToString();
                }
                else
                {
                    strTemp = dr["ZipCode"].ToString();
                }
            }
            else strTemp = "";
            if (dr["Invoice_Address"].ToString() != "")
            {
                strTemp += dr["Invoice_Address"].ToString();
                //Modify by GoodTV Tanya:收據地址加上「Attn」
                if (dr["Invoice_Attn"].ToString() != "")
                    strTemp += "(" + dr["Invoice_Attn"].ToString() + ")";
            }
            else
            {
                strTemp += dr["Address"].ToString();
                //Modify by GoodTV Tanya:收據地址加上「Attn」
                if (dr["Attn"].ToString() != "")
                    strTemp += "(" + dr["Attn"].ToString() + ")";
            }
            if (dr["IsAbroad_Invoice"].ToString() != "N")
            {
                strTemp += "&nbsp;" + dr["Invoice_OverseasCountry"].ToString();
            }
            

            cell.InnerHtml = strTemp == "" ? "&nbsp;" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            //css.Add("font-size", strFontSize);
            css.Add("font-size", "16px");
            css.Add("width", "100%");
            //css.Add("margin-top", "18mm");
            css.Add("margin-top", "12mm");
            //css.Add("margin-left", "105mm");
            //css.Add("margin-right", "40mm");
            css.Add("margin-left", "95mm");
            css.Add("margin-right", "30mm");
            //css.Add("height", "36mm");
            css.Add("height", "24mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();

            //mod by geo 20131227 for 收件人用Donor_Name1          
            if (dr["Title2"].ToString() != "")
            {
                if (dr["Title2"].ToString().IndexOf("寶號") > -1)
                    strTemp = Donor_Name1 + " 鈞啟";
                else
                    strTemp = Donor_Name1 + "&nbsp;" + dr["Title2"].ToString() + " 鈞啟";
            }
            else
            {
                if (dr["Title"].ToString().IndexOf("寶號") > -1)
                    strTemp = Donor_Name1 + " 鈞啟";
                else
                    strTemp = Donor_Name1 + "&nbsp;" + dr["Title"].ToString() + " 鈞啟";
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
            css.Add("text-align", "left");
            //css.Add("font-size", strFontSize);
            css.Add("font-size", "16px");
            css.Add("width", "100%");
            css.Add("margin-left", "95mm");
            css.Add("margin-right", "35mm");
            //css.Add("height", "6mm");       
            css.Add("height", "12mm");       
            row.Cells.Add(cell);
            table.Rows.Add(row);

            //row = new HtmlTableRow();
            //cell = new HtmlTableCell();
            //strTemp = dr["Invoice_No"].ToString();
            //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            //css = cell.Style;
            //css.Add("text-align", "left");
            //css.Add("font-size", "20px");
            //css.Add("width", "100%");
            //css.Add("line-height", "120%");
            //css.Add("margin-top", "52mm");
            //css.Add("margin-left", "89mm");
            //css.Add("height", "59mm");
            //row.Cells.Add(cell);
            //table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            strTemp = dr["Donor_Id"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "16px");
            css.Add("width", "100%");
            //css.Add("line-height", "120%");
            css.Add("margin-top", "10mm");
            css.Add("margin-left", "95mm");
            css.Add("height", "6mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            strTemp = dr["Invoice_No"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "20px");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            //css.Add("margin-top", "38mm");
            css.Add("margin-top", "42mm");
            css.Add("margin-left", "89mm");
            //css.Add("height", "39mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);


            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            //20131212Modify by GoodTV Tanya:應列印奉獻日期
            //DateTime dtDB = Util.GetDBDateTime();
            DateTime dtDonateDate = Convert.ToDateTime(dr["Donate_Date"].ToString());
            strTemp = dtDonateDate.Year.ToString() + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + dtDonateDate.Month.ToString("00") + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + dtDonateDate.Day.ToString("00");
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "20");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            //css.Add("margin-top", "3mm");
            css.Add("margin-left", "106mm");
            css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();          

            strTemp = Donor_Name;
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
            //string strChineseAmt = Util.ChineseMoney(Convert.ToInt64(Convert.ToDouble(dr["Donate_Amt"].ToString())));
            string strAmt = Convert.ToDouble(dr["Donate_Amt"].ToString()).ToString("C0");
            //strTemp = strChineseAmt + "&nbsp;" + strAmt;
            strTemp = strAmt;
            //1. 外幣(如果同時有新台幣及原幣) 顯示在 金額最後面的空位
            if (dr["Donate_Forign"].ToString().Trim() != "")
            {
                strTemp += "(" + dr["Donate_Forign"].ToString().Trim() + ")";
            }
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "20");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            css.Add("margin-left", "100mm");
            //add 20131105
            css.Add("margin-right", "30mm");
            css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            //2. 在性質 - 公益事業空位中增加捐款用途欄位
            //20131219Modify by GoodTV Tanya:此欄位改列印「收據備註」
            //strTemp = dr["Donate_Purpose"].ToString();
            strTemp = dr["Invoice_PrintComment"].ToString();            
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "20");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            css.Add("margin-left", "110mm");
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

            //row = new HtmlTableRow();
            //cell = new HtmlTableCell();
            //dr = dt.Rows[i];
            //strTemp = "&nbsp&nbsp&nbsp <br>";
            //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            //css = cell.Style;
            //css.Add("text-align", "right");
            //css.Add("font-size", "22");
            //css.Add("width", "100%");
            //css.Add("margin-right", "150px");
            //css.Add("height", "350px");
            //row.Cells.Add(cell);
            //table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            strTemp = Donor_Name;
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "16");
            css.Add("width", "100%");
            //css.Add("margin-top", "75mm");
            css.Add("margin-top", "66mm");
            css.Add("margin-left", "60mm");
            //css.Add("height", "79mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
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
            css.Add("font-size", "14px");
            css.Add("width", "100%");
            css.Add("margin-top", "4mm");
            css.Add("margin-left", "60mm");
            css.Add("margin-right", "108mm");
            css.Add("height", "28mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            strTemp = dr["Donor_Id"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "16px");
            css.Add("width", "100%");
            //css.Add("margin-top", "7mm");
            css.Add("margin-left", "60mm");
            css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            //row = new HtmlTableRow();
            //cell = new HtmlTableCell();
            //dr = dt.Rows[i];
            //strTemp = "&nbsp&nbsp&nbsp <br>";
            //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            //css = cell.Style;
            //css.Add("text-align", "right");
            //css.Add("font-size", "22");
            //css.Add("width", "100%");
            //css.Add("margin-right", "150px");
            //css.Add("height", "350px");
            //row.Cells.Add(cell);
            //table.Rows.Add(row);

            //轉成 html 碼
            StringWriter sw = new StringWriter();

            HtmlTextWriter htw = new HtmlTextWriter(sw);
            table.RenderControl(htw);

            strRet += htw.InnerWriter.ToString();

            if (i != dt.Rows.Count - 1)
            {
               //插入分頁符號
               strRet += "<br clear=all style='mso-special-character:line-break;page-break-before:always'>";
            }
        }
        //--------------------------------------------

        ////轉成 html 碼
        //StringWriter sw = new StringWriter();

        //HtmlTextWriter htw = new HtmlTextWriter(sw);
        //table.RenderControl(htw);

        //return htw.InnerWriter.ToString();
        return strRet;
    } //end of GetTable()

    //20131120 Add by GoodTV Tanya:增加更新收據列印紀錄
    protected void UpdatePrintRecord(string strSql)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        NpoDB.ExecuteSQLS(strSql, dict);
    }

    //20140106 Add by GoodTV Tanya:增加中文字驗證
    public static bool isChinese(string strChinese) 
    { 
        bool bresult = true; 
        int dRange = 0; 
        int dstringmax=Convert.ToInt32("9fff", 16); 
        int dstringmin=Convert.ToInt32("4e00", 16); 
        for (int i = 0; i < strChinese.Length; i++)
        { 
            dRange = Convert.ToInt32(Convert.ToChar(strChinese.Substring(i, 1))); 
            if (dRange >= dstringmin && dRange <dstringmax)
            { 
                bresult = true; 
                break; 
            } 
            else 
            { 
                bresult = false; 
            } 
        } 
        return bresult; 
    } 
}