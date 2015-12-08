using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class Invoice_Print_Data : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //Session["ProgID"] = "Print";

        if (!IsPostBack)
        {
            string strSql = Session["strSql_Print"].ToString();
            if (strSql != "")
            {
                //Response.Write(strSql);
                RrintReceipt(strSql);
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

        //Edit_Donate
        string strSql_Edit = Session["strSql_Edit"].ToString();
        if (strSql_Edit.Contains("do.Invoice_Print = '0' or do.Invoice_Print is null") == true) //20150325增加 更新 首印經手人和最後補印經手人
        {
            DataTable dt2 = NpoDB.QueryGetTable(strSql_Edit);
            DataRow dr2;
            int count2 = dt2.Rows.Count;
            for (int i = 0; i < count2; i++)
            {
                dr2 = dt2.Rows[i];
                //Edit
                //****變數宣告****//
                Dictionary<string, object> dict = new Dictionary<string, object>();

                //****設定SQL指令 +更新Invoice_Print_Date****//
                string strSql_Donate = " update Donate set ";

                strSql_Donate += "  Invoice_Print = @Invoice_Print";
                strSql_Donate += " ,Invoice_Print_Date = @Invoice_Print_Date";
                strSql_Donate += " ,Invoice_FirstPrint_User = @Invoice_FirstPrint_User";
                strSql_Donate += " where Donate_Id = @Donate_Id";

                dict.Add("Invoice_Print", "1");
                dict.Add("Invoice_Print_Date", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
                dict.Add("Invoice_FirstPrint_User", SessionInfo.UserName);
                dict.Add("Donate_Id", dr2["捐款編號"].ToString());
                NpoDB.ExecuteSQLS(strSql_Donate, dict);
            }
        }
        else
        { 
            DataTable dt2 = NpoDB.QueryGetTable(strSql_Edit);
            DataRow dr3;
            int count2 = dt2.Rows.Count;
            for (int i = 0; i < count2; i++)
            {
                dr3 = dt2.Rows[i];
                //Edit
                //****變數宣告****//*/
                Dictionary<string, object> dict = new Dictionary<string, object>();
                //****設定SQL指令 更新Invoice_RePrint_Date****/
                string strSql_Date = " update Donate set ";

                strSql_Date += "  Invoice_RePrint_Date = @Invoice_RePrint_Date";
                strSql_Date += " ,Invoice_LastPrint_User = @Invoice_LastPrint_User";
                strSql_Date += " where Donate_Id = @Donate_Id";

                dict.Add("Invoice_RePrint_Date", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
                dict.Add("Invoice_LastPrint_User", SessionInfo.UserName);
                dict.Add("Donate_Id", dr3["捐款編號"].ToString());
                NpoDB.ExecuteSQLS(strSql_Date, dict);
            }
        }

        Util.OutputTxt(s1 + GridList.Text + "</div></body>", "2", DateTime.Now.ToString("yyyyMMddHHmmss"));
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

        
        //--------------------------------------------
        //data
        int count = dt.Rows.Count;
        for (int i = 0; i < count; i++)
        {
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
            //string strFontSize = "18px";

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            dr = dt.Rows[i];
            
            //Add by GoodTV Tanya:收據抬頭中文長度大於20，僅列印至20
            string Donor_Name = dr["捐款人"].ToString();
            if (isChinese(Donor_Name))
            {
                if (Donor_Name.Length > 20)
                    Donor_Name = Donor_Name.Substring(0, 20);
            }
            else
            {
                if (Donor_Name.Length > 40)
                    Donor_Name = Donor_Name.Substring(0, 40);
            }

            //Add by GoodTV Tanya:收件人中文長度大於27時，僅取到27
            //Modify by GoodTV Tanya:收件資料上移並將收件人設定為可2列字數拉長到中文長度27
            string Donor_Name1 = dr["捐款人"].ToString();
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

            strTemp = dr["捐款人"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            //css.Add("font-size", strFontSize);
            css.Add("font-size", "13.5pt");
            css.Add("width", "100%");
            css.Add("margin-top", "11mm");
            css.Add("margin-left", "21mm");
            css.Add("height", "6mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            //strTemp = dr["ZipCode"].ToString() + "　" + dr["City"].ToString() + dr["Area"].ToString() + "<br>" + dr["Address"].ToString();
            //若是海外地址，不帶Zipcode
            if (dr["IsAbroad_Invoice"].ToString() != "Y")
            {

                if (dt.Columns["Invoice_ZipCode"] != null && dr["Invoice_ZipCode"].ToString() != "")
                {
                    strTemp = dr["Invoice_ZipCode"].ToString();
                }
                else
                {

                    strTemp = dr["郵遞區號"].ToString();
                }

            }
            else strTemp = "";
            if (dt.Columns["Invoice_Address"] != null && dr["Invoice_Address"].ToString() != "")
            {
                strTemp += dr["Invoice_Address"].ToString();
                //Modify by GoodTV Tanya:收據地址加上「Attn」
                if (dr["Invoice_Attn"].ToString() != "")
                    strTemp += "(" + dr["Invoice_Attn"].ToString() + ")";
            }
            else
            {
                strTemp += dr["地址"].ToString();
                //Modify by GoodTV Tanya:通訊地址加上「Attn」
                if (dr["Attn"].ToString() != "")
                    strTemp += "(" + dr["Attn"].ToString() + ")";
            }
            //strTemp = dr["郵遞區號"].ToString();
            //strTemp += dr["地址"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp;" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "12pt");
            //css.Add("font-size", "16px");
            css.Add("width", "100%");
            css.Add("margin-top", "12mm");
            css.Add("margin-left", "95mm");
            css.Add("margin-right", "30mm");
            css.Add("height", "28mm");//24mm
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            //if (dr["Title2"].ToString() != "")
            //{
            //    strTemp = dr["Donor_Name"].ToString() + " &nbsp " + dr["Title2"].ToString() + "收";
            //}
            //else
            //{
            //    strTemp = dr["Donor_Name"].ToString() + " &nbsp " + dr["Title"].ToString() + "收";
            //}
            //if (dr["Invoice_Title"].ToString().Trim() != "")
            //{
            //    strTemp += "<br>" + dr["Invoice_Title"].ToString() + " &nbsp " + dr["Title2"].ToString() + "收";
            //}
            //else
            //{
            //    strTemp += "<br>" + dr["Donor_Name"].ToString() + " &nbsp " + dr["Title"].ToString() + "收";
            //}
            //mod by geo 20131227 for 收件人用Donor_Name1          
            if (dr["稱謂2"].ToString() != "")
            {
                if (dr["稱謂2"].ToString().IndexOf("寶號") > -1)
                    strTemp = Donor_Name1 + " 鈞啟";
                else
                    strTemp = Donor_Name1 + "&nbsp;" + dr["稱謂2"].ToString() + " 鈞啟";
            }
            else
            {
                if (dr["稱謂"].ToString().IndexOf("寶號") > -1)
                    strTemp = Donor_Name1 + " 鈞啟";
                else
                    strTemp = Donor_Name1 + "&nbsp;" + dr["稱謂"].ToString() + " 鈞啟";
            }
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "12pt");
            //css.Add("font-size", "16px");
            css.Add("width", "100%");
            css.Add("margin-left", "95mm");
            css.Add("margin-right", "35mm");
            css.Add("height", "10mm"); //12mm
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
            strTemp = dr["編號"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "12pt");
            //css.Add("font-size", "16px");
            css.Add("width", "100%");
            css.Add("margin-top", "10mm");
            css.Add("margin-left", "95mm");
            css.Add("height", "8mm");//6mm
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            strTemp = dr["收據編號"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "15pt");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            css.Add("margin-top", "42mm");
            css.Add("margin-left", "89mm");
            //css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            //20131212Modify by GoodTV Tanya:應列印奉獻日期
            //DateTime dtDB = Util.GetDBDateTime();
            // 2014/5/22 捐款日期為null的防呆
            if (dr.IsNull("Donate_Date"))
            {
                strTemp = "";
            }
            else
            {
                DateTime dtDonateDate = Convert.ToDateTime(dr["Donate_Date"].ToString());
                strTemp = dtDonateDate.Year.ToString() + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + dtDonateDate.Month.ToString("00") + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + dtDonateDate.Day.ToString("00");
            }
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "15pt");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            css.Add("margin-left", "106mm");
            css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            //strTemp = dr["捐款人"].ToString();
            //20140821改為 收據抬頭
            strTemp = dr["Invoice_Title"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "15pt");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            css.Add("margin-left", "89mm");
            css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            //string strChineseAmt = Util.ChineseMoney(Convert.ToInt64(Convert.ToDouble(dr["Donate_Amt"].ToString())));
            //國字大寫
            //string Donate_Amt = GetMoneyUpper(Decimal.Parse(dr["金額"].ToString().Replace(",", "") + ".00"));

            //2014/5/22 修改的金額欄位為null會出現錯誤
            string strAmt = "";
            if (!dr.IsNull("金額"))
            {
                strAmt = Convert.ToDouble(dr["金額"].ToString()).ToString("C0");
            }
            //strTemp = strChineseAmt + "&nbsp;" + strAmt;
            //strTemp = Donate_Amt + " NT$" + dr["金額"].ToString();
            strTemp = strAmt;
            //1. 外幣(如果同時有新台幣及原幣) 顯示在 金額最後面的空位
            if (dr["外幣"].ToString().Trim() != "" && dr["外幣金額"].ToString().Trim() != "")
            {
                strTemp += "(" + dr["外幣"].ToString().Trim() + Convert.ToDouble(dr["外幣金額"].ToString()) + ")";
            }
            else if (dr["外幣"].ToString().Trim() != "" && dr["外幣金額"].ToString().Trim() != "0")
            {
                strTemp += "(" + dr["外幣"].ToString().Trim() + ")";
            }
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "15pt");
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
            css.Add("font-size", "15pt");
            css.Add("width", "100%");
            css.Add("line-height", "120%");
            css.Add("margin-left", "110mm");
            css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            //strTemp = Donor_Name;
            //20140821改為 收據抬頭
            strTemp = dr["Invoice_Title"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "12pt");
            css.Add("width", "100%");
            //css.Add("margin-top", "75mm");
            css.Add("margin-top", "66mm");
            css.Add("margin-left", "60mm");
            //css.Add("height", "79mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            if (dt.Columns["Invoice_ZipCode"] != null && dr["Invoice_ZipCode"].ToString() != "")
            {
                strTemp = dr["Invoice_ZipCode"].ToString();
            }
            else
            {
                strTemp = dr["郵遞區號"].ToString();
            }
            if (dt.Columns["Invoice_Address"] != null && dr["Invoice_Address"].ToString() != "")
            {
                strTemp += dr["Invoice_Address"].ToString();
                //2014/5/23 劃撥的收據地址也加上「Attn」
                if (dr["Invoice_Attn"].ToString() != "")
                    strTemp += "(" + dr["Invoice_Attn"].ToString() + ")";
            }
            else
            {
                strTemp += dr["地址"].ToString();
                //2014/5/23 劃撥的通訊地址也加上「Attn」
                if (dr["Attn"].ToString() != "")
                    strTemp += "(" + dr["Attn"].ToString() + ")";
            }
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "10.5pt");
            css.Add("width", "100%");
            css.Add("margin-top", "4mm");
            css.Add("margin-left", "60mm");
            css.Add("margin-right", "108mm");
            css.Add("height", "28mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            strTemp = "捐款人編號：" + dr["編號"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("text-align", "left");
            css.Add("font-size", "12pt");
            css.Add("width", "100%");
            //css.Add("margin-top", "7mm");
            css.Add("margin-left", "60mm");
            css.Add("height", "7mm");
            row.Cells.Add(cell);
            table.Rows.Add(row);

            //轉成 html 碼
            StringWriter sw = new StringWriter();

            HtmlTextWriter htw = new HtmlTextWriter(sw);
            table.RenderControl(htw);

            strRet += htw.InnerWriter.ToString();

            if (i != count - 1)
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
    public static string GetMoneyUpper(decimal d)
    {
        string o = d.ToString();
        string dq = "", dh = "";
        if (o.Contains("."))
        {
            dq = o.Split('.')[0];
            dh = o.Split('.')[1];
        }
        else
        {
            dq = o;
        }
        string ret = GetMoney(dq, true, "圓") + GetMoney(dh, false, "");
        if (ret.Contains("厘") || ret.Contains("分"))
            return ret;
        return ret + "整";
    }
    private static string GetMoney(string number, bool left, string lastdw)
    {
        string[] NTD = new string[10] { "零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖" };
        string[] DW = new string[8] { "厘", "分", "角", "", "拾", "佰", "仟", "萬" };
        int first = 4;
        string str = "";
        if (!left)
        {
            first = 1;
            if (number.Length == 1)
            {
                number += "00";
            }
            else if (number.Length == 2)
            {
                number += "0";
            }
            else number = number.Substring(0, 3);

        }
        else
        {
            if (number.Length >= 9)
            {
                return GetMoney(number.Substring(0, number.Length - 8), true, "億") + GetMoney(number.Substring(number.Length - 8, 8), true, "圓");
            }
            if (number.Length >= 5)
            {
                return GetMoney(number.Substring(0, number.Length - 4), true, "萬") + GetMoney(number.Substring(number.Length - 4, 4), true, "圓");
            }
        }
        bool has0 = false;
        for (int i = 0; i < number.Length; ++i)
        {
            int w = number.Length - i + first - 2;
            if (int.Parse(number[i].ToString()) == 0)
            {
                has0 = true;
                continue;
            }
            else
            {
                if (has0)
                {
                    str += "零";
                    has0 = false;
                }
            }
            str += NTD[int.Parse(number[i].ToString())];
            str += DW[w];
        }
        if (left)
            str += lastdw;
        return str;
    }
    //20140106 Add by GoodTV Tanya:增加中文字驗證
    public static bool isChinese(string strChinese)
    {
        bool bresult = true;
        int dRange = 0;
        int dstringmax = Convert.ToInt32("9fff", 16);
        int dstringmin = Convert.ToInt32("4e00", 16);
        for (int i = 0; i < strChinese.Length; i++)
        {
            dRange = Convert.ToInt32(Convert.ToChar(strChinese.Substring(i, 1)));
            if (dRange >= dstringmin && dRange < dstringmax)
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