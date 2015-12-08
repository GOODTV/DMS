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

public partial class DonateMgr_Donate_InvoicePrn : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            string Donate_Id = Util.GetQueryString("Donate_Id");
            if (Donate_Id != "")
            {
                Donate_Edit(Donate_Id);
                Print(Donate_Id);
            }
            else
            {
                Response.Write("empty string");
            }
        }
    }
    public void Print(string Donate_Id)
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = @"select dr.Donor_Id as [編號],
                                 dr.Donor_Name as [捐款人],dr.Title as [稱謂],dr.Title2 as [稱謂2],do.Invoice_No as [收據編號],Donate_Forign as [外幣],Donate_ForignAmt as [外幣金額],do.Donate_Date,CONVERT(nvarchar(4000),Invoice_PrintComment) as Invoice_PrintComment,
                                 do.Donate_Purpose as [款項用途],dr.Cellular_Phone as [電話], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,do.Donate_Amt),1),'.00','') as [金額], dr.Invoice_ZipCode as [郵遞區號],
                                 (Case When dr.Invoice_City='' Then Invoice_Address Else Case When A.mValue<>B.mValue Then A.mValue+B.mValue+Invoice_Address Else A.mValue+Invoice_Address End End) as [地址],
                                 dr.IsAbroad_Invoice,dr.Invoice_Attn,dr.Attn,dr.Invoice_OverseasCountry,do.Invoice_Title,
                                 IsNull(dr.Invoice_ZipCode,'') as [Invoice_ZipCode],(Case When dr.Invoice_City='' Then dr.Invoice_Address Else Case When A.mValue<>B.mValue Then A.mValue+B.mValue+dr.Invoice_Address Else A.mValue+dr.Invoice_Address End End) as [Invoice_Address],
                                 IsNull(dr.ZipCode,'') as [ZipCode],(Case When dr.City='' Then dr.Address Else Case When C.mValue<>D.mValue Then C.mValue+D.mValue+dr.Address Else C.mValue+dr.Address End End) as [Address]
                          from Donate do left join Donor dr on do.Donor_Id = dr.Donor_Id     
                              Left Join CODECITY As A On dr.Invoice_City=A.mCode Left Join CODECITY As B On dr.Invoice_Area=B.mCode 
                              Left Join CODECITY As C On dr.City=C.mCode Left Join CODECITY As D On dr.Area=D.mCode
                          where Donate_Id = @Donate_Id";
        dict.Add("Donate_Id", Donate_Id);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        DataRow dr = dt.Rows[0];

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

        GridList.Text = GetTable(dr);

        Util.OutputTxt(s1 + GridList.Text + "</div></body>", "2", DateTime.Now.ToString("yyyyMMddHHmmss"));
    }
    public string GetTable(DataRow dr)
    {
        string strRet = "";

        //組 table
        string strTemp = "";
        //data

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
        css.Add("font-size", strFontSize);
        css.Add("width", "100%");
        css.Add("margin-top", "11mm");
        css.Add("margin-left", "21mm");
        css.Add("height", "6mm");
        row.Cells.Add(cell);
        table.Rows.Add(row);

        row = new HtmlTableRow();
        cell = new HtmlTableCell();
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
        //strTemp = dr["郵遞區號"].ToString();
        //strTemp += dr["地址"].ToString();
        cell.InnerHtml = strTemp == "" ? "&nbsp;" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", "16px");
        css.Add("width", "100%");
        css.Add("margin-top", "12mm");
        css.Add("margin-left", "95mm");
        css.Add("margin-right", "30mm");
        css.Add("height", "28mm");
        row.Cells.Add(cell);
        table.Rows.Add(row);

        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        //strTemp = "<br>" + dr["捐款人"].ToString() + " &nbsp " + dr["稱謂"].ToString() + "收";
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
        css.Add("font-size", "16px");
        css.Add("width", "100%");
        css.Add("margin-left", "95mm");
        css.Add("margin-right", "35mm");
        css.Add("height", "10mm"); 
        row.Cells.Add(cell);
        table.Rows.Add(row);

        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        strTemp = dr["編號"].ToString();
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", "16px");
        css.Add("width", "100%");
        css.Add("margin-top", "10mm");
        css.Add("margin-left", "95mm");
        css.Add("height", "8mm");
        row.Cells.Add(cell);
        table.Rows.Add(row);

        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        strTemp = dr["收據編號"].ToString();
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", "20px");
        css.Add("width", "100%");
        css.Add("line-height", "120%");
        css.Add("margin-top", "42mm");
        css.Add("margin-left", "89mm");
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
        //國字大寫
        //string Donate_Amt = GetMoneyUpper(Decimal.Parse(dr["金額"].ToString().Replace(",", "") + ".00"));
        string strAmt = Convert.ToDouble(dr["金額"].ToString()).ToString("C0");
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

        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        //strTemp = Donor_Name;
        //20140821改為 收據抬頭
        strTemp = dr["Invoice_Title"].ToString();
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
        /*
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
        }*/
        if (dr["Invoice_ZipCode"].ToString() != "")
        {
            strTemp = dr["Invoice_ZipCode"].ToString();
        }
        else
        {
            strTemp = dr["郵遞區號"].ToString();
        }
        if (dr["Invoice_Address"].ToString() != "")
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
        //--------------------------------------------

        ////轉成 html 碼
        //StringWriter sw = new StringWriter();

        //HtmlTextWriter htw = new HtmlTextWriter(sw);
        //table.RenderControl(htw);

        //return htw.InnerWriter.ToString();
        return strRet;
    } 
    public void Donate_Edit(string Donate_Id)
    {
        //先確定是否有首印過
        string strSql = " select Invoice_Print_Date from  Donate where Donate_Id ='" + Donate_Id + "'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (dr["Invoice_Print_Date"].ToString() == "")
        //20150325增加 更新 首印經手人和最後補印經手人
        //無首印
        {
            //****設定SQL指令****//
            strSql = " update Donate set ";

            strSql += " Invoice_Print_Date = @Invoice_Print_Date";
            strSql += " ,Invoice_FirstPrint_User = @Invoice_FirstPrint_User";
            strSql += " ,Invoice_Print = @Invoice_Print";
            strSql += " where Donate_Id = @Donate_Id";

            dict.Add("Invoice_Print_Date", DateTime.Now.ToString("yyyy-MM-dd"));
            dict.Add("Invoice_FirstPrint_User", SessionInfo.UserName);
            dict.Add("Invoice_Print", "1");
            dict.Add("Donate_Id", Donate_Id);
        }
        else
        //有首印
        { 
            //****設定SQL指令****//
            strSql = " update Donate set ";

            strSql += " Invoice_RePrint_Date = @Invoice_RePrint_Date";
            strSql += " ,Invoice_LastPrint_User = @Invoice_LastPrint_User";
            strSql += " ,Invoice_Print = @Invoice_Print";
            strSql += " where Donate_Id = @Donate_Id";

            dict.Add("Invoice_RePrint_Date", DateTime.Now.ToString("yyyy-MM-dd"));
            dict.Add("Invoice_LastPrint_User", SessionInfo.UserName);
            dict.Add("Invoice_Print", "1");
            dict.Add("Donate_Id", Donate_Id);
        }
        NpoDB.ExecuteSQLS(strSql, dict);
    }
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
    public static string GetMoney(string number, bool left, string lastdw)
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