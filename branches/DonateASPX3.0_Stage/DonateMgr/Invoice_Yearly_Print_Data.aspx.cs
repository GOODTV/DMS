using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonateMgr_Invoice_Yearly_Print_Address : BasePage
{
    // 2014/4/21 取消未用變數
    //string TableWidth = "800px";
    // 2014/7/10 超過20筆捐款紀錄的跳頁提示訊息
    string strOverRecord = "";
    string strDonateIdGroup = "0";

    protected void Page_Load(object sender, EventArgs e)
    {
        Print();
        if (GridList.Text != "** 沒有符合條件的資料 **")
        {
            Response.Write("<script language=javascript>window.print();</script>");
        }
    }
    private void Print()
    {
        GridList.Text ="";
        string strSql = Session["strSql_Data"].ToString();
        string Condition = Session["Condition"].ToString();
        string Donate_Date_Year = Session["Donate_Date_Year"].ToString();
        string Invoice_Title_New = Session["Invoice_Title_New"].ToString();
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr;

        if (dt.Rows.Count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
            //ShowSysMsg("年度證明列印：共" + dt.Rows.Count + "筆");

            int count_Id = dt.Rows.Count;
            for (int i = 0; i < count_Id; i++)
            {
                dr = dt.Rows[i];
                // 2014/4/21 修改與ASP2.0相同
                //20140430 Add by GoodTV Tanya:增加單一捐款人時可「收據抬頭改印」
                //GridList.Text += GetTitle() + GetName(dr) + GetData(dr["捐款人編號"].ToString(), Donate_Date_Year) + GetEnd(dr["捐款人編號"].ToString(), Donate_Date_Year);
                //  2014/7/11 增加跨頁(超過20筆跳頁)
                if ((int)dr["row_cnt"] > 20)
                {
                    strOverRecord += strOverRecord + "No." + dr["捐款人編號"].ToString() + " " + dr["捐款人"].ToString() + "\\n";
                }
                int detail_cnt = (int)dr["row_cnt"] / 20;
                for (int j = 0; j <= detail_cnt; j++)
                {
                    if (count_Id == 1 && !string.IsNullOrEmpty(Invoice_Title_New))
                        GridList.Text += GetHead() + GetAddress(dr) + GetName(dr, Invoice_Title_New) + GetData(dr["捐款人編號"].ToString(), Donate_Date_Year, j) + (j == detail_cnt ? GetEnd(dr["捐款人編號"].ToString(), Donate_Date_Year) : "");
                    else
                        GridList.Text += GetHead() + GetAddress(dr) + GetName(dr, "") + GetData(dr["捐款人編號"].ToString(), Donate_Date_Year, j) + (j == detail_cnt ? GetEnd(dr["捐款人編號"].ToString(), Donate_Date_Year) : "");

                    if (j != detail_cnt)
                    {
                        GridList.Text += "<div style='page-break-after:always;'>&nbsp;</div>";  //列印強制換頁,最後一頁不用換頁

                    }
                }

                if (i != count_Id - 1)
                {
                    GridList.Text += "<div style='page-break-after:always;'>&nbsp;</div>";  //列印強制換頁,最後一頁不用換頁

                }

            }

            // 補漏產出年度證明後註記已列印
            UpdateDonatePrintFlag();

            ShowSysMsg((strOverRecord.Length > 0 ? strOverRecord + "跨頁\\n\\n" : "") + "年度證明列印：共" + count_Id + "筆");

        }
    }

    private void UpdateDonatePrintFlag()
    {
        //20150521 增加年度證明列印時間與列印經手人
        string Reprint = Session["RePrint"].ToString();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (Reprint == "Y") //補印
        {
            // 產出年度證明後註記已列印
            string strSql_Donate = " update Donate set ";
            strSql_Donate += " Invoice_Print_Yearly = @Invoice_Print_Yearly";
            strSql_Donate += " ,Invoice_Yearly_RePrint_Date = @Invoice_Yearly_RePrint_Date";
            strSql_Donate += " ,Invoice_Yearly_LastPrint_User = @Invoice_Yearly_LastPrint_User";
            strSql_Donate += " where Donate_Id in (" + strDonateIdGroup + ");";
            dict.Add("Invoice_Print_Yearly", "1");
            dict.Add("Invoice_Yearly_RePrint_Date", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
            dict.Add("Invoice_Yearly_LastPrint_User", SessionInfo.UserName);
            NpoDB.ExecuteSQLS(strSql_Donate, dict);
        }
        else //首印
        {
            // 產出年度證明後註記已列印
            string strSql_Donate = " update Donate set ";
            strSql_Donate += " Invoice_Print_Yearly = @Invoice_Print_Yearly";
            strSql_Donate += " ,Invoice_Yearly_Print_Date = @Invoice_Yearly_Print_Date";
            strSql_Donate += " ,Invoice_Yearly_FirstPrint_User = @Invoice_Yearly_FirstPrint_User";
            strSql_Donate += " where Donate_Id in (" + strDonateIdGroup + ");";
            dict.Add("Invoice_Print_Yearly", "1");
            dict.Add("Invoice_Yearly_Print_Date", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
            dict.Add("Invoice_Yearly_FirstPrint_User", SessionInfo.UserName);
            NpoDB.ExecuteSQLS(strSql_Donate, dict);
        }
        

    }

    private string GetHead()
    {
        //組 table
        string strTemp = "&nbsp";
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        table.Align = "center";
        css = table.Style;
        css.Add("width", "194mm");
        //css.Add("border-collapse", "collpase");
        //css.Add("margin-top", "5mm");
        //css.Add("margin-bottom", "8mm");
        //css.Add("margin-left", "5mm");
        //css.Add("margin-right", "5mm");
        //--------------------------------------------
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        cell.InnerHtml = strTemp;

        css = cell.Style;
        css.Add("height", "7mm");
        row.Cells.Add(cell);
        table.Rows.Add(row);
        //--------------------------------------------

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }

    private string GetAddress(DataRow dr)
    {
        string addr = "";
        //組 table
        string strTemp = "";
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        table.Align = "center";
        css = table.Style;
        //css.Add("font-size", "18px");
        css.Add("font-family", "標楷體");
        css.Add("width", "194mm");
        //css.Add("border-collapse", "collpase");
        //css.Add("margin-top", "5mm");
        //css.Add("margin-bottom", "8mm");
        //css.Add("margin-left", "5mm");
        //css.Add("margin-right", "5mm");

        string strFontSize = "18px";
        //--------------------------------------------
        row = new HtmlTableRow();
        // 第一格
        cell = new HtmlTableCell();
        cell.VAlign = "top";
        cell.Align = "center";
        strTemp = "&nbsp";
        cell.InnerHtml = strTemp;

        css = cell.Style;
        css.Add("width", "17mm");
        css.Add("height", "82mm");
        //css.Add("text-valign", "top");
        //css.Add("text-align", "center");
        row.Cells.Add(cell);

        // 第二格
        cell = new HtmlTableCell();
        cell.VAlign = "top";
        cell.Align = "left";
        strTemp = dr["Invoice_Title"].ToString();
        strTemp = strTemp == "" ? "&nbsp" : strTemp;
        cell.InnerHtml = strTemp += "";

        css = cell.Style;
        css.Add("font-size", strFontSize);
        css.Add("width", "75mm");
        css.Add("height", "85mm");
        //css.Add("valign", "top");
        //css.Add("align", "left");
        row.Cells.Add(cell);

        // 第三格
        //地址加Attn 20140709 by 詩儀
        if (dr["Attn"].ToString() != "")
        {
            addr = dr["郵遞區號"].ToString() + dr["City"].ToString() + dr["Area"].ToString() + dr["Address"].ToString() + "(" + dr["Attn"].ToString() + ")";
        }
        else 
        {
            addr = dr["郵遞區號"].ToString() + dr["City"].ToString() + dr["Area"].ToString() + dr["Address"].ToString();
        }
        cell = new HtmlTableCell();
        cell.VAlign = "top";
        strTemp = "<br><br><br><br><br>";
        if (dr["IsAbroad"].ToString() == "Y") //國外收據列印格式調整
        {
            strTemp += dr["捐款人"].ToString() + "　" + dr["稱謂"].ToString() + "<br>";
            if (addr.Length > 22)
            {
                //Response.Write(addr.Substring(0, 22).ToString() + "<br>" + addr.Substring(22));
                //Response.End();
                strTemp += addr.Substring(0, 22) + "<br>";
                strTemp += addr.Substring(22) + "<br>";
            }
            else
            {
                strTemp += addr + "<br><br>";
            }
        }
        else
        {
            if (addr.Length > 22)
            {
                //Response.Write(addr.Substring(0, 22).ToString() + "<br>" + addr.Substring(22));
                //Response.End();
                strTemp += addr.Substring(0, 22) + "<br>";
                strTemp += addr.Substring(22) + "<br>";
            }
            else
            {
                strTemp += addr + "<br><br>";
            }
            strTemp += dr["捐款人"].ToString() + "　" + dr["稱謂"].ToString() + "　　鈞啟<br>";
        }
        strTemp += dr["捐款人編號"].ToString();
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("font-size", strFontSize);
        css.Add("width", "102mm");
        css.Add("height", "82mm");
        //css.Add("text-valign", "top");
        //css.Add("text-align", "left");
        row.Cells.Add(cell);
        table.Rows.Add(row);
        //--------------------------------------------

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }

    private string GetTitle()
    {
        //組 table
        string strTemp = "";
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 1;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "16px");
        css.Add("font-family", "標楷體");
        css.Add("width", "100%");
        css.Add("border-collapse", "collpase");
        css.Add("margin-top", "20px");
        css.Add("margin-bottom", "20px");
        css.Add("margin-left", "20px");
        css.Add("margin-right", "20px");

        string strFontSize = "18px";
        //--------------------------------------------
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        strTemp = "親愛的<br>";
        strTemp += "奉獻天使平安:<br><br>";
        strTemp += "感謝 神使我們一起同工，<br>";
        strTemp += "將福音廣傳到地極。<br>";
        strTemp += "願主紀念您所擺上的一切，<br>";
        strTemp += "敞開天上的窗戶傾福與您。<br><br>";
        strTemp += "奉獻徵信名錄，請上GOOD TV網站查詢。<br><br>";
        strTemp += "財團法人加百列福音傳播基金會 敬上";
        cell.InnerHtml += strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("text-align", "left");
        css.Add("font-size", strFontSize);
        css.Add("height", "100mm");
        //css.Add("valign", "top");
        cell.VAlign = "Top";
        row.Cells.Add(cell);
        table.Rows.Add(row);
        //--------------------------------------------

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    private string GetName(DataRow dr, string Invoice_Title_New)
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
        table.Align = "center";
        css = table.Style;
        //css.Add("font-size", "16px");
        css.Add("font-family", "標楷體");
        css.Add("width", "194mm");
        //css.Add("border-collapse", "collpase");
        //css.Add("margin-top", "5mm");
        //css.Add("margin-bottom", "8mm");
        //css.Add("margin-left", "5mm");
        //css.Add("margin-right", "5mm");

        string strFontSize = "18px";
        //--------------------------------------------
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = "&nbsp";

        css = cell.Style;
        css.Add("width", "22mm");
        css.Add("height", "30mm");
        row.Cells.Add(cell);

        cell = new HtmlTableCell();
        cell.VAlign = "bottom";
        //20140430 Add by GoodTV Tanya:增加單一捐款人時可「收據抬頭改印」
        if (!string.IsNullOrEmpty(Invoice_Title_New))
            strTemp = Invoice_Title_New;
        else
            strTemp = dr["Invoice_Title"].ToString();
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;

        css = cell.Style;
        css.Add("font-size", strFontSize);
        css.Add("width", "128mm");
        css.Add("height", "30mm");
        //css.Add("text-valign", "bottom");
        row.Cells.Add(cell);

        cell = new HtmlTableCell();
        cell.VAlign = "bottom";
        strTemp = dr["Invoice_IDNo"].ToString();
        cell.InnerHtml += strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("font-size", strFontSize);
        css.Add("width", "44mm");
        css.Add("height", "30mm");
        //css.Add("text-valign", "bottom");
        row.Cells.Add(cell);
        table.Rows.Add(row);
        //--------------------------------------------
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.ColSpan = 3;
        cell.InnerHtml = "&nbsp";
        css = cell.Style;
        css.Add("height", "13mm");
        row.Cells.Add(cell);
        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    private string GetData(string Donor_Id, string Donate_Date_Year,int for_loop_cnt)
    {

        //抓出捐款的詳細資料---------------------------------

        // 2014/7/11 增加跨頁的迴圈
        int for_cnt = for_loop_cnt;
        string strSql = @" select Convert(CHAR(8),Donate_Date,112) as [Donate_Date],Donate_Id,
               ISNULL(Invoice_Pre,'')+Invoice_No as [Invoice_No], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','') as [Donate_Amt]
               from Donate 
               where Donor_Id = " + Donor_Id + " and Year(Donate_Date) = '" + Donate_Date_Year + "' and Issue_Type !='D' order by Donate_Date ";
        strSql += @"OFFSET " + for_cnt * 20 + "  ROWS FETCH NEXT 20 ROWS ONLY;";

        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr;
        int count_Donate = dt.Rows.Count;

        //組 table
        string strTemp = "";
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        table.Align = "center";
        css = table.Style;
        //css.Add("font-size", "14px");
        css.Add("font-family", "標楷體");
        css.Add("width", "194mm");
        //css.Add("border-collapse", "collpase");
        //css.Add("margin-top", "5mm");
        //css.Add("margin-bottom", "8mm");
        //css.Add("margin-left", "5mm");
        //css.Add("margin-right", "5mm");

        string strFontSize = "14px";
        //--------------------------------------------
        for (int i = 0; i < count_Donate; i = i + 2)
        {

            row = new HtmlTableRow();
            cell = new HtmlTableCell();
            dr = dt.Rows[i];

            strTemp = "";
            cell.InnerHtml += strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("font-size", strFontSize);
            css.Add("width", "8mm");
            css.Add("height", "6mm");
            row.Cells.Add(cell);

            cell = new HtmlTableCell();
            cell.Align = "left";
            cell.VAlign = "center";
            strTemp = dr["Donate_Date"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("font-size", strFontSize);
            css.Add("width", "25.5mm");
            css.Add("height", "6mm");
            //css.Add("text-align", "left");
            //css.Add("text-valign", "center");
            row.Cells.Add(cell);

            strDonateIdGroup += ", " + dr["Donate_Id"].ToString();

            cell = new HtmlTableCell();
            cell.Align = "left";
            cell.VAlign = "center";
            strTemp = dr["Invoice_No"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("font-size", strFontSize);
            css.Add("width", "29mm");
            css.Add("height", "6mm");
            //css.Add("text-align", "left");
            //css.Add("text-valign", "center");
            row.Cells.Add(cell);

            cell = new HtmlTableCell();
            cell.Align = "right";
            cell.VAlign = "center";
            strTemp = dr["Donate_Amt"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("font-size", strFontSize);
            css.Add("width", "22.5mm");
            css.Add("height", "6mm");
            //css.Add("text-align", "right");
            //css.Add("text-valign", "center");

            row.Cells.Add(cell);
            //--------------------
            if (i == count_Donate - 1 && count_Donate % 2 == 1)
            {
                cell = new HtmlTableCell();
                cell.Align = "right";
                cell.VAlign = "center";
                strTemp = "";
                cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                css = cell.Style;
                css.Add("font-size", strFontSize);
                css.Add("width", "22.5mm");
                css.Add("height", "6mm");
                //css.Add("text-align", "right");
                //css.Add("text-valign", "center");
                row.Cells.Add(cell);

                cell = new HtmlTableCell();
                cell.Align = "left";
                cell.VAlign = "center";
                strTemp = "";
                cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                css = cell.Style;
                css.Add("font-size", strFontSize);
                css.Add("width", "27mm");
                css.Add("height", "6mm");
                //css.Add("text-align", "left");
                //css.Add("text-valign", "center");
                row.Cells.Add(cell);

                cell = new HtmlTableCell();
                cell.Align = "left";
                cell.VAlign = "center";
                strTemp = "";
                cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                css = cell.Style;
                css.Add("font-size", strFontSize);
                css.Add("width", "28.5mm");
                css.Add("height", "6mm");
                //css.Add("text-align", "left");
                //css.Add("text-valign", "center");
                row.Cells.Add(cell);

                cell = new HtmlTableCell();
                cell.Align = "right";
                cell.VAlign = "center";
                strTemp = "";
                cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                css = cell.Style;
                css.Add("font-size", strFontSize);
                css.Add("width", "27mm");
                css.Add("height", "6mm");
                //css.Add("text-align", "right");
                //css.Add("text-valign", "center");
                row.Cells.Add(cell);

                cell = new HtmlTableCell();
                cell.Align = "right";
                cell.VAlign = "center";
                strTemp = "";
                cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
                css = cell.Style;
                css.Add("width", "7mm");
                css.Add("height", "6mm");
                //css.Add("text-align", "right");
                //css.Add("text-valign", "center");
                row.Cells.Add(cell);

                table.Rows.Add(row);

                break;
            }

            dr = dt.Rows[i + 1];

            cell = new HtmlTableCell();
            cell.Align = "right";
            cell.VAlign = "center";
            strTemp = "";
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("width", "22.5mm");
            css.Add("height", "6mm");
            //css.Add("text-align", "right");
            //css.Add("text-valign", "center");
            row.Cells.Add(cell);

            cell = new HtmlTableCell();
            cell.Align = "left";
            cell.VAlign = "center";
            strTemp = dr["Donate_Date"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("font-size", strFontSize);
            css.Add("width", "27mm");
            css.Add("height", "6mm");
            //css.Add("text-align", "left");
            //css.Add("text-valign", "center");
            row.Cells.Add(cell);

            strDonateIdGroup += ", " + dr["Donate_Id"].ToString();

            cell = new HtmlTableCell();
            cell.Align = "left";
            cell.VAlign = "center";
            strTemp = dr["Invoice_No"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("font-size", strFontSize);
            css.Add("width", "29mm");
            css.Add("height", "6mm");
            //css.Add("text-align", "left");
            //css.Add("text-valign", "center");
            row.Cells.Add(cell);

            cell = new HtmlTableCell();
            cell.Align = "right";
            cell.VAlign = "center";
            strTemp = dr["Donate_Amt"].ToString();
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("font-size", strFontSize);
            css.Add("width", "22.5mm");
            css.Add("height", "6mm");
            //css.Add("text-align", "right");
            //css.Add("text-valign", "center");
            row.Cells.Add(cell);

            cell = new HtmlTableCell();
            cell.Align = "right";
            cell.VAlign = "center";
            strTemp = "";
            cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            css = cell.Style;
            css.Add("font-size", strFontSize);
            css.Add("width", "12mm");
            css.Add("height", "6mm");
            //css.Add("text-align", "right");
            //css.Add("text-valign", "center");
            row.Cells.Add(cell);

            table.Rows.Add(row);
        }
        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    private string GetEnd(string Donor_Id, string Donate_Date_Year)
    {
        //抓出捐款總金額-------------------------------------
        string strSql = @" select REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when Donate_Amt!='0' then Donate_Amt else '0' end)),1),'.00','') as Donate_Amt_Num 
                                    from Donate 
                                    where Donor_Id = '" + Donor_Id + "' and Year(Donate_Date) = '" + Donate_Date_Year + "' and Issue_Type !='D'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr;
        dr = dt.Rows[0];

        //組 table
        string strTemp = "";
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        table.Align = "center";
        css = table.Style;
        //css.Add("font-size", "16px");
        css.Add("font-family", "標楷體");
        css.Add("width", "194mm");
        //css.Add("border-collapse", "collpase");
        //css.Add("margin-top", "5mm");
        //css.Add("margin-bottom", "8mm");
        //css.Add("margin-left", "5mm");
        //css.Add("margin-right", "5mm");

        string strFontSize = "18px";
        //--------------------------------------------
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        cell.Align = "center";
        cell.VAlign = "center";
        string Donate_Amt = GetMoneyUpper(Decimal.Parse(dr["Donate_Amt_Num"].ToString().Replace(",", "") + ".00"));
        strTemp = Session["Donate_Date_Year"].ToString() + "年度捐款金額總計:新台幣" + Donate_Amt + "&nbsp;(NT$" + dr["Donate_Amt_Num"].ToString() + ")";
        cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        css = cell.Style;
        css.Add("font-size", strFontSize);
        css.Add("width", "194mm");
        css.Add("height", "10mm");
        //css.Add("text-align", "center");
        //css.Add("text-valign", "center");
        row.Cells.Add(cell);
        table.Rows.Add(row);
        //--------------------------------------------
        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
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
        string ret = GetMoney(dq, true, "元") + GetMoney(dh, false, "");
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
                return GetMoney(number.Substring(0, number.Length - 8), true, "億") + GetMoney(number.Substring(number.Length - 8, 8), true, "元");
            }
            if (number.Length >= 5)
            {
                return GetMoney(number.Substring(0, number.Length - 4), true, "萬") + GetMoney(number.Substring(number.Length - 4, 4), true, "元");
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
}
