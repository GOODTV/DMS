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

public partial class Report_Custom_Report_Donate_Month_Report2_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Year_month2.Value = Util.GetQueryString("Year");
            HFD_Month_month2.Value = Util.GetQueryString("Month");
            HFD_Donate_Payment_month2.Value = Util.GetQueryString("Donate_Payment");
            HFD_Is_Abroad_month2.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_month2.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_month2.Value = Util.GetQueryString("Sex");
            Print_Excel();
        }
    }
    //---------------------------------------------------------------------------
    private void Print_Excel()
    {
        string strSql = Sql();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 1)
        {
            ShowSysMsg("查無資料!!!");
            Response.Write("<script language=javascript>window.close();</script>");
            return;
        }

        GridList.Text = GetTitle(dt.Rows[0]) + GetTable1(dt.Rows[0], strSql);
        //GetTableFooter();
        Util.OutputTxt(GridList.Text, "1", "donate_month_report2");
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
        string strTitle = "定期奉獻授權到期明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>終止年月：" + HFD_Year_month2.Value + "年" + HFD_Month_month2.Value + "月" + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人編號排序</span>";
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 14;
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

        DataTable dtRet = CaseUtil.Donate_Month_Report2_Print(dt);

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
                //cell.Width = "40mm";
                cell.Height = "40mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-left", ".5pt solid windowtext");
            }
            else if (iCtrl == 1)
            {
                //cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 2)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 3)
            {
                //ell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 4)
            {
                //cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 5)
            {
                //cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 6)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 7)
            {
                //cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 8)
            {
                //cell.Width = "50mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 9)
            {
                //cell.Width = "70mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 10)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 11)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 12)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 13)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-right", ".5pt solid windowtext");
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
                else if (Col == 13)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                    cell.Style.Add("text-align", "left");
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
        string strSql = @"SELECT Case When C.Donate_FromDate > A.Donate_FromDate AND C.Donate_ToDate > A.Donate_ToDate  
									  Then 'V' Else '' END '新增'  
								 ,Case When C.Donate_FromDate > A.Donate_FromDate AND C.Donate_ToDate > A.Donate_ToDate  
									  Then C.Pledge_Id Else NULL END '新授權書號'  
								 ,(CONVERT(varchar(12) , A.Donate_ToDate, 111 )) AS '授權迄日'  
								 ,Case When C.Donate_FromDate > A.Donate_FromDate AND C.Donate_ToDate > A.Donate_ToDate  
									  Then REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,C.Donate_Amt),1),'.00','') Else NULL END '新授權額' ,A.Donate_Payment AS '授權方式'  
								 ,A.Card_Bank AS '銀行名稱' ,A.Card_Type AS '卡別'  
								 ,A.Pledge_Id AS '授權書號' ,A.Donor_Id AS '編號' ,D.Donor_Name AS '天使姓名' ,ISNULL(D.Title,'') AS '稱謂'  
								 ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
											 (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
											 Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
									   Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
									   Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話'  
								 ,IsNull(Cellular_Phone,'') AS '手機',IsNull(D.ZipCode,'') AS '郵遞區號'  
								 ,Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'')  
									   else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
					   	 FROM dbo.PLEDGE A  
						 LEFT JOIN (SELECT Donor_Id,Donate_Payment,MAX(Pledge_Id) AS Pledge_Id  
						 			FROM dbo.PLEDGE WHERE Donate_Type='長期捐款'  
									GROUP BY Donor_Id,Donate_Payment) B  
						 ON A.Donor_Id = B.Donor_Id AND A.Donate_Payment = B.Donate_Payment  
						 LEFT JOIN (SELECT * FROM dbo.PLEDGE WHERE Donate_Type='長期捐款') C  
						 ON B.Donor_Id=C.Donor_Id AND B.Donate_Payment = C.Donate_Payment AND B.Pledge_Id = C.Pledge_Id  
						 LEFT JOIN dbo.DONOR D  ON A.Donor_Id = D.Donor_Id  
                         LEFT JOIN dbo.CODECITYNew C1 ON D.City = C1.ZipCode
                         LEFT JOIN dbo.CODECITYNew C2 ON D.Area = C2.ZipCode
						 WHERE A.Donate_Type='長期捐款' ";

        //搜尋條件
        string[] Donate_Payment;
        int count = 0;
        if (HFD_Donate_Payment_month2.Value != "")
        {
            char[] ch = new char[] { ',' };
            Donate_Payment = HFD_Donate_Payment_month2.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
            count = Donate_Payment.Length;

            bool first = true; int cnt = 0;
            for (int i = 0; i < count; i++)
            {
                if (first)
                {
                    strSql += " and (A.Donate_Payment='" + Donate_Payment[i] + "' "; first = false; cnt++;
                }
                else
                {
                    strSql += " or A.Donate_Payment='" + Donate_Payment[i] + "' "; cnt++;
                }
            }
            if (cnt != 0) strSql += ')';
        }
        else
        {
            strSql += " And A.Donate_Payment <> '' ";
        }
        DateTime Donate_Date_Begin, Donate_Date_End;
        if (HFD_Year_month2.Value != "" && HFD_Month_month2.Value != "")
        {
            Donate_Date_Begin = Convert.ToDateTime(HFD_Year_month2.Value + "/" + HFD_Month_month2.Value + "/1");
            Donate_Date_End = Donate_Date_Begin.AddMonths(1).AddDays(-1);
            strSql += " And A.Donate_ToDate >= '" + Donate_Date_Begin.ToString("yyyy/MM/dd") + "' And A.Donate_ToDate <= '" + Donate_Date_End.ToString("yyyy/MM/dd") + "' ";
        }
        if (HFD_Is_Abroad_month2.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_month2.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_month2.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }

        //ORDER BY D.Donate_Purpose,D.Donate_Amt,D.Donate_Date,Last_DonateDate ";
        //以下是串UNION的SQL語法
        strSql += @" Union 
					 select '總筆數：',count(A.Pledge_Id),'','','筆','','','','999999','','','','','',''
					 FROM dbo.PLEDGE A  
									LEFT JOIN (SELECT Donor_Id,Donate_Payment,MAX(Pledge_Id) AS Pledge_Id  
												FROM dbo.PLEDGE WHERE Donate_Type='長期捐款'  
												GROUP BY Donor_Id,Donate_Payment) B  
									ON A.Donor_Id = B.Donor_Id AND A.Donate_Payment = B.Donate_Payment  
									LEFT JOIN (SELECT * FROM dbo.PLEDGE WHERE Donate_Type='長期捐款') C  
									ON B.Donor_Id=C.Donor_Id AND B.Donate_Payment = C.Donate_Payment AND B.Pledge_Id = C.Pledge_Id  
									LEFT JOIN dbo.DONOR D  
									ON A.Donor_Id = D.Donor_Id  
					WHERE A.Donate_Type='長期捐款' ";
        //搜尋條件
        if (HFD_Donate_Payment_month2.Value != "")
        {
            char[] ch = new char[] { ',' };
            Donate_Payment = HFD_Donate_Payment_month2.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
            count = Donate_Payment.Length;

            bool first = true; int cnt = 0;
            for (int i = 0; i < count; i++)
            {
                if (first)
                {
                    strSql += " and (A.Donate_Payment='" + Donate_Payment[i] + "' "; first = false; cnt++;
                }
                else
                {
                    strSql += " or A.Donate_Payment='" + Donate_Payment[i] + "' "; cnt++;
                }
            }
            if (cnt != 0) strSql += ')';
        }
        else
        {
            strSql += " And A.Donate_Payment <> '' ";
        }
        if (HFD_Year_month2.Value != "" && HFD_Month_month2.Value != "")
        {
            Donate_Date_Begin = Convert.ToDateTime(HFD_Year_month2.Value + "/" + HFD_Month_month2.Value + "/1");
            Donate_Date_End = Donate_Date_Begin.AddMonths(1).AddDays(-1);
            strSql += " And A.Donate_ToDate >= '" + Donate_Date_Begin.ToString("yyyy/MM/dd") + "' And A.Donate_ToDate <= '" + Donate_Date_End.ToString("yyyy/MM/dd") + "' ";
        }
        if (HFD_Is_Abroad_month2.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_month2.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_month2.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        strSql += " Order by A.Donor_Id ";
        return strSql;
    }
}