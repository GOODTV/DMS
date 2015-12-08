﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;
using System.Web.UI.HtmlControls;

public partial class Report_Custom_Report_Donate_Month_Report4_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Donate_Amt_month4.Value = Util.GetQueryString("Donate_Amt");
            HFD_Donate_Total_Amt_month4.Value = Util.GetQueryString("Donate_Total_Amt");
            HFD_Donate_Purpose_month4.Value = Util.GetQueryString("Donate_Purpose");
            HFD_Is_Abroad_month4.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_month4.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_month4.Value = Util.GetQueryString("Sex");
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
        Util.OutputTxt(GridList.Text, "1", "donate_month_report4");
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
        string strTitle = "捐款人EMAIL明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + DateTime.Now.ToString() + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人編號排序</span>";
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 13;
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

        DataTable dtRet = CaseUtil.Donate_Month_Report4_Print(dt,true);

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
                cell.Width = "55mm";
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
                cell.Width = "50mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 3)
            {
                //cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 4)
            {
                //cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 5)
            {
                //cell.Width = "25mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 6)
            {
                //cell.Width = "25mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 7)
            {
                //cell.Width = "100mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 8)
            {
                cell.Width = "100mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 9)
            {
                //cell.Width = "60mm";
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
                else if (Col == 12)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                    cell.Style.Add("text-align", "center");
                }
                else if (Col == 2)
                {
                    cell.Width = "50mm";
                    cell.Style.Add("text-align", "center");
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
        string strSql = @" SELECT D.Donor_Id AS '編號'   
								 ,M.Donor_Name AS '天使姓名' ,IsNull(M.Title,'') AS '稱謂' ,IsNull(Donor_Type,'') AS '身份別'  
								 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Amt),1),'.00','') AS '累計奉獻金額' ,D.Times AS '奉獻次數'  
								 ,IsNull(M.ZipCode,'') AS '郵遞區號'  
								 ,Case When M.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'')  
									Else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
								 ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
									(Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
									Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
									Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
									Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話'  
								 ,IsNull(Cellular_Phone,'') AS '手機'  
								 ,Email AS 'Email' ,CONVERT(VARCHAR, Begin_DonateDate, 111 ) AS '首捐日期' ,CONVERT(VARCHAR, Last_DonateDate, 111 ) AS '末捐日期'  
						   FROM dbo.DONOR M  
						   RIGHT JOIN (SELECT Donor_Id ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times  
						               FROM DONATE  
                                       Where ISNULL(Issue_Type,'') != 'D' ";

        //搜尋條件
        string[] Donate_Purpose;
        int count = 0;
        if (HFD_Donate_Purpose_month4.Value != "")
        {
            char[] ch = new char[] { ',' };
            Donate_Purpose = HFD_Donate_Purpose_month4.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
            count = Donate_Purpose.Length;

            bool first = true; int cnt = 0;
            for (int i = 0; i < count; i++)
            {
                if (first)
                {
                    strSql += " and (Donate_Purpose='" + Donate_Purpose[i] + "' "; first = false; cnt++;
                }
                else
                {
                    strSql += " or Donate_Purpose='" + Donate_Purpose[i] + "' "; cnt++;
                }
            }
            if (cnt != 0) strSql += ')';
        }
        if (HFD_Donate_Amt_month4.Value != "")
        {
            strSql += " AND Donate_Amt >'" + HFD_Donate_Amt_month4.Value + "' ";
        }
        else
        {
            strSql += " AND Donate_Amt >0 ";
        }
        if (HFD_Donate_Total_Amt_month4.Value != "")
        {
            strSql += " GROUP BY Donor_Id HAVING SUM(Donate_Amt) >= ' " + HFD_Donate_Total_Amt_month4.Value + "' )D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id ) D ";
        }
        strSql += @" ON M.Donor_Id = D.Donor_Id
                     LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
                     LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode
                     WHERE Email <> ''";
        if (HFD_Is_Abroad_month4.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_month4.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_month4.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        //以下是串UNION的SQL語法
        strSql += @" Union 
					 select count(D.Donor_Id),'總筆數：','筆','','0','','總計','捐款金額：',REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(D.Amt)),1),'.00',''),'元','','',''
                          FROM dbo.DONOR M  
						       RIGHT JOIN (SELECT Donor_Id ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times  
					                       FROM DONATE Where ISNULL(Issue_Type,'') != 'D' ";
        //搜尋條件
        if (HFD_Donate_Purpose_month4.Value != "")
        {
            char[] ch = new char[] { ',' };
            Donate_Purpose = HFD_Donate_Purpose_month4.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
            count = Donate_Purpose.Length;

            bool first = true; int cnt = 0;
            for (int i = 0; i < count; i++)
            {
                if (first)
                {
                    strSql += " and (Donate_Purpose='" + Donate_Purpose[i] + "' "; first = false; cnt++;
                }
                else
                {
                    strSql += " or Donate_Purpose='" + Donate_Purpose[i] + "' "; cnt++;
                }
            }
            if (cnt != 0) strSql += ')';
        }
        if (HFD_Donate_Amt_month4.Value != "")
        {
            strSql += " AND Donate_Amt >'" + HFD_Donate_Amt_month4.Value + "' ";
        }
        else
        {
            strSql += " AND Donate_Amt >0 ";
        }
        if (HFD_Donate_Total_Amt_month4.Value != "")
        {
            strSql += " GROUP BY Donor_Id HAVING SUM(Donate_Amt) >= ' " + HFD_Donate_Total_Amt_month4.Value + "' )D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id ) D ";
        }
        strSql += " ON M.Donor_Id = D.Donor_Id WHERE Email <> ''";
        if (HFD_Is_Abroad_month4.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_month4.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_month4.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        strSql += " ORDER BY '累計奉獻金額' Desc ";
        return strSql;
    }
}