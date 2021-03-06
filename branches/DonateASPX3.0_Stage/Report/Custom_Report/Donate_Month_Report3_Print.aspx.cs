﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class Report_Custom_Report_Donate_Month_Report3_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_IsMember_month3.Value = Util.GetQueryString("IsMember");
            HFD_Is_Abroad_month3.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_month3.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_month3.Value = Util.GetQueryString("Sex");
            PrintReport();
        }
    }
    //---------------------------------------------------------------------------
    private void PrintReport()
    {
        string strSql = Sql();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 1)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
            return;
        }

        DataTable dtRet = CaseUtil.Donate_Month_Report3_Print(dt);

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        table.Width = "770";
        css = table.Style;
        css.Add("font-size", "12px");
        css.Add("font-family", "細明體");
        css.Add("line-height", "20px");
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
                //cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 7)
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
        int count = 0;
        foreach (DataRow dr in dtRet.Rows)
        {
            int Col = 0;
            row = new HtmlTableRow();

            foreach (object objItem in dr.ItemArray)
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
                else if (Col == 7)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                    cell.Style.Add("text-align", "center");
                }
                else if (Col == 3)
                {
                    cell.Style.Add("text-align", "left");
                }
                else
                {
                    cell.Style.Add("text-align", "center");
                }
                row.Cells.Add(cell);
                Col++;

            }
            //印出大小不確定,先不強制分頁
            ////size 45 
            //count++;
            //if (count % 45 == 0)
            //{
            //    cell = new HtmlTableCell();
            //    string strTemp = "<br clear=all style='page-break-before:always'>";
            //    cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
            //    css = cell.Style;
            //    row.Cells.Add(cell);
            //}
            table.Rows.Add(row);
        }
        //轉成 html 碼
        StringWriter sw = new StringWriter();

        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        GridList.Text = "月刊郵寄明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + DateTime.Now.ToString() + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人編號排序</span><br>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql = @"SELECT MAIN.Donor_Id AS '編號' ,Donor_Name + ' ' + ISNULL(Title,'') AS '天使姓名'   
                                 ,MAIN.ZipCode AS '郵遞區號'  
							     ,Case When MAIN.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'')  
	                              else (Case when ISNULL(Attn,'') <> '' then IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address+'('+ISNULL(Attn,'')+')' 
				                        Else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address End) END '地址' ,Email AS 'Email' 
								 ,IsNull(Donor_Type,'') AS '身份別' ,IsNull(IsSendNewsNum,'') AS '月刊份數'  
								 ,Case when IsNull(IsContact,'') = 'N' Then '是' Else '' End '不主動聯絡'  
						  FROM DONOR MAIN   
                              LEFT JOIN dbo.CODECITYNew C1 ON MAIN.City = C1.ZipCode
                              LEFT JOIN dbo.CODECITYNew C2 ON MAIN.Area = C2.ZipCode 
						  WHERE IsSendNews='Y' and DeleteDate is NULL ";

        //搜尋條件
        if (HFD_IsMember_month3.Value == "Y")
        {
            strSql += " And MAIN.Donor_Type like '%讀者%' ";
        }
        else
        {
            strSql += " And MAIN.Donor_Type not like '%讀者%' ";
        }
        if (HFD_Is_Abroad_month3.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_month3.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_month3.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }

        //ORDER BY D.Donate_Purpose,D.Donate_Amt,D.Donate_Date,Last_DonateDate ";
        //以下是串UNION的SQL語法
        strSql += @" Union 
					 select '99999',convert(nvarchar, count(MAIN.Donor_Id)),'','','','',Sum (Cast (IsSendNewsNum as int)),''
					 FROM DONOR MAIN  
					 WHERE IsSendNews='Y' and DeleteDate is NULL ";
        //搜尋條件
        if (HFD_IsMember_month3.Value == "Y")
        {
            strSql += " And MAIN.Donor_Type like '%讀者%' ";
        }
        else
        {
            strSql += " And MAIN.Donor_Type not like '%讀者%' ";
        }
        if (HFD_Is_Abroad_month3.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_month3.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_month3.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        strSql += " Order by MAIN.Donor_Id ";
        return strSql;
    }
}