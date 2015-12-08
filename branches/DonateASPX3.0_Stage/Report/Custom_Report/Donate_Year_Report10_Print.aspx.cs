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

public partial class Report_Custom_Report_Donate_Year_Report10_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Year_year10.Value = Util.GetQueryString("Year");
            HFD_DonorId_year10.Value = Util.GetQueryString("DonorId");
            HFD_Is_Abroad_year10.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_year10.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_year10.Value = Util.GetQueryString("Sex");
            PrintReport();
        }
    }
    //---------------------------------------------------------------------------
    private void PrintReport()
    {
        string strSql = Sql();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
            return;
        }

        DataTable dtRet = CaseUtil.Donate_Year_Report10_Print(dt);

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
        /////格式很怪，若需要顯示的話需要再修改
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
                //cell.Width = "70mm";
                cell.Height = "40mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-left", ".5pt solid windowtext");
            }
            else if (iCtrl == 1)
            {
                //cell.Width = "50mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 2)
            {
                //cell.Width = "80mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 3)
            {
                //cell.Width = "60mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 4)
            {
                //cell.Width = "60mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 5)
            {
                //cell.Width = "30mm";
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
                //cell.Width = "85mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 9)
            {
                //cell.Width = "85mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 10)
            {
                //cell.Width = "40mm";
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
                else if (Col == 10)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                    cell.Style.Add("text-align", "center");
                }
                else if (Col == 7)
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

        GridList.Text = HFD_Year_year10.Value + "年 年度收據列印明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + DateTime.Now.ToString() + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人編號排序</span><br>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql = @"  SELECT M.Donor_Id AS '捐款人編號' ,M.Donor_Name AS '捐款人' ,Invoice_Title  AS '收據抬頭'
		                            ,ISNULL(M.Title,'') AS '稱謂',ISNULL(M.Sex,'') AS '性別', ISNULL(IDNo,'') as '身份證/統編' 
		                            ,D.Invoice_Type AS '收據開立'  
                                    ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Sum_Amt),1),'.00','') AS '累計奉獻金額', D.Times AS '奉獻次數'   
	                                ,Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') 
	                                                Else IsNull(M.ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '通訊地址'
                                    ,Case When IsAbroad = 'Y' Then IsNull(Invoice_OverseasCountry,'') + ' ' +IsNull(Invoice_OverseasAddress,'')
	                                                Else IsNull(M.Invoice_ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Invoice_Address END ' 收據地址'
	    
                            FROM DONOR M  
                            inner JOIN (select  Donor_Id ,Invoice_Type, COUNT(Donor_Id) AS Times , SUM(Donate_Amt) AS Sum_Amt  
                                        from dbo.DONATE 
                                        where ISNULL(Issue_Type,'') != 'D'  ";

        //搜尋條件
        if (HFD_Year_year10.Value != "")
        {
            strSql += " and DATEPART(YEAR,Donate_Date) = '" + HFD_Year_year10.Value + "'";
        }
        strSql += @" group by Donor_Id,Invoice_Type) D
                     ON M.Donor_Id = D.Donor_Id 
                        LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
                        LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode
                     Where 1=1 ";
        if (HFD_DonorId_year10.Value != "")
        {
            strSql += " And M.Donor_Id='" + HFD_DonorId_year10.Value + "' ";
        }
        if (HFD_Is_Abroad_year10.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        else
        {
            strSql += " And IsAbroad='Y' ";
        }
        if (HFD_Is_ErrAddress_year10.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_year10.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        //以下是串UNION的SQL語法
        /*strSql += @" Union select count(D.Donor_Id),REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(D.Sum_Amt)),1),'.00',''),'','','','','','0','','',sum(D.Sum_Amt)
                                 FROM dbo.DONOR M  
                                 INNER JOIN (select  Donor_Id ,Invoice_Type, COUNT(Donor_Id) AS Times , SUM(Donate_Amt) AS Sum_Amt
                                             FROM DONATE 
                                             WHERE ISNULL(Issue_Type,'') != 'D' ";
        //搜尋條件
        if (HFD_Year_year10.Value != "")
        {
            strSql += " and DATEPART(YEAR,Donate_Date) = '" + HFD_Year_year10.Value + "' group by Donor_Id,Invoice_Type) D ";
        }
        strSql += " ON M.Donor_Id = D.Donor_Id  Where 1=1 ";
        if (HFD_DonorId_year10.Value != "")
        {
            strSql += " And M.Donor_Id='" + HFD_DonorId_year10.Value + "' ";
        }
        if (HFD_Is_Abroad_year10.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_year10.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_year10.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        //strSql += " Order by D.Sum_Amt ";*/
        return strSql;
    }
}