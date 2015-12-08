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

public partial class Report_Custom_Report_Donate_Season_Report2_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Donate_Amt_season2.Value = Util.GetQueryString("Donate_Amt");
            HFD_DonateDateS_season2.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_season2.Value = Util.GetQueryString("DonateDateE");
            HFD_Is_Abroad_season2.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_season2.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_season2.Value = Util.GetQueryString("Sex");
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

        DataTable dtRet = CaseUtil.Donate_Season_Report2_Print(dt);

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
            else if (iCtrl == 10)
            {
                //cell.Width = "70mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-right", ".5pt solid windowtext");
            }
            else
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
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
                    cell.Style.Add("text-align", "left");
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

        GridList.Text = "單筆捐款金額明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + HFD_DonateDateS_season2.Value + "~" + HFD_DonateDateE_season2.Value + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依累計奉獻金額排序</span><br>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql = @"  SELECT D.Donor_Id AS '編號'   
                                    ,M.Donor_Name AS '天使姓名' ,M.Title AS '稱謂'  
								    ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Max_Amt),1),'.00','') AS '最高奉獻金額'
								    ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Sum_Amt),1),'.00','') AS '累計奉獻金額' 
								    ,D.Times AS '奉獻次數',IsNull(M.ZipCode,'') AS '郵遞區號'  
								    ,Case When M.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'')  
											else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址',ISNULL(Email,'') AS 'Email'  
									,CONVERT(VARCHAR, Begin_DonateDate, 111) AS '首捐日期' ,CONVERT(VARCHAR, Last_DonateDate, 111) AS '末捐日期' ,D.Sum_Amt 
							FROM dbo.DONOR M  
							RIGHT JOIN (SELECT Donor_Id ,Max(Donate_Amt) AS Max_Amt, Sum(Donate_Amt) AS Sum_Amt ,Count(Donor_Id) AS Times  
							            FROM DONATE 
                                        WHERE ISNULL(Issue_Type,'') != 'D' ";

        //搜尋條件
        if (HFD_Donate_Amt_season2.Value != "")
        {
            strSql += " and Donate_Amt > '" + HFD_Donate_Amt_season2.Value + "' ";
        }
        else
        {
            strSql += " and Donate_Amt > 0 ";
        }
        if (HFD_DonateDateS_season2.Value != "")
        {
            strSql += " And Donate_Date>='" + HFD_DonateDateS_season2.Value + "' ";
        }
        if (HFD_DonateDateE_season2.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_season2.Value + "' ";
        }
        strSql += @" GROUP BY Donor_Id ) D 
                     ON M.Donor_Id = D.Donor_Id 
                     LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
                     LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode
                     WHERE 1=1 ";
        if (HFD_Is_Abroad_season2.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_season2.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_season2.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        //以下是串UNION的SQL語法
        strSql += @" Union select count(D.Donor_Id),'','',REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(D.Sum_Amt)),1),'.00',''),'0','','','','','','',sum(D.Sum_Amt)
                           FROM dbo.DONOR M  
                           RIGHT JOIN (SELECT Donor_Id ,Max(Donate_Amt) AS Max_Amt, Sum(Donate_Amt) AS Sum_Amt ,Count(Donor_Id) AS Times  
						               FROM DONATE 
                                       WHERE ISNULL(Issue_Type,'') != 'D' ";
        //搜尋條件
        if (HFD_Donate_Amt_season2.Value != "")
        {
            strSql += " and Donate_Amt > '" + HFD_Donate_Amt_season2.Value + "' ";
        }
        else
        {
            strSql += " and Donate_Amt > 0 ";
        }
        if (HFD_DonateDateS_season2.Value != "")
        {
            strSql += " And Donate_Date>='" + HFD_DonateDateS_season2.Value + "' ";
        }
        if (HFD_DonateDateE_season2.Value != "")
        {
            strSql += " And Donate_Date<='" + HFD_DonateDateE_season2.Value + "' ";
        }
        strSql += " GROUP BY Donor_Id ) D ON M.Donor_Id = D.Donor_Id WHERE 1=1 ";
        if (HFD_Is_Abroad_season2.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_season2.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_season2.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        strSql += " ORDER BY D.Sum_Amt";
        return strSql;
    }
}