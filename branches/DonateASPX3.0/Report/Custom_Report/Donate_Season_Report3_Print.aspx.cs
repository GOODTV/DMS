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

public partial class Report_Custom_Report_Donate_Season_Report3_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Donate_Amt_season3.Value = Util.GetQueryString("Donate_Amt");
            HFD_Donate_Total_Amt_season3.Value = Util.GetQueryString("Donate_Total_Amt");
            HFD_Donate_Purpose_season3.Value = Util.GetQueryString("Donate_Purpose");
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

        DataTable dtRet = CaseUtil.Donate_Season_Report3_Print(dt);

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
                cell.Width = "65mm";
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
                //cell.Width = "80mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 3)
            {
                //ell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 4)
            {
                //cell.Width = "60mm";
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
                //cell.Width = "70mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 11)
            {
                //cell.Width = "70mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 12)
            {
                //cell.Width = "70mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 13)
            {
                //cell.Width = "70mm";
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
                else if (Col == 13)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                    cell.Style.Add("text-align", "left");
                }
                else if (Col == 8)
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

        GridList.Text = DateTime.Now.AddYears(-1911).Year + "年" + DateTime.Now.Month + "月份  捐款用途各項總額明細表<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + DateTime.Now.ToString() + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依累計奉獻金額排序</span><br>" + htw.InnerWriter.ToString();
        //Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql = @"  SELECT D.Donor_Id AS '編號'   
									,M.Donor_Name AS '天使姓名' ,IsNull(M.Title,'') AS '稱謂' ,IsNull(Donor_Type,'') AS '身份別'  
									,D.Donate_Purpose AS '奉獻用途',REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Amt),1),'.00','') AS '累計奉獻金額' ,D.Times AS '奉獻次數'  
									,IsNull(M.ZipCode,'') AS '郵遞區號'  
									,Case When M.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'')  
									    else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
									,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
												(Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
												Else Tel_Office + ' <br>Ext.' + Tel_Office_Ext End)  
									Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
										  Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' <br>Ext.' + Tel_Office_Ext End) End '電話'  
									,IsNull(Cellular_Phone,'') AS '手機'  
									,CONVERT(VARCHAR, Birthday, 111) AS '生日' ,CONVERT(VARCHAR, Begin_DonateDate, 111) AS '首捐日期' ,CONVERT(VARCHAR, Last_DonateDate, 111) AS '末捐日期'  
							FROM dbo.DONOR M  
							RIGHT JOIN (SELECT Donor_Id ,Donate_Purpose ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times  
										FROM DONATE 
                                        WHERE ISNULL(Issue_Type,'') != 'D' ";

        //搜尋條件
        string[] Donate_Purpose;
        int count = 0;
        if (HFD_Donate_Purpose_season3.Value != "")
        {
            char[] ch = new char[] { ',' };
            Donate_Purpose = HFD_Donate_Purpose_season3.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
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
        else
        {
            strSql += " and Donate_Purpose <> '' ";
        }
        if (HFD_Donate_Amt_season3.Value != "")
        {
            strSql += " And Donate_Amt>'" + HFD_Donate_Amt_season3.Value + "' ";
        }
        else
        {
            strSql += " AND Donate_Amt > 0 ";
        }
        if (HFD_Donate_Total_Amt_season3.Value != "")
        {
            strSql += " GROUP BY Donor_Id, Donate_Purpose HAVING SUM(Donate_Amt) >= '" + HFD_Donate_Total_Amt_season3.Value + "'  ) D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id, Donate_Purpose ) D ";
        }
        strSql += @" ON M.Donor_Id = D.Donor_Id
                     LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
                     LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode ";
        
        //以下是串UNION的SQL語法
        strSql += @" Union select '99999',REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(D.Amt)),1),'.00',''),'','','','0','','','','','','','',''
                               FROM dbo.DONOR M 
                               RIGHT JOIN (SELECT Donor_Id ,Donate_Purpose ,Sum(Donate_Amt) AS Amt ,Count(Donor_Id) AS Times  
						                   FROM DONATE 
                                           WHERE ISNULL(Issue_Type,'') != 'D' ";
        //搜尋條件
        if (HFD_Donate_Purpose_season3.Value != "")
        {
            char[] ch = new char[] { ',' };
            Donate_Purpose = HFD_Donate_Purpose_season3.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
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
        else
        {
            strSql += " and Donate_Purpose <> '' ";
        }
        if (HFD_Donate_Amt_season3.Value != "")
        {
            strSql += " And Donate_Amt>'" + HFD_Donate_Amt_season3.Value + "' ";
        }
        else
        {
            strSql += " AND Donate_Amt > 0 ";
        }
        if (HFD_Donate_Total_Amt_season3.Value != "")
        {
            strSql += " GROUP BY Donor_Id, Donate_Purpose HAVING SUM(Donate_Amt) >= '" + HFD_Donate_Total_Amt_season3.Value + "'  ) D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id, Donate_Purpose ) D ";
        }
        strSql += " ON M.Donor_Id = D.Donor_Id ";
        strSql += " ORDER BY '累計奉獻金額'  Desc ";
        return strSql;
    }
}