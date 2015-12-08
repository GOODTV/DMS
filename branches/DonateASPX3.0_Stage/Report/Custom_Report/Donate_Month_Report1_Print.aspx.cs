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

public partial class Report_Custom_Report_Donate_Month_Report1_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Donate_Amt_month1.Value = Util.GetQueryString("Donate_Amt");
            HFD_Donate_Total_Begin_month1.Value = Util.GetQueryString("Donate_Total_Begin");
            HFD_Donate_Total_End_month1.Value = Util.GetQueryString("Donate_Total_End");
            HFD_Donor_Type_month1.Value = Util.GetQueryString("Donor_Type");
            HFD_Birthday_Month_month1.Value = Util.GetQueryString("Birthday_Month");
            HFD_Is_Abroad_month1.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_month1.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_month1.Value = Util.GetQueryString("Sex");
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

        DataTable dtRet = CaseUtil.Donate_Month_Report1_Print(dt);

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
                cell.Width = "60mm";
                cell.Height = "40mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-left", ".5pt solid windowtext");
            }
            else if (iCtrl == 1)
            {
                cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 2)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 3)
            {
                cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 4)
            {
                cell.Width = "50mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 5)
            {
                cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 6)
            {
                cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 7)
            {
                //cell.Width = "80mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 8)
            {
                //cell.Width = "100mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 9)
            {
                //cell.Width = "60mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 10)
            {
                //cell.Width = "60mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 11)
            {
                cell.Width = "30mm";
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
                else if (Col == 11)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                    cell.Style.Add("text-align", "center");
                }
                else if (Col == 7)
                {
                    cell.Style.Add("text-align", "left");
                }
                else if (Col == 4)
                {
                    cell.Style.Add("text-align", "right");
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

        GridList.Text = "捐款人生日明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + DateTime.Now.ToString() + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人生日日期排序</span><br>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql = @"SELECT M.Donor_Id AS '編號'   
                                    ,M.Donor_Name AS '天使姓名' ,ISNULL(M.Title,'') AS '稱謂'  
								    ,M.Donor_Type AS '身分別'  
								    ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Sum_Amt),1),'.00','') AS '累計奉獻金額' ,D.Times AS '奉獻次數'  
								    ,IsNull(M.ZipCode,'') AS '郵遞區號'  
								    ,Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') 
										  else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
								    ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
												(Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
												Else Tel_Office + ' <br>Ext.' + Tel_Office_Ext End)  
										  Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
												Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' <br>Ext.' + Tel_Office_Ext End) End '電話'  
								    ,IsNull(Cellular_Phone,'') AS '手機',IsNull(Email,'') AS 'Email'  
								    ,RIGHT(CONVERT(varchar(12) , Birthday, 111 ),5)  AS '生日' 
							FROM DONOR M  
							RIGHT JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
							            FROM DONATE 
                                        WHERE ISNULL(Issue_Type,'') != 'D' ";

        //搜尋條件
        if (HFD_Donate_Amt_month1.Value != "")
        {
            strSql += " and Donate_Amt > '" + HFD_Donate_Amt_month1.Value + "'";
        }
        else
        {
            strSql += " and Donate_Amt > 0";
        }
        if (HFD_Donate_Total_Begin_month1.Value != "" && HFD_Donate_Total_End_month1.Value != "")
        {
            strSql += " GROUP BY Donor_Id Having SUM(Donate_Amt) BETWEEN '" + HFD_Donate_Total_Begin_month1.Value + "' AND '" + HFD_Donate_Total_End_month1.Value + "' ) D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id ) D";
        }
        strSql += @" ON M.Donor_Id = D.Donor_Id 
                     LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
                     LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode ";
        if (HFD_Birthday_Month_month1.Value != "")
        {
            strSql += " WHERE DATEPART(Month,M.Birthday) = '" + HFD_Birthday_Month_month1.Value + "' ";
        }
        else
        {
            strSql += " WHERE DATEPART(Month,M.Birthday) <> '' ";
        }
        string[] Donor_Type;
        int count = 0;
        if (HFD_Donor_Type_month1.Value != "")
        {
            char[] ch = new char[] { ',' };
            Donor_Type = HFD_Donor_Type_month1.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
            count = Donor_Type.Length;

            bool first = true; int cnt = 0;
            for (int i = 0; i < count; i++)
            {
                if (first)
                {
                    strSql += " and (Donor_Type='" + Donor_Type[i] + "' "; first = false; cnt++;
                }
                else
                {
                    strSql += " or Donor_Type='" + Donor_Type[i] + "' "; cnt++;
                }
            }
            if (cnt != 0) strSql += ')';
        }
        if (HFD_Is_Abroad_month1.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_month1.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_month1.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        //ORDER BY D.Donate_Purpose,D.Donate_Amt,D.Donate_Date,Last_DonateDate ";
        //以下是串UNION的SQL語法
        strSql += @" Union 
						select '','總筆數：','筆','','',count(M.Donor_Id),'','','','','','9999/12/31'
						    FROM DONOR M  
					        RIGHT JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
						FROM DONATE 
                        WHERE ISNULL(Issue_Type,'') != 'D' ";
        //搜尋條件
        if (HFD_Donate_Amt_month1.Value != "")
        {
            strSql += " and Donate_Amt > '" + HFD_Donate_Amt_month1.Value + "'";
        }
        else
        {
            strSql += " and Donate_Amt > 0";
        }
        if (HFD_Donate_Total_Begin_month1.Value != "" && HFD_Donate_Total_End_month1.Value != "")
        {
            strSql += " GROUP BY Donor_Id Having SUM(Donate_Amt) BETWEEN '" + HFD_Donate_Total_Begin_month1.Value + "' AND '" + HFD_Donate_Total_End_month1.Value + "' ) D ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id ) D";
        }
        strSql += " ON M.Donor_Id = D.Donor_Id ";
        if (HFD_Birthday_Month_month1.Value != "")
        {
            strSql += " WHERE DATEPART(Month,M.Birthday) = '" + HFD_Birthday_Month_month1.Value + "' ";
        }
        else
        {
            strSql += " WHERE DATEPART(Month,M.Birthday) <> '' ";
        }
        if (HFD_Donor_Type_month1.Value != "")
        {
            char[] ch = new char[] { ',' };
            Donor_Type = HFD_Donor_Type_month1.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
            count = Donor_Type.Length;

            bool first = true; int cnt = 0;
            for (int i = 0; i < count; i++)
            {
                if (first)
                {
                    strSql += " and (Donor_Type='" + Donor_Type[i] + "' "; first = false; cnt++;
                }
                else
                {
                    strSql += " or Donor_Type='" + Donor_Type[i] + "' "; cnt++;
                }
            }
            if (cnt != 0) strSql += ')';
        }
        if (HFD_Is_Abroad_month1.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_month1.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_month1.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        strSql += " Order By RIGHT(CONVERT(varchar(12) , Birthday, 111 ),5) ";
        return strSql;
    }
}