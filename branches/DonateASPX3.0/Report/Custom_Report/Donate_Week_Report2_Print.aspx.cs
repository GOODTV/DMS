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

public partial class Report_Custom_Report_Donate_Week_Report2_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_DonateDateS_week2.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_week2.Value = Util.GetQueryString("DonateDateE");
            HFD_Donate_Total_Begin_week2.Value = Util.GetQueryString("Donate_Total_Begin");
            HFD_Donate_Total_End_week2.Value = Util.GetQueryString("Donate_Total_End");
            HFD_Donate_Purpose_week2.Value = Util.GetQueryString("Donate_Purpose");
            HFD_Is_Abroad_week2.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_week2.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_week2.Value = Util.GetQueryString("Sex");
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

        DataTable dtRet = CaseUtil.Donate_Week_Report2_Print(dt);

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
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
                cell.Width = "100mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 2)
            {
                cell.Width = "70mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 3)
            {
                cell.Width = "50mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 4)
            {
                cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 5)
            {
                cell.Width = "50mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 6)
            {
                //cell.Width = "100mm";
                cell.Style.Add("text-align", "left");
            }
            else if (iCtrl == 7)
            {
                //cell.Width = "80mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 8)
            {
                //cell.Width = "70mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 9)
            {
                //cell.Width = "70mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 10)
            {
                cell.Width = "70mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 11)
            {
                cell.Width = "30mm";
                cell.Style.Add("text-align", "left");
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
                else if (Col == 1 || Col == 6)
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

        GridList.Text = "新捐款人明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + HFD_DonateDateS_week2.Value + "~" + HFD_DonateDateE_week2.Value + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人編號排序</span><br>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql = @" SELECT M.Donor_Id AS 'Donor_Id' ,M.Donor_Name+' '+ISNULL(M.Title,'') AS '天使姓名' ,CONVERT(VARCHAR, M.Begin_DonateDate, 111 )  AS '首捐日期'  
                                          ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Sum_Amt),1),'.00','') AS '累計奉獻金額', D.Times AS '奉獻次數'   
	                                      ,IsNull(M.ZipCode,'') AS '郵遞區號'  
	                                      ,Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'')  
		                                          else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
	                                      ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
				                                      (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
				                                      Else Tel_Office + ' <br>Ext.' + Tel_Office_Ext End)  
		                                          Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
				                                      Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' <br>Ext.' + Tel_Office_Ext End) End '電話'  
                                          ,IsNull(Cellular_Phone,'') AS '手機' ,IsNull(Email,'') AS 'Email' 
	                                      ,Case When IsSendNews='Y' Then '寄' Else '不寄' End '月刊寄送' ,Case When IsContact='N' Then 'V' Else '' End '不主動聯絡'  
                                    FROM DONOR M  
                                    inner JOIN (select  Donor_Id ,COUNT(Donor_Id) AS Times , SUM(Donate_Amt) AS Sum_Amt  
                                                from dbo.DONATE 
                                                where ISNULL(Issue_Type,'') != 'D' ";

        //搜尋條件
        if (HFD_Donate_Total_Begin_week2.Value != "")
        {
            strSql += " And Donate_Amt >= '" + HFD_Donate_Total_Begin_week2.Value + "'";
        }
        else
        {
            strSql += " And Donate_Amt > 0";
        }
        if (HFD_Donate_Total_End_week2.Value != "")
        {
            strSql += " And Donate_Amt <= '" + HFD_Donate_Total_End_week2.Value + "'";
        }
        else
        {
            strSql += "";
        }

        if (HFD_DonateDateS_week2.Value != "")
        {
            strSql += " And Create_Date >= '" + HFD_DonateDateS_week2.Value + "'";
        }
        if (HFD_DonateDateS_week2.Value != "")
        {
            strSql += " And Create_Date <= '" + HFD_DonateDateE_week2.Value + "'";
        }
        string[] Donate_Purpose;
        int count = 0;
        if (HFD_Donate_Purpose_week2.Value != "")
        {
            char[] ch = new char [] {','};
            Donate_Purpose = HFD_Donate_Purpose_week2.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
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
        strSql += "group by Donor_Id) D ";
        strSql += @"ON M.Donor_Id = D.Donor_Id 
                    LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
                    LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode
                    WHERE M.Begin_DonateDate >= '" + HFD_DonateDateS_week2.Value + @"'
                    AND M.Begin_DonateDate <= '" + HFD_DonateDateE_week2.Value + "' ";
        if (HFD_Is_Abroad_week2.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_week2.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_week2.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }

        //以下是串UNION的SQL語法
        strSql += @"Union 
                    select '99999' as 'Donor_Id', '總筆數：'  AS [天使姓名],'' AS '首捐日期','筆' AS '累計奉獻金額',count(M.Donor_Id) AS '奉獻次數' ,'總計' AS '郵遞區號', '捐款金額：'as '地址' ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(D.Sum_Amt)),1),'.00','') +'元' as '電話','' as '手機','' as 'Email','' as '月刊寄送','' as '不主動聯絡' 
                    FROM DONOR M  
                    inner JOIN (select  Donor_Id ,COUNT(Donor_Id) AS Times , SUM(Donate_Amt) AS Sum_Amt  
                                from dbo.DONATE 
                                where ISNULL(Issue_Type,'') != 'D' ";
        //搜尋條件
        if (HFD_Donate_Total_Begin_week2.Value != "")
        {
            strSql += " And Donate_Amt >= '" + HFD_Donate_Total_Begin_week2.Value + "'";
        }
        else
        {
            strSql += " And Donate_Amt > 0";
        }
        if (HFD_Donate_Total_End_week2.Value != "")
        {
            strSql += " And Donate_Amt <= '" + HFD_Donate_Total_End_week2.Value + "'";
        }
        else
        {
            strSql += "";
        }

        if (HFD_DonateDateS_week2.Value != "")
        {
            strSql += " And Create_Date >= '" + HFD_DonateDateS_week2.Value + "'";
        }
        if (HFD_DonateDateE_week2.Value != "")
        {
            strSql += " And Create_Date <= '" + HFD_DonateDateE_week2.Value + "'";
        }
        if (HFD_Donate_Purpose_week2.Value != "")
        {
            char[] ch = new char [] {','};
            Donate_Purpose = HFD_Donate_Purpose_week2.Value.Split(ch, StringSplitOptions.RemoveEmptyEntries);
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
        strSql += "group by Donor_Id) D ";
        strSql += @"ON M.Donor_Id = D.Donor_Id  
                    WHERE M.Begin_DonateDate >= '" + HFD_DonateDateS_week2.Value + @"'
                    AND M.Begin_DonateDate <= '" + HFD_DonateDateE_week2.Value + "'";
        if (HFD_Is_Abroad_week2.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_week2.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_week2.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        strSql += "  Order by M.Donor_Id ";
        return strSql;
    }
}