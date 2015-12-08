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

public partial class Report_Custom_Report_Donate_Week_Report4_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Donate_Amt_week4.Value = Util.GetQueryString("Donate_Amt");
            HFD_DonateDateS_week4.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_week4.Value = Util.GetQueryString("DonateDateE");
            HFD_AccumulateDateS_week4.Value = Util.GetQueryString("AccumulateDateS");
            HFD_AccumulateDateE_week4.Value = Util.GetQueryString("AccumulateDateE");
            HFD_Is_Abroad_week4.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_week4.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_week4.Value = Util.GetQueryString("Sex");
            Print_Excel();
        }
    }
    //---------------------------------------------------------------------------
    private void Print_Excel()
    {
        string strSql = Sql();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            ShowSysMsg("查無資料!!!");
            return;
        }
        string strSql2 = Sql2();
        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        DataTable dt2 = NpoDB.GetDataTableS(strSql2, dict2);

        GridList.Text = GetTitle(dt.Rows[0]) + GetTable1(dt.Rows[0], strSql, dt2.Rows[0], strSql2);
        //GetTableFooter();
        Util.OutputTxt(GridList.Text, "1", "donate_week_report4");
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
        string strTitle = "週報用大額捐款人明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>捐款期間：" + HFD_DonateDateS_week4.Value + "~" + HFD_DonateDateE_week4.Value + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款日期排序</span>";
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
    private string GetTable1(DataRow dr, string strSql, DataRow dr2, string strSql2)
    {
        //組 table
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;
        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        DataTable dt2 = NpoDB.GetDataTableS(strSql2, dict2);
        DataTable dtRet = CaseUtil.Donate_Week_Report4_Print(dt,dt2,true);

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
                cell.Width = "50mm";
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
                cell.Width = "50mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 3)
            {
                cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 4)
            {
                cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 5)
            {
                cell.Width = "50mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 6)
            {
                cell.Width = "40mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 7)
            {
                cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 8)
            {
                cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 9)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
            }
            else if (iCtrl == 10)
            {
                cell.Width = "100mm";
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
                //cell.Width = "40mm";
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
                    cell.Style.Add("text-align", "center");
                }
                else if (Col == 10)
                {
                    cell.Style.Add("text-align", "left");
                    cell.Width = "100mm";
                }
                else if (Col == 4)
                {
                    cell.Style.Add("text-align", "right");
                }
                else if (Col == 7)
                {
                    cell.Style.Add("text-align", "right");
                }
                else if (Col == 2)
                {
                    cell.Width = "50mm";
                    cell.Style.Add("text-align", "left");
                }
                else if (Col == 5)
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
        string strSql = @"  SELECT D.Donor_Id AS '編號',CONVERT(VARCHAR, D.Donate_Date, 111 ) AS '捐款日期'   
                                 ,M.Donor_Name AS '天使姓名' ,M.Title AS '稱謂'  
								 ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Donate_Amt),1),'.00','') AS '捐款金額' 
                                 ,Case When CONVERT(VARCHAR, Begin_DonateDate, 111 ) >= Cast('" + HFD_DonateDateS_week4.Value + "'as date) and CONVERT(VARCHAR, Begin_DonateDate, 111 ) <= Cast('" + HFD_DonateDateE_week4.Value + "'as date) Then '新' Else '舊' END '新/舊天使' ";
               strSql += @"      ,Case When M.Donor_Name like '%主知名' then CONVERT(VARCHAR, D.Donate_Date, 111 ) ELSE CONVERT(VARCHAR, Begin_DonateDate, 111 ) END AS '首捐日期'
                                 ,Case When M.Donor_Name like '%主知名' then '' ELSE REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D3.Sum_Amt),1),'.00','') END AS '累計金額',D.Donate_Purpose AS '捐款用途'  
								 ,IsNull(M.ZipCode,'') AS '郵遞區號'  
								 ,Case When M.IsAbroad = 'Y' Then IsNull(OverseasAddress,'')  
								 		else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
							     ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = ''   
										Then (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
								        Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office   
										Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話' ,IsNull(Cellular_Phone,'') AS '手機' ,IsNull(Email,'') AS 'Email'   
							FROM dbo.DONOR M  
							inner JOIN (SELECT Donor_Id ,Donate_Id,Donate_Date,Donate_Amt,Donate_Purpose,Donate_Payment  
							            FROM DONATE ";

        //搜尋條件
        string Donate_Where = "";
        string Donate_Where1 = "";
        string Donate_Where2 = "";
        if (HFD_Donate_Amt_week4.Value != "")
        {
            Donate_Where = " WHERE Donate_Amt >= '" + HFD_Donate_Amt_week4.Value + "'";
        }
        else
        {
            Donate_Where = " WHERE Donate_Amt > 0";
        }
        if (HFD_DonateDateS_week4.Value != "")
        {
            Donate_Where1 += " And Donate_Date >= '" + HFD_DonateDateS_week4.Value + "'";
        }
        if (HFD_DonateDateE_week4.Value != "")
        {
            Donate_Where1 += " And Donate_Date <= '" + HFD_DonateDateE_week4.Value + "'";
        }
        if (HFD_AccumulateDateS_week4.Value != "")
        {
            Donate_Where2 += " And Donate_Date >= '" + HFD_AccumulateDateS_week4.Value + "'";
        }
        if (HFD_AccumulateDateE_week4.Value != "")
        {
            Donate_Where2 += " And Donate_Date <= '" + HFD_AccumulateDateE_week4.Value + "'";
        }

        strSql += Donate_Where + Donate_Where1 + @" and Issue_Type != 'D') D ON M.Donor_Id = D.Donor_Id 
                  left JOIN (SELECT Donor_Id ,Count(Donor_Id) AS Times FROM dbo.DONATE ";
        strSql += Donate_Where + Donate_Where1 + @" and Issue_Type != 'D' group by Donor_Id ) D1 ON M.Donor_Id = D1.Donor_Id 
                  left JOIN (SELECT Donor_Id ,Count(Donor_Id) AS Times FROM dbo.DONATE ";
        strSql += Donate_Where + Donate_Where2 + @" and Issue_Type != 'D' group by Donor_Id ) D2 ON M.Donor_Id = D2.Donor_Id 
                  left JOIN (SELECT Donor_Id , Sum(Donate_Amt) AS Sum_Amt FROM dbo.DONATE ";
        strSql += "WHERE Donate_Amt > 0 " + Donate_Where2 + @" and Issue_Type != 'D' group by Donor_Id ) D3 ON M.Donor_Id = D3.Donor_Id 
                                                                LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
                                                                LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode ";
        //ORDER BY D.Donate_Purpose,D.Donate_Amt,D.Donate_Date,Last_DonateDate ";
        if (HFD_Is_Abroad_week4.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_week4.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_week4.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        strSql += " ORDER BY D.Donate_Purpose,D.Donate_Amt,D.Donate_Date,Last_DonateDate ";
        return strSql;
    }
    private string Sql2()
    {
        string strSql2 = @"  select count(D.Donor_Id),REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(D.Donate_Amt)),1),'.00','')
							FROM dbo.DONOR M  
							inner JOIN (SELECT Donor_Id ,Donate_Id,Donate_Date,Donate_Amt,Donate_Purpose,Donate_Payment  
							            FROM DONATE ";
        //搜尋條件
        string Donate_Where = "";
        string Donate_Where1 = "";
        string Donate_Where2 = "";
        if (HFD_Donate_Amt_week4.Value != "")
        {
            Donate_Where = " WHERE Donate_Amt >= '" + HFD_Donate_Amt_week4.Value + "'";
        }
        else
        {
            Donate_Where = " WHERE Donate_Amt > 0";
        }
        if (HFD_DonateDateS_week4.Value != "")
        {
            Donate_Where1 += " And Donate_Date >= '" + HFD_DonateDateS_week4.Value + "'";
        }
        if (HFD_DonateDateE_week4.Value != "")
        {
            Donate_Where1 += " And Donate_Date <= '" + HFD_DonateDateE_week4.Value + "'";
        }
        if (HFD_AccumulateDateS_week4.Value != "")
        {
            Donate_Where2 += " And Donate_Date >= '" + HFD_AccumulateDateS_week4.Value + "'";
        }
        if (HFD_AccumulateDateE_week4.Value != "")
        {
            Donate_Where2 += " And Donate_Date <= '" + HFD_AccumulateDateE_week4.Value + "'";
        }

        strSql2 += Donate_Where + Donate_Where1 + @" and Issue_Type != 'D') D ON M.Donor_Id = D.Donor_Id 
                  left JOIN (SELECT Donor_Id ,Count(Donor_Id) AS Times FROM dbo.DONATE ";
        strSql2 += Donate_Where + Donate_Where1 + @" and Issue_Type != 'D' group by Donor_Id ) D1 ON M.Donor_Id = D1.Donor_Id 
                  left JOIN (SELECT Donor_Id ,Count(Donor_Id) AS Times FROM dbo.DONATE ";
        strSql2 += Donate_Where + Donate_Where2 + @" and Issue_Type != 'D' group by Donor_Id ) D2 ON M.Donor_Id = D2.Donor_Id 
                  left JOIN (SELECT Donor_Id , Sum(Donate_Amt) AS Sum_Amt FROM dbo.DONATE ";
        strSql2 += "WHERE Donate_Amt > 0 " + Donate_Where2 + @" and Issue_Type != 'D' group by Donor_Id ) D3 ON M.Donor_Id = D3.Donor_Id 
                                                                LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
                                                                LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode ";
       
        if (HFD_Is_Abroad_week4.Value == "True")
        {
            strSql2 += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_week4.Value == "True")
        {
            strSql2 += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_week4.Value == "True")
        {
            strSql2 += " And IsNull(Sex,'') <> '歿' ";
        }
        return strSql2;
    }
}