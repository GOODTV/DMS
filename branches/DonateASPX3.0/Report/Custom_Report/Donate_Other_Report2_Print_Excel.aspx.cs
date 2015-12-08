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

public partial class Report_Custom_Report_Donate_Other_Report2_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Linked2_Name_other2.Value = Util.GetQueryString("Linked2_Name");
            HFD_ContributeDateS_other2.Value = Util.GetQueryString("ContributeDateS");
            HFD_ContributeDateE_other2.Value = Util.GetQueryString("ContributeDateE");
            HFD_Is_Abroad_other2.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_other2.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_other2.Value = Util.GetQueryString("Sex");
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
        Util.OutputTxt(GridList.Text, "1", "donate_other_report2");
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
        string strTitle = "DVD贈品索取人報表<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>贈送日期：" + HFD_ContributeDateS_other2.Value + "~" + HFD_ContributeDateE_other2.Value + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人編號排序</span>";
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 8;
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

        DataTable dtRet = CaseUtil.Donate_Other_Report2_Print(dt);

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
                //cell.Width = "55mm";
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
                //cell.Width = "55mm";
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
                else if (Col == 7)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
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
        string strSql = @"  SELECT distinct M.Donor_Id AS '編號' ,Donor_Name AS '天使姓名' ,ISNULL(Title,'') AS '稱謂'
                                   ,IsNull(M.ZipCode,'') AS '郵遞區號' 
                                   ,Case When M.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'')
	                                     Else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'	 
                                   ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then 
		                               (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office
                                        Else Tel_Office + ' Ext.' + Tel_Office_Ext End)
                                   Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office
                                        Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話' 
                                   ,G2.Goods_Qty AS '份數' ,CONVERT(varchar(12) , G2.Create_Date, 111 ) AS '資料輸入日期'
                            FROM dbo.DONOR M
                            INNER JOIN dbo.GIFT G
                            ON M.Donor_Id=G.Donor_Id
                            INNER JOIN dbo.GIFTData G2
                            ON M.Donor_Id=G2.Donor_Id and G.Contribute_Id=G2.Contribute_Id
                            INNER JOIN dbo.LINKED2 L2
                            ON L2.Ser_No=G2.Goods_Id
                            INNER JOIN dbo.LINKED L
                            ON L.Linked_Id=L2.Linked_Id
                            LEFT JOIN dbo.CODECITYNew C1
                            ON M.City = C1.ZipCode
                            LEFT JOIN dbo.CODECITYNew C2
                            ON M.Area = C2.ZipCode ";

        //搜尋條件
        if (HFD_ContributeDateS_other2.Value != "")
        {
            strSql += " WHERE Contribute_Date>='" + HFD_ContributeDateS_other2.Value + "' ";
        }
        if (HFD_ContributeDateE_other2.Value != "")
        {
            strSql += " And Contribute_Date<='" + HFD_ContributeDateE_other2.Value + "' ";
        }
        if (HFD_Linked2_Name_other2.Value != "")
        {
            strSql += " AND Goods_Name = '" + HFD_Linked2_Name_other2.Value + "' ";
        }

        if (HFD_Is_Abroad_other2.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_other2.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_other2.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        //以下是串UNION的SQL語法
        strSql += @" Union select '999999999','總筆數：',CONVERT(VARCHAR,count(*)),'筆','','','','9999/12/31'
                                FROM DONOR M
                            INNER JOIN dbo.GIFT G
                            ON M.Donor_Id=G.Donor_Id
                            INNER JOIN dbo.GIFTData G2
                            ON M.Donor_Id=G2.Donor_Id and G.Contribute_Id=G2.Contribute_Id
                            INNER JOIN dbo.LINKED2 L2
                            ON L2.Ser_No=G2.Goods_Id
                            INNER JOIN dbo.LINKED L
                            ON L.Linked_Id=L2.Linked_Id";
        //搜尋條件
        if (HFD_ContributeDateS_other2.Value != "")
        {
            strSql += " WHERE Contribute_Date>='" + HFD_ContributeDateS_other2.Value + "' ";
        }
        if (HFD_ContributeDateE_other2.Value != "")
        {
            strSql += " And Contribute_Date<='" + HFD_ContributeDateE_other2.Value + "' ";
        }
        if (HFD_Linked2_Name_other2.Value != "")
        {
            strSql += " AND Goods_Name = '" + HFD_Linked2_Name_other2.Value + "' ";
        }

        if (HFD_Is_Abroad_other2.Value == "True")
        {
            strSql += " And IsAbroad='N' ";
        }
        if (HFD_Is_ErrAddress_other2.Value == "True")
        {
            strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
        }
        if (HFD_Sex_other2.Value == "True")
        {
            strSql += " And IsNull(Sex,'') <> '歿' ";
        }
        strSql += " ORDER BY M.Donor_Id";
        return strSql;
    }
}