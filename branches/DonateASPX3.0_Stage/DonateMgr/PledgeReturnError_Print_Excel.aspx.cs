using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonateMgr_PledgeReturnError_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "1450px";

    protected void Page_Load(object sender, EventArgs e)
    {
        string strSql = @"select ROW_NUMBER() OVER(ORDER BY P.Pledge_Id) AS '序號',P.Pledge_Id as '授權編號'
                            ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,SR.Donate_Amt),1),'.00','') as '授權金額'
                            , D.Donor_Id as '捐款人編號', D.Donor_Name as '捐款人姓名'
    						,p.[Card_Bank] as [發卡銀行],p.[Account_No] as [信用卡卡號],p.[Valid_Date] as [效期],Authorize as [末三碼CVV]
                            ,IsNull(D.ZipCode,'') AS '郵遞區號'  
                            ,Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') 
		                            else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
                            ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
			                            (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
			                            Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
		                            Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
			                            Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話'  
                            ,IsNull(Cellular_Phone,'') AS '手機'
                            ,PS.ErrorCode as '授權失敗碼' ,PS.Note as '授權失敗原因'
                            from Pledge_Send_Return SR
                            left join Pledge P on SR.Pledge_Id = P.Pledge_Id
                            left join Donor D on P.Donor_Id = D.Donor_Id
                            LEFT JOIN dbo.CODECITYNew C1 ON D.City = C1.ZipCode
                            LEFT JOIN dbo.CODECITYNew C2 ON D.Area = C2.ZipCode
                            left join Pledge_Status PS on SR.Return_Status_No = PS.ErrorCode
                            where SR.Return_Status='N'
                            order by SR.Pledge_Id";
        Print_Excel(strSql);
    }
    //---------------------------------------------------------------------------
    private void Print_Excel(string strSql)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            ShowSysMsg("查無資料!!!");
            return;
        }

        GridList.Text = GetTitle(dt.Rows[0]) + GetTable1(dt.Rows[0], strSql);
        //GetTableFooter();
        Util.OutputTxt(GridList.Text, "1", "PledgeReturnError");
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
        css.Add("font-size", "13px");
        css.Add("font-family", "標楷體");
        css.Add("width", TableWidth);

        string strFontSize = "24px";

        string strTitle = "台銀授權失敗回覆檔暨捐款人資料<br/> ";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 17;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    //---------------------------------------------------------------------------
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

        DataTable dtRet = CaseUtil.PledgeReturnError_Print(dt);

        table.Border = 1;
        table.CellPadding = 3;
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
            cell.Style.Add("background-color", " #FFE4C4 ");
            /*
            cell.Style.Add("border-left", ".5pt solid windowtext");
            cell.Style.Add("border-top", ".5pt solid windowtext");
            cell.Style.Add("border-right", ".5pt solid windowtext");
            cell.Style.Add("border-bottom", ".5pt solid windowtext");
            */

            if (iCtrl == 0)
            {
                cell.Width = "30";
            }
            else
            {
                cell.Width = "90mm";
            }
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp;" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

        foreach (DataRow drow in dtRet.Rows)
        {

            row = new HtmlTableRow();
            foreach (object objItem in drow.ItemArray)
            {
                cell = new HtmlTableCell();
                /*
                cell.Style.Add("border-left", ".5pt solid windowtext");
                cell.Style.Add("border-top", ".5pt solid windowtext");
                cell.Style.Add("border-right", ".5pt solid windowtext");
                cell.Style.Add("border-bottom", ".5pt solid windowtext");
                */

                cell.InnerHtml = objItem.ToString() == "" ? "&nbsp;" : objItem.ToString();
                row.Cells.Add(cell);
            }
            table.Rows.Add(row);
        }

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
}