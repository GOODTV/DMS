using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonateMgr_PledgeReturnError_Print : BasePage
{
    string TableWidth = "800px";

    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        PrintReport();
    }
    //---------------------------------------------------------------------------
    private void PrintReport()
    {
        string strSql = @"select ROW_NUMBER() OVER(ORDER BY P.Pledge_Id) AS '序號',P.Pledge_Id as '授權編號',REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,SR.Donate_Amt),1),'.00','') as '授權金額', D.Donor_Id as '捐款人編號', D.Donor_Name as '捐款人姓名'
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
        Dictionary<string, object> dict = new Dictionary<string, object>();

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
        }

        DataTable dtRet = CaseUtil.PledgeReturnError_Print(dt);

        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        table.Width = "100%";
        css = table.Style;
        css.Add("font-size", "12px");
        css.Add("font-family", "細明體");
        css.Add("line-height", "20px");

        row = new HtmlTableRow();

        int iCtrl = 0;
        foreach (DataColumn dc in dtRet.Columns)
        {
            cell = new HtmlTableCell();
            cell.Style.Add("background-color", " #FFE4C4 ");
            cell.Style.Add("border-left", ".5pt solid windowtext");
            cell.Style.Add("border-top", ".5pt solid windowtext");
            cell.Style.Add("border-right", ".5pt solid windowtext");
            cell.Style.Add("border-bottom", ".5pt solid windowtext");

            //if (iCtrl == 0 || iCtrl == 9 || iCtrl == 10 || iCtrl == 11)
            //{
            //    cell.Width = "50mm";
            //}
            //else if (iCtrl == 4)
            //{
            //    cell.Width = "70mm";
            //}
            //else if (iCtrl == 1 || iCtrl == 7 || iCtrl == 8)
            //{
            //    cell.Width = "90mm";
            //}
            //else if (iCtrl == 6)
            //{
            //    cell.Width = "100mm";
            //}
            //else
            //{
            //    cell.Width = "80mm";
            //}
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

        foreach (DataRow dr in dtRet.Rows)
        {
            row = new HtmlTableRow();
            foreach (object objItem in dr.ItemArray)
            {
                cell = new HtmlTableCell();
                cell.Style.Add("border-left", ".5pt solid windowtext");
                cell.Style.Add("border-top", ".5pt solid windowtext");
                cell.Style.Add("border-right", ".5pt solid windowtext");
                cell.Style.Add("border-bottom", ".5pt solid windowtext");
                cell.InnerHtml = objItem.ToString() == "" ? "&nbsp" : objItem.ToString();
                row.Cells.Add(cell);
            }
            table.Rows.Add(row);
        }
        //轉成 html 碼
        StringWriter sw = new StringWriter();

        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);
        GridList.Text = "台銀授權失敗回覆檔暨捐款人資料<br/>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
}