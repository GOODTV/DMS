using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class Donate_DonateMonthReport_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string Condition = Session["Condition"].ToString();
        PrintReport(Condition);
    }
    //---------------------------------------------------------------------------
    private void PrintReport(string Condition)
    {
        string strSql = @" select ROW_NUMBER() OVER(ORDER BY do.Invoice_No) AS 序號 , do.Invoice_No as 收據編號, CONVERT(VARCHAR(10) , do.Donate_Date, 111 ) as 捐款日期,
                            Dr.Donor_Name as 捐款人, do.Donor_Id as 編號, 
                            Case when Donate_Forign <> '' then REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,do.Donate_Amt),1),'.00','')+'('+Donate_Forign+ISNULL(CONVERT(VARCHAR,CONVERT(MONEY,Donate_ForignAmt),1),'')+')' else REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,do.Donate_Amt),1),'.00','') end as 捐款金額,
                            do.Donate_Purpose as 捐款用途, do.Donate_Payment as 捐款方式,do.Invoice_Pre,dp.CodeName as 付款方式 
                            From Donate do 
                            join Dept dn on do.Dept_Id=dn.DeptId
                            left join Donor dr on dr.Donor_Id = do.Donor_Id
                            left join (select * from (select ROW_NUMBER() OVER(PARTITION by orderid ORDER BY Ser_No) AS ROWID,*
                                 from DONATE_IEPAY) as P1 where ROWID = 1) as di on do.od_sob = di.orderid 
                            left join Donate_IePayType dp on di.paytype = dp.CodeID 
                            where dr.DeleteDate is null and do.Issue_Type !='D'";
        strSql += Condition;
        strSql += @"union
                        select '99999999' as [序號] , '列印日期：' as [收據編號], CONVERT(VARCHAR, GETDATE(), 120 ) as [捐款日期],'總計' as [捐款人],'' as [編號],
                        REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(do.Donate_Amt)),1),'.00','') as [捐款金額], '' as [捐款用途],'','' ,''  
                         from Donate do left join Donor dr on do.Donor_Id = Dr.Donor_Id 
                         left join (select * from (select ROW_NUMBER() OVER(PARTITION by orderid ORDER BY Ser_No) AS ROWID,*
                         from DONATE_IEPAY) as P1 where ROWID = 1) as di on do.od_sob = di.orderid 
                         left join Donate_IePayType dp on di.paytype = dp.CodeID 
                         where Dr.DeleteDate is null and do.Issue_Type != 'D'";
        strSql += Condition;
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "沒有符合條件的資料可以列印!!";
            return;
        }

        DataTable dtRet = CaseUtil.DonateMonthReport_Print(dt);

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
            cell.Style.Add("border-top", "1px solid windowtext");
            //cell.Style.Add("border-right", ".5pt solid windowtext");
            cell.Style.Add("border-bottom", "1px solid windowtext");
            cell.Style.Add("text-align", "center");
            if (iCtrl == 0)
            {
                //cell.Width = "90mm";
                cell.Style.Add("border-left", "1px solid windowtext");
            }
            if (iCtrl == 1)
            {
                //cell.Width = "90mm";
            }
            if (iCtrl == 2)
            {
                //cell.Width = "90mm";
            }
            if (iCtrl == 3)
            {
                //cell.Width = "100mm";
            }
            if (iCtrl == 4)
            {
                //cell.Width = "100mm";
            }
            if (iCtrl == 5)
            {
                //cell.Width = "60mm";
            }
            if (iCtrl == 6)
            {
                //cell.Width = "90mm";
                cell.Style.Add("border-right", "1px solid windowtext");
            }
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

        int Row_cut = 0;

        foreach (DataRow dr in dtRet.Rows)
        {
            int Col = 0;
            Row_cut++;
            row = new HtmlTableRow();
            foreach (object objItem in dr.ItemArray)
            {
                cell = new HtmlTableCell();
                //cell.Style.Add("border-left", ".5pt solid windowtext");
                cell.Style.Add("border-top", "1px solid windowtext");
                //cell.Style.Add("border-right", ".5pt solid windowtext");
                cell.Style.Add("border-bottom", "1px solid windowtext");
                cell.Style.Add("text-align", "center");
                // 2014/8/1 增加最後一行的簽名欄調高1倍的高度   by Samson
                if (Row_cut == dtRet.Rows.Count)
                    cell.Style.Add("line-height", "40px");

                cell.InnerHtml = objItem.ToString() == "" ? "&nbsp" : objItem.ToString();
                //if()
                if (Col == 0)
                {
                    cell.Style.Add("border-left", "1px solid windowtext");
                }
                if (Col == 3)
                {
                    cell.Style.Add("width","180px");
                }
                if (Col == 6)
                {
                    cell.Style.Add("border-right", "1px solid windowtext");
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

        GridList.Text = "捐款報表(財務)<br/>" + "<span style=' font-size: 9pt; font-family: 新細明體;'></span><br><span style=' font-size: 9pt; font-family: 新細明體;'>依收據編號排序</span><br><div align='right'><span style=' font-size: 9pt; font-family: 新細明體;'>幣別：新台幣&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
}