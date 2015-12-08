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

public partial class Report_Custom_Report_Donate_Other_Report3_Print_Excel : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Is_SameName_other3.Value = Util.GetQueryString("Is_SameName");
            HFD_Is_SameAddress_other3.Value = Util.GetQueryString("Is_SameAddress");
            HFD_Address_other3.Value = Util.GetQueryString("Address");
            HFD_Is_SameTel_other3.Value = Util.GetQueryString("Is_SameTel");
            HFD_Tel_other3.Value = Util.GetQueryString("Tel");
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
        Util.OutputTxt(GridList.Text, "1", "donate_other_report3");
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
        string strTitle = "雷同聯絡資料明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>查詢日期：" + DateTime.Now.ToString() + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人姓名或地址排序</span>";
        row = new HtmlTableRow();
        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = 11;
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

        DataTable dtRet = CaseUtil.Donate_Other_Report3_Print(dt);

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
            else if (iCtrl == 10)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-right", ".5pt solid windowtext");
            }
            else
            {
                cell.Style.Add("text-align", "center");
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
                else if (Col == 10)
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
        string strSql = @"  Select M.Donor_Id AS '捐款人編號',Case When IsNull(Donor_Pwd,'') = '' then ''+Donor_Name else '*'+Donor_Name End AS '捐款人姓名',ISNULL(Sex,'') AS '性別',ISNULL(Email,'') AS 'Email'
                                    ,S.付款方式,S.狀態,S.捐款時間
                                    ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then 
				                                     (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office
                                                      Else Tel_Office + ' Ext.' + Tel_Office_Ext End)
                                                 Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office
                                                       Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話'
                                    ,IsNull(Cellular_Phone,'') AS '手機' 
                                    ,Case When IsAbroad = 'Y' Then IsNull(OverseasAddress,'')
	                                             Else IsNull(M.ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '通訊地址'
                                    ,Case When IsAbroad = 'Y' Then IsNull(Invoice_OverseasAddress,'')
	                                             Else IsNull(M.Invoice_ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Invoice_Address END ' 收據地址'
                                    From dbo.DONOR M
                                    left Join (select distinct P.[param] as [捐款人編號]
                                    ,(case when isnull(I.codename,'') <> '' then i.codename else I.CodeID end ) as [付款方式] 
                                    ,(case when P.IEPAY_returnOK is not null then '<font color=cornflowerblue>金流處理中</font>('+IEPAY_returnOK+')' else I.PendingMsg end) as [狀態]
                                    ,P.authdate as [捐款時間] 
                                    from DONATE_IEPAY  as P 
                                    left Join Donate_IePayType AS I on P.paytype = I.CodeID
                                    left join Donate as N on P.orderid=N.od_sob
                                    inner join Donor D on D.Donor_Id=P.param
                                    where Status ='0' and paytype <> '1' 
                                    --and N.Invoice_No is NULL
                                    ) as S on S.[捐款人編號] = M.Donor_Id
                                    LEFT JOIN dbo.CODECITYNew C1
                                    ON M.City = C1.ZipCode
                                    LEFT JOIN dbo.CODECITYNew C2
                                    ON M.Area = C2.ZipCode
                                    Where M.Donor_Id In
                                    (Select D1.Donor_Id From dbo.DONOR D1
                                    Left Join dbo.DONOR D2 ";

        //搜尋條件
        if (HFD_Is_SameName_other3.Value == "True")
        {
            if (HFD_Is_SameAddress_other3.Value == "False")  //只選同姓名 或 同電話
            {
                strSql += @" ON D1.Donor_Name = D2.Donor_Name";
                if (HFD_Is_SameTel_other3.Value == "True")
                {
                    if (HFD_Tel_other3.Value == "1")  //手機
                    {
                        strSql += " And ISNULL(D1.Cellular_Phone,'') = ISNULL(D2.Cellular_Phone,'') and D1.Cellular_Phone <>'' and D2.Cellular_Phone <>''";
                    }
                    else if (HFD_Tel_other3.Value == "2") //室內電話
                    {
                        strSql += " And ISNULL(D1.Tel_Office,'') = ISNULL(D2.Tel_Office,'') and D1.Tel_Office <>'' and D2.Tel_Office <>''";
                    }
                }
                strSql += @" Where D1.Donor_Id <> D2.Donor_Id 
                             and D1.DeleteDate is null and D2.DeleteDate is null
                             Group by D1.Donor_Id)
                             Order by Donor_Name, M.Donor_Id, IsNull(M.ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address ";
            }
            else  //同時選同姓名、同地址或同電話
            {
                strSql += " ON D1.Donor_Name = D2.Donor_Name";
                if (HFD_Address_other3.Value == "1")  //通訊地址
                {
                    strSql += @" And IsNull(D1.City,'') = IsNull(D2.City,'') 
                                 And IsNull(D1.Area,'') = IsNull(D2.Area,'')
                                 And IsNull(D1.Street,'') = IsNull(D2.Street,'')
                                 --And IsNull(D1.Address,'') = IsNull(D2.Address,'')
                                 And IsNull(D1.OverseasAddress,'') = IsNull(D2.OverseasAddress,'')";
                    if (HFD_Is_SameTel_other3.Value == "True")
                    {
                        if (HFD_Tel_other3.Value == "1")  //手機
                        {
                            strSql += " And ISNULL(D1.Cellular_Phone,'') = ISNULL(D2.Cellular_Phone,'') and D1.Cellular_Phone <>'' and D2.Cellular_Phone <>''";
                        }
                        else if (HFD_Tel_other3.Value == "2") //室內電話
                        {
                            strSql += " And ISNULL(D1.Tel_Office,'') = ISNULL(D2.Tel_Office,'') and D1.Tel_Office <>'' and D2.Tel_Office <>''";
                        }
                    }
                    strSql += @" Where D1.Donor_Id <> D2.Donor_Id 
                                 and D1.DeleteDate is null and D2.DeleteDate is null
                                 Group by D1.Donor_Id) 
                                 Order by Donor_Name,IsNull(M.ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address,  M.Donor_Id ";
                }
                else  //收據地址
                {
                    strSql += @" And IsNull(D1.Invoice_City,'') = IsNull(D2.Invoice_City,'') 
                                 And IsNull(D1.Invoice_Area,'') = IsNull(D2.Invoice_Area,'')
                                 And IsNull(D1.Invoice_Street,'') = IsNull(D2.Invoice_Street,'')
                                 --And IsNull(D1.Invoice_Address,'') = IsNull(D2.Invoice_Address,'')
                                 And IsNull(D1.Invoice_OverseasAddress,'') = IsNull(D2.Invoice_OverseasAddress,'')";
                    if (HFD_Is_SameTel_other3.Value == "True")
                    {
                        if (HFD_Tel_other3.Value == "1")  //手機
                        {
                            strSql += " And ISNULL(D1.Cellular_Phone,'') = ISNULL(D2.Cellular_Phone,'') and D1.Cellular_Phone <>'' and D2.Cellular_Phone <>''";
                        }
                        else if (HFD_Tel_other3.Value == "2") //室內電話
                        {
                            strSql += " And ISNULL(D1.Tel_Office,'') = ISNULL(D2.Tel_Office,'') and D1.Tel_Office <>'' and D2.Tel_Office <>''";
                        }
                    }
                    strSql += @" Where D1.Donor_Id <> D2.Donor_Id 
                                 and D1.DeleteDate is null and D2.DeleteDate is null
                                 Group by D1.Donor_Id)
                                 Order by Donor_Name,IsNull(M.Invoice_ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Invoice_Address,  M.Donor_Id ";
                }
            }
        }
        else
        {
            if (HFD_Is_SameAddress_other3.Value == "True")  //只選同地址 或同電話
            {
                if (HFD_Address_other3.Value == "1")  //通訊地址
                {
                    strSql += @" On IsNull(D1.City,'') = IsNull(D2.City,'') 
                                 And IsNull(D1.Area,'') = IsNull(D2.Area,'') 
                                 And IsNull(D1.Street,'') = IsNull(D2.Street,'') 
                                 --And IsNull(D1.Address,'') = IsNull(D2.Address,'')
                                 And IsNull(D1.OverseasAddress,'') = IsNull(D2.OverseasAddress,'')";
                    if (HFD_Is_SameTel_other3.Value == "True")
                    {
                        if (HFD_Tel_other3.Value == "1")  //手機
                        {
                            strSql += " And ISNULL(D1.Cellular_Phone,'') = ISNULL(D2.Cellular_Phone,'') and D1.Cellular_Phone <>'' and D2.Cellular_Phone <>''";
                        }
                        else if (HFD_Tel_other3.Value == "2") //室內電話
                        {
                            strSql += " And ISNULL(D1.Tel_Office,'') = ISNULL(D2.Tel_Office,'') and D1.Tel_Office <>'' and D2.Tel_Office <>''";
                        }
                    }
                    strSql += @" Where D1.Donor_Id <> D2.Donor_Id 
                                 And (IsNull(D1.Address,'') <> '' Or IsNull(D1.OverseasAddress,'') <> '') 
                                 and D1.DeleteDate is null and D2.DeleteDate is null
                                 Group by D1.Donor_Id)
                                 Order by IsNull(M.ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address, M.Donor_Id, Donor_Name ";
                }
                else  //收據地址
                {
                    strSql += @" On IsNull(D1.Invoice_City,'') = IsNull(D2.Invoice_City,'') 
                                 And IsNull(D1.Invoice_Area,'') = IsNull(D2.Invoice_Area,'') 
                                 And IsNull(D1.Invoice_Street,'') = IsNull(D2.Invoice_Street,'') 
                                 --And IsNull(D1.Invoice_Address,'') = IsNull(D2.Invoice_Address,'')
                                 And IsNull(D1.Invoice_OverseasAddress,'') = IsNull(D2.Invoice_OverseasAddress,'')";
                    if (HFD_Is_SameTel_other3.Value == "True")
                    {
                        if (HFD_Tel_other3.Value == "1")  //手機
                        {
                            strSql += " And ISNULL(D1.Cellular_Phone,'') = ISNULL(D2.Cellular_Phone,'') and D1.Cellular_Phone <>'' and D2.Cellular_Phone <>''";
                        }
                        else if (HFD_Tel_other3.Value == "2") //室內電話
                        {
                            strSql += " And ISNULL(D1.Tel_Office,'') = ISNULL(D2.Tel_Office,'') and D1.Tel_Office <>'' and D2.Tel_Office <>''";
                        }
                    }
                    strSql += @" Where D1.Donor_Id <> D2.Donor_Id 
                                 And (IsNull(D1.Invoice_Address,'') <> '' Or IsNull(D1.Invoice_OverseasAddress,'') <> '') 
                                 and D1.DeleteDate is null and D2.DeleteDate is null
                                 Group by D1.Donor_Id)
                                 Order by IsNull(M.Invoice_ZipCode,'')+IsNull(C1.Name,'') + IsNull(C2.Name,'') + Invoice_Address, M.Donor_Id, Donor_Name ";
                }
            }
        }


        //以下是串UNION的SQL語法
        /*strSql += @" Union select '999999999','總筆數：',CONVERT(VARCHAR,count(*)),'筆','','',''
                                    From dbo.DONOR M
                                    LEFT JOIN dbo.CODECITYNew C1
                                    ON M.City = C1.ZipCode
                                    LEFT JOIN dbo.CODECITYNew C2
                                    ON M.Area = C2.ZipCode
                                    Where Donor_Id In
                                    (Select D1.Donor_Id From dbo.DONOR D1
                                    Left Join dbo.DONOR D2 ";
        //搜尋條件
        if (HFD_Is_SameName_other3.Value == "True")
        {
            if (HFD_Is_SameAddress_other3.Value == "False")  //只選同姓名
            {
                strSql += @" ON D1.Donor_Name = D2.Donor_Name 
                         Where D1.Donor_Id <> D2.Donor_Id 
                         Group by D1.Donor_Id) ";
            }
            else  //同時選同姓名、同地址
            {
                if (HFD_Address_other3.Value == "1")  //通訊地址
                {
                    strSql += @" And IsNull(D1.Address,'') = IsNull(D2.Address,'') 
                                 Where D1.Donor_Id <> D2.Donor_Id 
                                 And (IsNull(D1.Address,'') <> '' Or IsNull(D1.OverseasAddress,'') <> '')  
                                 Group by D1.Donor_Id)  ";
                }
                else  //收據地址
                {
                    strSql += @" And IsNull(D1.Invoice_Address,'') = IsNull(D2.Invoice_Address,'') 
                                 Where D1.Donor_Id <> D2.Donor_Id 
                                 And (IsNull(D1.Invoice_Address,'') <> '' Or IsNull(D1.Invoice_OverseasAddress,'') <> '') 
                                 Group by D1.Donor_Id)  ";
                }
            }
        }
        else
        {
            if (HFD_Is_SameAddress_other3.Value == "True")  //只選同地址
            {
                if (HFD_Address_other3.Value == "1")  //通訊地址
                {
                    strSql += @" On IsNull(D1.City,'') = IsNull(D2.City,'') And IsNull(D1.Area,'') = IsNull(D2.Area,'') And IsNull(D1.Address,'') = IsNull(D2.Address,'') 
                                 Where D1.Donor_Id <> D2.Donor_Id 
                                 And (IsNull(D1.Address,'') <> '' Or IsNull(D1.OverseasAddress,'') <> '') 
                                 Group by D1.Donor_Id)  ";
                }
                else  //收據地址
                {
                    strSql += @" On IsNull(D1.Invoice_City,'') = IsNull(D2.Invoice_City,'') And IsNull(D1.Invoice_Area,'') = IsNull(D2.Invoice_Area,'') And IsNull(D1.Invoice_Address,'') = IsNull(D2.Invoice_Address,'') 
                                 Where D1.Donor_Id <> D2.Donor_Id 
                                 And (IsNull(D1.Invoice_Address,'') <> '' Or IsNull(D1.Invoice_OverseasAddress,'') <> '') 
                                 Group by D1.Donor_Id)  ";
                }
            }
        }*/
        return strSql;
    }
}