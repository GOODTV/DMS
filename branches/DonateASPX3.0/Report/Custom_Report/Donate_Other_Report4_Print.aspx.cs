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

public partial class Report_Custom_Report_Donate_Other_Report4_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_IsMember_other4.Value = Util.GetQueryString("IsMember");
            HFD_DonateDateS_other4.Value = Util.GetQueryString("DonateDateS");
            HFD_DonateDateE_other4.Value = Util.GetQueryString("DonateDateE");
            HFD_Is_Abroad_other4.Value = Util.GetQueryString("Is_Abroad");
            HFD_Is_ErrAddress_other4.Value = Util.GetQueryString("Is_ErrAddress");
            HFD_Sex_other4.Value = Util.GetQueryString("Sex");
            PrintReport();
        }
    }
    private void PrintReport()
    {
        string strSql = Sql();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
            return;
        }

        DataTable dtRet = CaseUtil.Donate_Other_Report4_Print(dt);

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
            //else if (iCtrl == 1)
            //{
            //    //cell.Width = "40mm";
            //    cell.Style.Add("text-align", "center");
            //}
            else if (iCtrl == 17)
            {
                //cell.Width = "30mm";
                cell.Style.Add("text-align", "center");
                cell.Style.Add("border-right", ".5pt solid windowtext");
            }
            else
            {
                //cell.Width = "40mm";
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
                else if (Col == 17)
                {
                    cell.Style.Add("border-right", ".5pt solid windowtext");
                    cell.Style.Add("text-align", "center");
                }
                else if (Col == 5)
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

        GridList.Text = "捐款人/讀者資料明細<br/>" + "<span  style=' font-size: 9pt; font-family: 新細明體'>統計日期：" + DateTime.Now.ToString() + "</span><br><span  style=' font-size: 9pt; font-family: 新細明體'>依捐款人編號排序</span><br>" + htw.InnerWriter.ToString();
        Response.Write("<script language=javascript>window.print();</script>");
    }
    private string Sql()
    {
        string strSql;
        if (HFD_IsMember_other4.Value == "Y")  //讀者
        {
            strSql = @"SELECT M.Donor_Id AS '編號'   
                        ,M.Donor_Name AS '姓名' ,IsNUll(M.Sex,'') AS '性別' ,DATEDIFF(year ,Birthday,getdate()) as '年齡' ,M.Donor_Type AS '身分別'
                        ,Case When M.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'')  
		                        else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
                        ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
				                            (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
				                            Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
		                                Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
				                            Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話'  
                                ,IsNull(Cellular_Phone,'') AS '手機' ,IsNull(Email,'') AS 'Email'
                        ,'' AS '奉獻方式' ,'' AS '累計奉獻金額' ,'' AS '奉獻次數'  
                        ,CONVERT(VARCHAR, Begin_DonateDate, 111) AS '首捐日期' ,CONVERT(VARCHAR, Last_DonateDate, 111) AS '末捐日期' 
                        ,Case When IsSendNews='Y' Then 'V' Else '' End AS '月刊寄否',IsNull(IsSendNewsNum,0) AS '月刊份數'  
                        ,Case When IsContact='N' Then 'V' Else '' End '不主動聯絡'
                        ,Case When IsErrAddress='Y' Then 'V' Else '' End '是否錯址'  
                        FROM dbo.DONOR M  
                        LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
                                    LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode
                                    Where M.Donor_Type like '%讀者%' --IsMember = 'Y' 
			                        and DeleteDate is NULL ";
            if (HFD_Is_Abroad_other4.Value == "True")
            {
                strSql += " And IsAbroad='N' ";
            }
            if (HFD_Is_ErrAddress_other4.Value == "True")
            {
                strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
            }
            if (HFD_Sex_other4.Value == "True")
            {
                strSql += " And IsNull(Sex,'') <> '歿' ";
            }
        }
        else  //捐款人
        {
            strSql = @"SELECT D.Donor_Id AS '編號'   
                        ,M.Donor_Name AS '姓名' ,IsNUll(M.Sex,'') AS '性別' ,DATEDIFF(year ,Birthday,getdate()) as '年齡' ,M.Donor_Type AS '身分別'
                        ,Case When M.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'')  
		                        else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
                        ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
				                            (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
				                            Else Tel_Office + ' <br>Ext.' + Tel_Office_Ext End)  
		                                Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
				                            Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' <br>Ext.' + Tel_Office_Ext End) End '電話'  
                                ,IsNull(Cellular_Phone,'') AS '手機' ,IsNull(Email,'') AS 'Email'
                        ,(SELECT DISTINCT Donate_Payment + '.' FROM dbo.DONATE as D1 WHERE M.Donor_Id=D1.Donor_Id FOR XML PATH('') ) AS '奉獻方式'
                        ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Sum_Amt),1),'.00','') AS '累計奉獻金額' ,D.Times AS '奉獻次數'  
                        ,CONVERT(VARCHAR, Begin_DonateDate, 111) AS '首捐日期' ,CONVERT(VARCHAR, Last_DonateDate, 111) AS '末捐日期' 
                        ,Case When IsSendNews='Y' Then 'V' Else '' End AS '月刊寄否',IsNull(IsSendNewsNum,0) AS '月刊份數'  
                        ,Case When IsContact='N' Then 'V' Else '' End '不主動聯絡'
                        ,Case When IsErrAddress='Y' Then 'V' Else '' End '是否錯址'  
                        FROM dbo.DONOR M  
                        INNER JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
			                        FROM DONATE 
                                    WHERE ISNULL(Issue_Type,'') != 'D' ";

            //搜尋條件
            if (HFD_DonateDateS_other4.Value != "")
            {
                strSql += " And Donate_Date>='" + HFD_DonateDateS_other4.Value + "'";
            }
            if (HFD_DonateDateS_other4.Value != "")
            {
                strSql += " And Donate_Date<='" + HFD_DonateDateE_other4.Value + "'";
            }
            strSql += @"GROUP BY Donor_Id) D
			ON M.Donor_Id = D.Donor_Id 
            LEFT JOIN dbo.CODECITYNew C1 ON M.City = C1.ZipCode
            LEFT JOIN dbo.CODECITYNew C2 ON M.Area = C2.ZipCode
            Where DeleteDate is NULL  ";
            if (HFD_Is_Abroad_other4.Value == "True")
            {
                strSql += " And IsAbroad='N' ";
            }
            if (HFD_Is_ErrAddress_other4.Value == "True")
            {
                strSql += " And (IsErrAddress='N' or IsErrAddress is NULL) ";
            }
            if (HFD_Sex_other4.Value == "True")
            {
                strSql += " And IsNull(Sex,'') <> '歿' ";
            }
        }

        //以下是串UNION的SQL語法
//        strSql += @" Union 
//					 select '99999',convert(nvarchar, count(MAIN.Donor_Id)),'','','',Sum (Cast (IsSendNewsNum as int)),''
//					 FROM DONOR MAIN  
//					 WHERE IsSendNews='Y' and DeleteDate is NULL ";
//        //搜尋條件
//        if (HFD_IsMember_other4.Value == "Y")
//        {
//            strSql += " And IsMember='Y' ";
//        }
//        else
//        {
//            strSql += " And IsMember='N' ";
//        }
//        if (HFD_DonateDateS_other4.Value != "")
//        {
//            strSql = " And Donate_Date>='" + HFD_DonateDateS_other4.Value + "'";
//        }
//        if (HFD_DonateDateS_other4.Value != "")
//        {
//            strSql += " And Donate_Date<='" + HFD_DonateDateS_other4.Value + "'";
//        }
        strSql += " Order by M.Donor_Id ";
        return strSql;
    }
}