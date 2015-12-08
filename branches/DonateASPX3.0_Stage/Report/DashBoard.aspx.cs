using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Report_DashBoard : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "DashBoard";
        //權控處理
        AuthrityControl();
        if (!IsPostBack)
        {
            lblyear1.Text = DateTime.Now.Year.ToString();
            lblyear2.Text = DateTime.Now.Year.ToString();
            lblyear3.Text = DateTime.Now.Year.ToString();
            lblDate_Today.Text = DateTime.Now.ToString("yyyy/MM/dd");
            lblThis_YearMonth.Text = DateTime.Now.Year.ToString() +"/"+ DateTime.Now.Month.ToString();
            lblThis_Year.Text = DateTime.Now.Year.ToString();
            lblLast_Year.Text = (DateTime.Now.Year -1).ToString(); //Year-1
            lblLast_YearMonth.Text = (DateTime.Now.Year - 1).ToString() + "/" + DateTime.Now.Month.ToString(); //Year-1

            String InfoSQL, SelectSQL, WhereSQL;
            DataTable dt;
            //天使總人數
            InfoSQL = "SELECT Replace(Convert(Varchar,Convert(money,COUNT(*)),1),'.00','') Angel_Count FROM dbo.DONOR WHERE IsMember <> 'Y'";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr = dt.Rows[0];
            lblAngel_Count.Text = dr["Angel_Count"].ToString();
            //讀者總人數
            InfoSQL = "SELECT Replace(Convert(Varchar,Convert(money,COUNT(*)),1),'.00','') Reader_Count FROM dbo.DONOR WHERE IsMember = 'Y'";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr1 = dt.Rows[0];
            lblReader_Count.Text = dr1["Reader_Count"].ToString();
            //今日累計奉獻金額,筆數,平均
            SelectSQL = @"SELECT Replace(Convert(Varchar,A.Count,1),'.00','') Count 
                                ,Replace(Convert(Varchar,A.SumAmt,1),'.00','') SumAmt 
                                ,Replace(Convert(Varchar,Case When A.Count > 2 Then Round((A.SumAmt - A.Max_Amt - A.Min_Amt) / (A.Count - 2),0) 			
                                Else 0 End,1),'.00','') DonateAmt_Avg 
  				                FROM (SELECT Convert(money,COUNT(*)) Count ,IsNull(SUM(Donate_Amt),0) SumAmt 
                                ,IsNull(Max(Donate_Amt),0) Max_Amt,IsNull(Min(Donate_Amt),0) Min_Amt 
                                FROM dbo.DONATE " ;
            WhereSQL = " WHERE Convert(Varchar(10),Donate_Date,111) = Convert(Varchar(10),GETDATE(),111) and Issue_Type <> 'D' ";
            InfoSQL = SelectSQL + WhereSQL + " ) A";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr2 = dt.Rows[0];
            lblToday_Count.Text = dr2["Count"].ToString();
            lblToday_SumAmt.Text = dr2["SumAmt"].ToString();
            lblToday_DonateAmt_Avg.Text = dr2["DonateAmt_Avg"].ToString();
            //今日線上累計奉獻筆數
            InfoSQL = @"select (select COUNT(*) as online_count from DONATE_IEPAY P left join DONATE_Web W on P.orderid=W.od_sob where status = '0' 
                          and Convert(Varchar(10),Donate_CreateDate,111)= Convert(Varchar(10),GETDATE(),111))+   
                          (select COUNT(*) as pledge_count from PLEDGE_Temp 				
                          where  Convert(Varchar(10),Create_Date,111) = Convert(Varchar(10),GETDATE(),111)) as count ";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr3 = dt.Rows[0];
            lblToday_Online_Count.Text = dr3["Count"].ToString();
            //本月累計奉獻金額,筆數,平均
            WhereSQL=" WHERE Year(Donate_Date) = Year(GETDATE()) AND MONTH(Donate_Date) = MONTH(GETDATE()) and Issue_Type <> 'D' ";
            InfoSQL = SelectSQL + WhereSQL + " ) A";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr4 = dt.Rows[0];
            lblMonth_Count.Text = dr4["Count"].ToString();
            lblMonth_SumAmt.Text = dr4["SumAmt"].ToString();
            lblMonth_DonateAmt_Avg.Text = dr4["DonateAmt_Avg"].ToString();
            //本月線上累計奉獻筆數
            InfoSQL = @"select (select COUNT(*) as online_count from DONATE_IEPAY P left join DONATE_Web W on P.orderid=W.od_sob where status = '0' 
                                and Year(Donate_CreateDate) = Year(GETDATE()) AND MONTH(Donate_CreateDate) = MONTH(GETDATE()))+ 
                                (select COUNT(*) as pledge_count from PLEDGE_Temp 			
                                where  Year(Create_Date) = Year(GETDATE()) AND MONTH(Create_Date) = MONTH(GETDATE())) as count ";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr5 = dt.Rows[0];
            lblMonth_Online_Count.Text = dr5["Count"].ToString();
            //本年度累計奉獻金額,筆數,平均
            WhereSQL = " WHERE Year(Donate_Date) = Year(GETDATE()) and Issue_Type <> 'D' ";
            InfoSQL = SelectSQL + WhereSQL + " ) A";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr6 = dt.Rows[0];
            lblYear_Count.Text = dr6["Count"].ToString();
            lblYear_SumAmt.Text = dr6["SumAmt"].ToString();
            lblYear_DonateAmt_Avg.Text = dr6["DonateAmt_Avg"].ToString();
            //本年度線上累計奉獻筆數
            InfoSQL = @"select (select COUNT(*) as online_count from DONATE_IEPAY P left join DONATE_Web W on P.orderid=W.od_sob where status = '0' 
                                and Year(Donate_CreateDate) = Year(GETDATE()))+    
                                (select COUNT(*) as pledge_count from PLEDGE_Temp 				
                                where  Year(Create_Date) = Year(GETDATE())) as count ";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr7 = dt.Rows[0];
            lblYear_Online_Count.Text = dr7["Count"].ToString();
            //上一年度奉獻總金額,筆數
            WhereSQL = " WHERE Year(Donate_Date) = Year(GETDATE())-1 and Issue_Type <> 'D' ";
            InfoSQL = SelectSQL + WhereSQL + " ) A";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr8 = dt.Rows[0];
            lblLastYear_Count.Text = dr8["Count"].ToString();
            lblLastYear_SumAmt.Text = dr8["SumAmt"].ToString();
            lblLastYear_DonateAmt_Avg.Text = dr8["DonateAmt_Avg"].ToString();
            //上一年度本月奉獻總金額,筆數
            WhereSQL = " WHERE Year(Donate_Date) = Year(GETDATE())-1 AND MONTH(Donate_Date) = MONTH(GETDATE()) and Issue_Type <> 'D' ";
            InfoSQL = SelectSQL + WhereSQL + " ) A";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr9 = dt.Rows[0];
            lblLastYearMonth_Count.Text = dr9["Count"].ToString();
            lblLastYearMonth_SumAmt.Text = dr9["SumAmt"].ToString();
            lblLastYearMonth_DonateAmt_Avg.Text = dr9["DonateAmt_Avg"].ToString();
            //歷年累計奉獻100萬以上天使筆數
            InfoSQL = @"SELECT Replace(Convert(Varchar,Convert(money,COUNT(Donor_Id)),1),'.00','') Over100_Count 
                                FROM dbo.DONOR WHERE Donor_Id IN 
                                (SELECT Donor_Id FROM dbo.DONATE WHERE Donate_Amt > 0 GROUP BY Donor_Id HAVING SUM(Donate_Amt) >= 1000000)";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr10 = dt.Rows[0];
            lblOver100_Count.Text = dr10["Over100_Count"].ToString();
            //歷年累計奉獻50萬以上天使筆數
            InfoSQL = @"SELECT Replace(Convert(Varchar,Convert(money,COUNT(Donor_Id)),1),'.00','') Over50_Count 
                                FROM dbo.DONOR WHERE Donor_Id IN 
                                (SELECT Donor_Id FROM dbo.DONATE WHERE Donate_Amt > 0 GROUP BY Donor_Id HAVING SUM(Donate_Amt) >= 500000)";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr11 = dt.Rows[0];
            lblOver50_Count.Text = dr11["Over50_Count"].ToString();
            //歷年累計奉獻30萬以上天使筆數
            InfoSQL = @"SELECT Replace(Convert(Varchar,Convert(money,COUNT(Donor_Id)),1),'.00','') Over30_Count 
                                FROM dbo.DONOR WHERE Donor_Id IN 
                                (SELECT Donor_Id FROM dbo.DONATE WHERE Donate_Amt > 0 GROUP BY Donor_Id HAVING SUM(Donate_Amt) >= 300000)";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr12 = dt.Rows[0];
            lblOver30_Count.Text = dr12["Over30_Count"].ToString();
            //歷年累計奉獻10萬以上天使筆數
            InfoSQL = @"SELECT Replace(Convert(Varchar,Convert(money,COUNT(Donor_Id)),1),'.00','') Over10_Count 
                                FROM dbo.DONOR WHERE Donor_Id IN 
                                (SELECT Donor_Id FROM dbo.DONATE WHERE Donate_Amt > 0 GROUP BY Donor_Id HAVING SUM(Donate_Amt) >= 100000)";
            dt = NpoDB.QueryGetTable(InfoSQL);
            DataRow dr13 = dt.Rows[0];
            lblOver10_Count.Text = dr13["Over10_Count"].ToString();
        }
    }
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
    }
}