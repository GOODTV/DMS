using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonateMgr_WebPledgeList : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "WebPledgeList";
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            LoadFormData();
        }
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql = Sql();

        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
            lblAmt.Text = "0";
        }
        else
        {
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            //npoGridView.Keys.Add("授權編號");
            //-------------------------------------------------------------------------
            NPOGridViewColumn col = new NPOGridViewColumn("選擇");
            col.ColumnType = NPOColumnType.Button;
            col.ShowConfirmDialog = true;
            col.ConfirmDialogMsg = "確定是否授權轉帳？";
            col.Link = "WebPledge_Transfer.aspx?Pledge_Id=";

            col.ControlValue.Add("");
            col.ControlKeyColumn.Add("Pledge_Id");
            col.ColumnName.Add("Pledge_Id");
            col.ControlText.Add("確認授權");
            col.ControlName.Add("PledgeId");
            col.ControlId.Add("Button");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("授權期間是否重複");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("授權期間是否重複");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("授權編號");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("授權編號");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("捐款人編號");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("捐款人編號");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("捐款人");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("捐款人");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("授權方式");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("授權方式");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("指定用途");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("指定用途");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("扣款金額");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("扣款金額");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("授權起日");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("授權起日");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("授權迄日");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("授權迄日");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("轉帳週期");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("轉帳週期");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            /*col = new NPOGridViewColumn("下次扣款日期");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("下次扣款日期");
            npoGridView.Columns.Add(col);*/
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("信用卡有效月年");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("信用卡有效月年");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("授權狀態");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("授權狀態");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("問卷");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("問卷");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("建檔日期");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("建檔日期");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            npoGridView.Keys.Add("捐款人編號");
            npoGridView.EditLink = Util.RedirectByTime("../DonorMgr/DonorInfo_Edit.aspx", "Donor_Id=");
            lblGridList.Text = npoGridView.Render();
            lblAmt.Text = Donate_Amt();
        }
    }
    private string Sql()
    {
        string strSql = "";
        strSql = @"Select Pledge_Id,
                          Case when (SELECT case when CONVERT(VARCHAR(12),MIN(P.Donate_FromDate),111)<= CONVERT(VarChar,PT.donate_fromdate,111) 
                            and CONVERT(VARCHAR(12),MAX(P.Donate_ToDate),111) >= CONVERT(VarChar,PT.donate_todate,111) then '1' else '0' end  count
                            FROM dbo.PLEDGE P
                            WHERE P.Status='授權中' AND P.Donor_Id=PT.Donor_Id)='1' then '重複' else ' ' end as [授權期間是否重複],
                          Pledge_Id as [授權編號],PT.Donor_Id as [捐款人編號],D.Donor_Name as [捐款人],
                          Donate_Payment as [授權方式],Donate_Purpose as [指定用途],
                          REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End)),1),'.00','')  as [扣款金額],
                          CONVERT(VarChar,donate_fromdate,111) as [授權起日],CONVERT(VarChar,donate_todate,111) as [授權迄日],
                          donate_period as [轉帳週期],CONVERT(VarChar,Next_DonateDate,111) as [下次扣款日期],
                          (Case When valid_date<>'' Then Left(valid_date,2)+'/'+Right(Valid_date,2) Else '' End) as [信用卡有效月年],status  as [授權狀態],
                          Case when(select SerNo from Donate_OnlineQuestion where DonateWay='線上定期定額' and Donor_Id=PT.Donor_Id and CONVERT(VarChar,Create_Date,111)=CONVERT(VarChar,PT.Create_Date,111)) !='' then 'V' else '' end as [問卷], 
                          CONVERT(VarChar,PT.Create_Date,111) as [建檔日期]
                   From PLEDGE_Temp PT Join DONOR D On PT.Donor_Id=D.Donor_Id 
                   Where Status = '授權中'
                   Order By Pledge_Id Desc";
        return strSql;
    }
    private string Donate_Amt()
    {
        string strSql = @"select REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when Status = '授權中' then Donate_Amt else 0 end)),1),'.00','') as Donate_Amt
                          from PLEDGE_Temp PT Join DONOR D On PT.Donor_Id=D.Donor_Id  where 1=1 ";
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Donate_Amt = dr["Donate_Amt"].ToString();
        return Donate_Amt;
    }
}