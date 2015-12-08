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

public partial class ContributeMgr_ContributeIssue_Detail : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            HFD_Uid.Value = Util.GetQueryString("Issue_Id");
            Form_DataBind();
        }
        Export();
    }
    //----------------------------------------------------------------------
    //帶入資料
    public void Form_DataBind()
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;

        //****變數設定****//
        uid = HFD_Uid.Value;

        //****設定查詢****//
        strSql = @" select * 
                    from Contribute_Issue CI left join Dept D on CI.Dept_Id = D.DeptID 
                    where CI.Issue_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);

        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("ContributeIssueList.aspx");

        DataRow dr = dt.Rows[0];

        //機構
        tbxDept.Text = dr["DeptShortName"].ToString().Trim();
        //領取人
        tbxIssue_Processor.Text = dr["Issue_Processor"].ToString().Trim();
        //領用日期
        tbxIssue_Date.Text = DateTime.Parse(dr["Issue_Date"].ToString()).ToString("yyyy/MM/dd");
        //領用用途
        tbxIssue_Purpose.Text = dr["Issue_Purpose"].ToString().Trim();
        //出貨單位
        tbxIssue_Org.Text = dr["Issue_Org"].ToString().Trim();
        //收據號碼
        tbxIssue_No.Text = dr["Issue_Pre"].ToString().Trim() + dr["Issue_No"].ToString().Trim();
        //捐款備註
        tbxComment.Text = dr["Issue_Comment"].ToString().Trim();

        //載入捐贈內容
        string strSql2, uid2;
        DataTable dt2;
        uid2 = HFD_Uid.Value;
        strSql2 = " select ROW_NUMBER() OVER(ORDER BY Ser_No) as [ROWID], ROW_NUMBER() OVER(ORDER BY Ser_No) as [序號], \n";
        strSql2 += " CID.Goods_Name as [物品名稱], CID.Goods_Qty as [數量] , CID.Goods_Unit as [單位], Goods_Comment as [備註], \n";
        strSql2 += " (case when Contribute_IsStock='Y' then 'V' else '' end) as [庫存品]\n";
        strSql2 += " from Contribute_IssueData CID right join Goods G on CID.Goods_Id = G.Goods_Id\n";
        strSql2 += " where Issue_Id='" + uid2 + "'";

        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        dt2 = NpoDB.GetDataTableS(strSql2, dict2);
        //Grid initial
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt2;
        npoGridView.DisableColumn.Add("ROWID");
        npoGridView.ShowPage = false;
        lblGridList.Text = npoGridView.Render();
    }
    public void Export()
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;
        //****變數設定****//
        uid = HFD_Uid.Value;
        //****設定查詢****//
        strSql = @" select Export
                    from Contribute_Issue CI
                    where CI.Issue_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);

        DataRow dr = dt.Rows[0];
        //判斷是否作廢
        if (dr["Export"].ToString() == "N")
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Export_N();", true);
        }
        if (dr["Export"].ToString() == "Y")
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Export_Y();", true);
        }
    }
    //----------------------------------------------------------------------
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("ContributeIssue_Edit.aspx", "Issue_Id=" + HFD_Uid.Value));
    }
    protected void btnPrint_Click(object sender, EventArgs e)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Contribute_Issue set ";
        strSql += " Issue_Print= @Issue_Print";
        strSql += " where Issue_Id = @Issue_Id";
        dict.Add("Issue_Print", "1");
        dict.Add("Issue_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    protected void btnExport_Click(object sender, EventArgs e)
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;
        //****變數設定****//
        uid = HFD_Uid.Value;
        //****設定查詢****//
        strSql = @" select Issue_Print 
                    from Contribute_Issue CI 
                    where CI.Issue_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Issue_Print = dr["Issue_Print"].ToString().Trim();
        //作廢單據前要先列印
        if (Issue_Print == "0")
        {
            ShowSysMsg("作廢單據前請先列印");
            return;
        }
        else
        {
            //****變數宣告****//
            Dictionary<string, object> dict = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql2 = " update Contribute_Issue set ";
            strSql2 += " Export= @Export";
            strSql2 += " where Issue_Id = @Issue_Id";
            dict.Add("Export", "Y");
            dict.Add("Issue_Id", HFD_Uid.Value);
            NpoDB.ExecuteSQLS(strSql2, dict);

            //找出Contribute_IssueData有幾筆資料
            Contribute_IssueData_Edit_Export();

            SetSysMsg("領取單已作廢！");
            Response.Redirect(Util.RedirectByTime("ContributeIssue_Detail.aspx", "Issue_Id=" + HFD_Uid.Value));
        }
    }
    protected void btn_ReExport_Click(object sender, EventArgs e)
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Contribute_Issue set ";
        strSql += " Export= @Export";
        strSql += " where Issue_Id = @Issue_Id";
        dict.Add("Export", "N");
        dict.Add("Issue_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);

        //找出Contribute_IssueData有幾筆資料
        Contribute_IssueData_Edit_ReExport();

        SetSysMsg("領取單已還原！");
        Response.Redirect(Util.RedirectByTime("ContributeIssue_Detail.aspx", "Issue_Id=" + HFD_Uid.Value));
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("ContributeIssueList.aspx"));
    }
    public void Contribute_IssueData_Edit_Export()
    {
        //****設定查詢****//
        string strSql = " select *  from Contribute_IssueData where Issue_Id='" + HFD_Uid.Value + "'";

        //****執行語法****//
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            DataRow dr = dt.Rows[i];
            string Goods_Id = dr["Goods_Id"].ToString();
            string Goods_Qty = dr["Goods_Qty"].ToString();

            //******修改Goods的資料******//
            Dictionary<string, object> dict_Goods = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Goods = " update Goods set ";
            strSql_Goods += "  Goods_Qty = Goods_Qty +'" + Goods_Qty + "'";
            strSql_Goods += " where Goods_Id = @Goods_Id";
            dict_Goods.Add("Goods_Id", Goods_Id);
            NpoDB.ExecuteSQLS(strSql_Goods, dict_Goods);
        }
    }
    public void Contribute_IssueData_Edit_ReExport()
    {
        //****設定查詢****//
        string strSql = " select *  from Contribute_IssueData where Issue_Id='" + HFD_Uid.Value + "'";

        //****執行語法****//
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            DataRow dr = dt.Rows[i];
            string Goods_Id = dr["Goods_Id"].ToString();
            string Goods_Qty = dr["Goods_Qty"].ToString();

            //******修改Goods的資料******//
            Dictionary<string, object> dict_Goods = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Goods = " update Goods set ";
            strSql_Goods += "  Goods_Qty = Goods_Qty -'" + Goods_Qty + "'";
            strSql_Goods += " where Goods_Id = @Goods_Id";
            dict_Goods.Add("Goods_Id", Goods_Id);
            NpoDB.ExecuteSQLS(strSql_Goods, dict_Goods);
        }
    }
}