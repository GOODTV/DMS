using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
/*
 * 群組名稱設定
 * 資料庫: Pledge
 * List頁面: PledgeInfo
 * 新增頁面: Pledge_Add.aspx 
 * 修改頁面: Pledge_Edit.aspx
 */
public partial class DonateMgr_PledgeQry : BasePage
{
    #region NpoGridView 處理換頁相關程式碼
    Button btnNextPage, btnPreviousPage, btnGoPage;
    HiddenField HFD_CurrentPage, HFD_CurrentQuerye;

    override protected void OnInit(EventArgs e)
    {
        CreatePageControl();
        base.OnInit(e);
    }
    private void CreatePageControl()
    {
        // Create dynamic controls here.
        btnNextPage = new Button();
        btnNextPage.ID = "btnNextPage";
        Form1.Controls.Add(btnNextPage);
        btnNextPage.Click += new System.EventHandler(btnNextPage_Click);

        btnPreviousPage = new Button();
        btnPreviousPage.ID = "btnPreviousPage";
        Form1.Controls.Add(btnPreviousPage);
        btnPreviousPage.Click += new System.EventHandler(btnPreviousPage_Click);

        btnGoPage = new Button();
        btnGoPage.ID = "btnGoPage";
        Form1.Controls.Add(btnGoPage);
        btnGoPage.Click += new System.EventHandler(btnGoPage_Click);

        HFD_CurrentPage = new HiddenField();
        HFD_CurrentPage.Value = "1";
        HFD_CurrentPage.ID = "HFD_CurrentPage";
        Form1.Controls.Add(HFD_CurrentPage);

        HFD_CurrentQuerye = new HiddenField();
        HFD_CurrentQuerye.Value = "Query";
        HFD_CurrentQuerye.ID = "HFHFD_CurrentQuerye";
        Form1.Controls.Add(HFD_CurrentQuerye);
    }
    protected void btnPreviousPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.MinusStringNumber(HFD_CurrentPage.Value);
        LoadFormData();
    }
    protected void btnNextPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.AddStringNumber(HFD_CurrentPage.Value);
        LoadFormData();
    }
    protected void btnGoPage_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    #endregion NpoGridView 處理換頁相關程式碼
    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "PledgeInfo";
        //權控處理
        AuthrityControl();
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            Session["cType"] = "";
            LoadDropDownListData();
            // 2014/4/10 修正進入預計不載入資料
            //LoadFormData();
        }

        // 2014/4/10 取得焦點
        tbxDonor_Id.Focus();

    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_AddNew", btnAdd);
        Authrity.CheckButtonRight("_Query", btnQuery);
        Authrity.CheckButtonRight("_Query", btnPrint);
        Authrity.CheckButtonRight("_Query", btnToxls);
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.Items.Insert(0, new ListItem("", ""));
        ddlDept.SelectedIndex = 1;

        //捐款授權到期日
        Util.FillDropDownList(ddlYear_Donate_ToDate, Int32.Parse(DateTime.Now.Year.ToString()), Int32.Parse(DateTime.Now.Year.ToString()) + 10, true, 0);
        Util.FillDropDownList(ddlMonth_Donate_ToDate, 1, 12, true, 0);

        //信用卡有效月年
        Util.FillDropDownList(ddlMonth_Valid_Date, 1, 12, true, 0);
        Util.FillDropDownList(ddlYear_Valid_Date, Int32.Parse(DateTime.Now.Year.ToString()), Int32.Parse(DateTime.Now.Year.ToString()) + 10, true, 0);
        
        //20140415 Modify by GoodTV Tanya:授權方式改從資料庫抓代碼清單
        //授權方式
        Util.FillDropDownList(ddlDonate_Payment, Util.GetDataTable("CaseCode", "GroupName", "授權方式", "", ""), "CaseName", "CaseName", false);
        ddlDonate_Payment.Items.Insert(0, new ListItem("", ""));
        //ddlDonate_Payment.Items.Insert(1, new ListItem("信用卡授權書", "信用卡授權書"));
        //ddlDonate_Payment.Items.Insert(2, new ListItem("帳戶轉帳授權書", "帳戶轉帳授權書"));
        ddlDonate_Payment.SelectedIndex = 0;

        //指定用途
        Util.FillDropDownList(ddlDonate_Purpose, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseID", false);
        ddlDonate_Purpose.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Purpose.SelectedIndex = 0;

        //轉帳週期
        ddlDonate_Period.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Period.Items.Insert(1, new ListItem("單筆", "單筆"));
        ddlDonate_Period.Items.Insert(2, new ListItem("月繳", "月繳"));
        ddlDonate_Period.Items.Insert(3, new ListItem("季繳", "季繳"));
        ddlDonate_Period.Items.Insert(4, new ListItem("半年繳", "半年繳"));
        ddlDonate_Period.Items.Insert(5, new ListItem("年繳", "年繳"));
        ddlDonate_Period.SelectedIndex = 0;
        //狀態
        ddlStatus.Items.Insert(0, new ListItem("授權中", "授權中"));
        ddlStatus.Items.Insert(1, new ListItem("停止", "停止"));
        ddlStatus.SelectedIndex = 0;

        //經手人
        Util.FillDropDownList(ddlCreate_User, Util.GetDataTable("AdminUser", "1", "1", "", ""), "UserName", "UserName", false);
        ddlCreate_User.Items.Insert(0, new ListItem("", ""));
        ddlCreate_User.SelectedIndex = 0;
    }
    public void LoadFormData()
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = Sql();
        strSql += SqlCondition();
        if (tbxPledge_Id.Text.Trim() != "")
        {
            // 2014/5/26 修改不用 SQL variable
            //strSql += " and Pledge_Id=@Pledge_Id";
            //dict.Add("Pledge_Id", tbxPledge_Id.Text.Trim());
            strSql += " and Pledge_Id=" + tbxPledge_Id.Text.Trim();
        }
        if (tbxDonor_Id.Text.Trim() != "")
        {
            // 2014/5/26 修改不用 SQL variable
            //strSql += " and P.Donor_Id=@Donor_Id";
            //dict.Add("Donor_Id", tbxDonor_Id.Text.Trim());
            strSql += " and P.Donor_Id=" + tbxDonor_Id.Text.Trim();
        }
        strSql += " order by Pledge_Id desc";

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
            npoGridView.Keys.Add("授權編號");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            Session["cType"] = "PledgeInfo";
            npoGridView.EditLink = Util.RedirectByTime("Pledge_Edit.aspx", "Pledge_Id=");
            lblGridList.Text = npoGridView.Render();
            lblAmt.Text = Donate_Amt();
        }
        Session["strSql"] = strSql;
    }
    //---------------------------------------------------------------------------
    protected void btnAuthorize_Click(object sender, EventArgs e)
    {
        Response.Redirect("PledgeTodate.aspx");
    }
    protected void btnCreditCard_Click(object sender, EventArgs e)
    {
        Response.Redirect("PledgeValid.aspx");
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Session["cType"] = "PledgeInfo";
        Response.Redirect("PledgeInfoQry.aspx");
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        string strSql = Sql();
        strSql += SqlCondition();
        strSql += " order by Pledge_Id desc";
        if (tbxPledge_Id.Text.Trim() != "")
        {
            strSql += " and Pledge_Id='" + tbxPledge_Id.Text.Trim() + "'";
        }
        if (tbxDonor_Id.Text.Trim() != "")
        {
            strSql += " and P.Donor_Id='" + tbxDonor_Id.Text.Trim() + "'";
        }
        Session["strSql"] = strSql;

        Response.Redirect("PledgeInfo_Print_Excel.aspx");
    }
    private string Sql()
    {
        //20140428 Modify by GoodTV:修正部份卡號/帳號因其他欄位Null造成無法顯示的問題
        string strSql = "";
        strSql = "select P.Pledge_Id as [授權編號], D.Donor_Name as [捐款人], D.Donor_Id as [捐款人編號], P.Donate_Payment as [授權方式],\n";
        strSql += " P.Donate_Purpose as [指定用途], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,P.Donate_Amt),1),'.00','') as [扣款金額],\n";
        strSql += " CONVERT(NVARCHAR, P.Donate_FromDate, 111) as [授權起日], CONVERT(NVARCHAR, P.Donate_ToDate, 111) as [授權迄日],\n";
        strSql += " P.Donate_Period as [轉帳週期], CONVERT(NVARCHAR, P.Next_DonateDate, 111) as [下次扣款日期],\n";
        strSql += " Card_Bank + '-' + Card_Type as [發卡銀行-卡別], (case when Account_No<>'' or Post_SavingsNo<>'' or  P_RCLNO<>'' then IsNull(Account_No,'') + IsNull(Post_SavingsNo,'') + IsNull(P_RCLNO,'') else '' end) as [卡號/帳號],\n";
        strSql += " substring(P.Valid_Date,1,2) +'/'+ substring(P.Valid_Date,3,2) as [信用卡有效月年],\n";
        strSql += " P.Status as [授權狀態], P.Create_User as [經手人]\n";
        strSql += " from Pledge P left join Donor D on P.Donor_Id = D.Donor_Id where 1=1";
        return strSql;
    }
    private string SqlCondition()
    {
        string strSql="";
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and P.Dept_Id='" + ddlDept.SelectedValue + "'";
        }
        if (ddlYear_Donate_ToDate.SelectedIndex != 0)
        {
            strSql += " and Year(P.Donate_ToDate)='" + ddlYear_Donate_ToDate.SelectedItem.Text + "'";
        }
        if (ddlMonth_Donate_ToDate.SelectedIndex != 0)
        {
            strSql += " and Month(P.Donate_ToDate)='" + ddlMonth_Donate_ToDate.SelectedItem.Text + "'";
        }
        if (ddlDonate_Payment.SelectedIndex != 0)
        {
            strSql += " and P.Donate_Payment='" + ddlDonate_Payment.SelectedItem.Text + "'";
        }
        if (ddlDonate_Purpose.SelectedIndex != 0)
        {
            strSql += " and P.Donate_Purpose='" + ddlDonate_Purpose.SelectedItem.Text + "'";
        }
        if (tbxAccount_No.Text.Trim() != "")
        {
            strSql += " and (Account_No like '%" + tbxAccount_No.Text.Trim() + "%' or Post_SavingsNo like '%" + tbxAccount_No.Text.Trim() + "%' or P_RCLNO like '%" + tbxAccount_No.Text.Trim() + "%')";
        }
        if (ddlMonth_Valid_Date.SelectedIndex != 0)
        {
            strSql += " and substring(P.Valid_Date,1,2) ='" + ddlMonth_Valid_Date.SelectedItem.Text.PadLeft(2, '0') + "'";
        }
        if (ddlYear_Valid_Date.SelectedIndex != 0)
        {
            strSql += " and substring(P.Valid_Date,3,2) ='" + ddlYear_Valid_Date.SelectedItem.Text.Substring(2, 2) + "'";
        }
        if (txtDonateDate.Text.Trim() != "")
        {
            strSql += " and P.Donate_ToDate >='" + txtDonateDate.Text.Trim() + "'";
        }
        if (ddlDonate_Period.SelectedIndex != 0)
        {
            strSql += " and P.Donate_Period='" + ddlDonate_Period.SelectedItem.Text + "'";
        }
        strSql += " and P.Status='" + ddlStatus.SelectedItem.Text + "'";
        if (ddlCreate_User.SelectedIndex != 0)
        {
            strSql += " and P.Create_User='" + ddlCreate_User.SelectedItem.Text + "'";
        }
        //20140523 新增by Ian 持卡人向度
        if (tbxCard_Owner.Text.Trim() != "")
        {
            strSql += " and P.Card_Owner like '%" + tbxCard_Owner.Text.Trim() +"%' ";
        }
        return strSql;
    }
    private string Donate_Amt()
    {
        string strSql = @"select REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when Status = '授權中' then Donate_Amt else 0 end)),1),'.00','') as Donate_Amt
                          from Pledge P left join Donor D on P.Donor_Id = D.Donor_Id where 1=1 ";
        strSql += SqlCondition();
        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (tbxPledge_Id.Text.Trim() != "")
        {
            // 2014/5/26 修改不用 SQL variable
            //strSql += " and Pledge_Id=@Pledge_Id";
            //dict.Add("Pledge_Id", tbxPledge_Id.Text.Trim());
            strSql += " and Pledge_Id=" + tbxPledge_Id.Text.Trim();
        }
        if (tbxDonor_Id.Text.Trim() != "")
        {
            // 2014/5/26 修改不用 SQL variable
            //strSql += " and P.Donor_Id=@Donor_Id";
            //dict.Add("Donor_Id", tbxDonor_Id.Text.Trim());
            strSql += " and P.Donor_Id=" + tbxDonor_Id.Text.Trim();
        }
        
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);
        DataRow dr = dt.Rows[0];
        string Donate_Amt = dr["Donate_Amt"].ToString();
        return Donate_Amt;
    }
}