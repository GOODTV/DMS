using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
/*
 * 群組名稱設定
 * 資料庫: Contribute_Issue
 * List頁面: ContributeIssueList
 * 新增頁面: ContributeIssue_Add.aspx 
 * 詳細頁面: ContributeIssue_Detail.aspx
 * 修改頁面: ContributeIssue_Edit.aspx
 */
public partial class ContributeMgr_ContributeIssueList : BasePage
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
        Session["ProgID"] = "ContributeIssueList";
        //權控處理
        AuthrityControl();
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
        if (!IsPostBack)
        {
            LoadDropDownListData();
            LoadFormData();
        }
    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_Query", btnQuery);
        Authrity.CheckButtonRight("_Print", btnPrint);
        Authrity.CheckButtonRight("_Print", btnToxls);
        Authrity.CheckButtonRight("_AddNew", btnAdd);
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.Items.Insert(0, new ListItem("", ""));
        ddlDept.SelectedIndex = 1;

        //領取用途
        Util.FillDropDownList(ddlIssue_Purpose, Util.GetDataTable("CaseCode", "GroupName", "物品領取用途", "", ""), "CaseName", "CaseName", false);
        ddlIssue_Purpose.Items.Insert(0, new ListItem("", ""));
        ddlIssue_Purpose.SelectedIndex = 0;

        //經手人
        Util.FillDropDownList(ddlCreate_User, Util.GetDataTable("AdminUser", "1", "1", "", ""), "UserName", "UserName", false);
        ddlCreate_User.Items.Insert(0, new ListItem("", ""));
        ddlCreate_User.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = "select CI.Issue_Id, Issue_Processor as [領取人], CONVERT(VARCHAR(10) , CI.Issue_Date, 111 ) as [領取日期] \n";
        strSql += " , CI.Issue_Purpose as [領取用途], CI.Issue_Pre + CI.Issue_No as [領取編號], (Case When CI.Issue_Print='1' Then 'V' Else '' End) as [列印] \n";
        strSql += " , Case When CI.Export='Y' Then '作廢' Else Case When CI.Issue_Type='M' Then '手開' Else '' End End as [狀態], CI.Create_User as [經手人]\n";
        strSql += " WhereClause";
        strSql += " from Contribute_Issue CI\n";
        strSql += " where 1=1";
        
        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and Dept_Id = '" + ddlDept.SelectedValue + "'";
        }
        if (tbxIssue_Processor.Text.Trim() != "")
        {
            strSql += " and Issue_Processor like '%" + tbxIssue_Processor.Text.Trim() + "%'";
        }
        if (ddlIssue_Purpose.SelectedIndex != 0)
        {
            strSql += " and Issue_Purpose = '" + ddlIssue_Purpose.SelectedItem.Text + "'";
        }
        if (ddlCreate_User.SelectedIndex != 0)
        {
            strSql += " and Create_User = '" + ddlCreate_User.SelectedItem.Text + "'";
        }
        if (txtIssue_DateS.Text.Trim() != "" )
        {
            strSql += " and Issue_Date >= '" + Util.DateTime2String(txtIssue_DateS.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "' ";
        }
        if ( txtIssue_DateE.Text.Trim() != "")
        {
            strSql += " and Issue_Date <= '" + Util.DateTime2String(txtIssue_DateE.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "' ";
        }
        if (tbxIssue_NoS.Text.Trim() != "")
        {
            if (Util.IsNumeric(tbxIssue_NoS.Text.Trim()))
                strSql += " and Issue_No >= '" + tbxIssue_NoS.Text.Trim() + "' ";
            else
                strSql += " and Issue_Pre+Issue_No >= '" + tbxIssue_NoS.Text.Trim().PadRight(7, '0') + "' ";
        }
        if (tbxIssue_NoE.Text.Trim() != "")
        {
            if (Util.IsNumeric(tbxIssue_NoE.Text.Trim()))
                strSql += " and Issue_No <= '" + tbxIssue_NoE.Text.Trim() + "' ";
            else
                strSql += " and Issue_Pre+Issue_No <= '" + tbxIssue_NoE.Text.Trim().PadRight(7, '0') + "' ";
        }
        if (cbxIssue_Pre.Checked == true)
        {
            strSql += " and Issue_Type = 'M'";
        }
        if (cbxExport.Checked == true)
        {
            strSql += " and Export = 'Y'";
        }
        Session["strSql"] = strSql;
        strSql = strSql.Replace("WhereClause", "");
        dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.Keys.Add("Issue_Id");
            npoGridView.DisableColumn.Add("Issue_Id");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            npoGridView.EditLink = Util.RedirectByTime("ContributeIssue_Detail.aspx", "Issue_Id=");
            lblGridList.Text = npoGridView.Render();
        }

    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Response.Redirect("ContributeIssueList_Print_Excel.aspx");
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("ContributeIssue_Add.aspx"));
    }
}