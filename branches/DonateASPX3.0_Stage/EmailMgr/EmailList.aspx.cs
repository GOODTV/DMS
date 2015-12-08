using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
/*
 * 群組名稱設定
 * 資料庫: EmailMgr
 * 相關程式:
 * List頁面: EmailList.aspx
 * 新增頁面: Email_Add.aspx 
 * 修改頁面: Email_Edit.aspx
 */
public partial class EmailMgr_EmailList : BasePage
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
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "EmailList";
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
        Authrity.CheckButtonRight("_AddNew", btnAdd);
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //郵件類別
        Util.FillDropDownList(ddlEmailMgr_Type, Util.GetDataTable("CaseCode", "GroupName", "郵件類別", "", ""), "CaseName", "CaseName", false);
        ddlEmailMgr_Type.Items.Insert(0, new ListItem("", ""));
        ddlEmailMgr_Type.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = "select Ser_No, EmailMgr_Type as [郵件類別], EmailMgr_Subject as [郵件標題]\n";
        strSql += " from EmailMgr where 1=1 ";
        if (tbxEmailMgr_Subject.Text.Trim() != "")
        {
            strSql += " and EmailMgr_Subject like '%" + tbxEmailMgr_Subject.Text.Trim() + "%'";
        }
        if (ddlEmailMgr_Type.SelectedIndex != 0)
        {
            strSql += " and EmailMgr_Type ='" + ddlEmailMgr_Type.SelectedItem.Text + "'";
        }
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dt = NpoDB.GetDataTableS(strSql, dict);

        //Grid initial
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("Ser_No");
        npoGridView.DisableColumn.Add("Ser_No");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
        npoGridView.EditLink = Util.RedirectByTime("Email_Edit.aspx", "Ser_No=");
        lblGridList.Text = npoGridView.Render();

    }
    //---------------------------------------------------------------------------
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("Email_Add.aspx"));
    }
}