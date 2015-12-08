using System;
using System.Collections.Generic;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonorGroupMgr_GroupClassQry : BasePage
{
    GroupAuthrity ga = null;

    #region NpoGridView 處理換頁相關程式碼
    Button btnNextPage, btnPreviousPage, btnGoPage;
    HiddenField HFD_CurrentPage;

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
        Session["ProgID"] = "GroupClass";
        //權控處理
        AuthrityControl();

        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
        if (!IsPostBack)
        {
            LoadDropDownListData();
            LoadFormData();
        }
    }
    //----------------------------------------------------------------------
    public void AuthrityControl()
    {
        ga = Authrity.GetGroupRight();
        if (Authrity.CheckPageRight(ga.Focus) == false)
        {
            return;
        }
        btnAdd.Visible = ga.AddNew;
        btnQuery.Visible = ga.Query;
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql = "";
        strSql = @"
                   select gc.uid, 
                   gc.GroupClassName as 群組類別,
                   gc.Supplement as 備註
                   from GroupClass gc
                   where gc.DeleteDate is null
                  ";

        if (txtGroupClassName.Text != "")
        {
            strSql += "and gc.GroupClassName like @GroupClassName\n";
        }
        if (txtSupplement.Text != "")
        {
            strSql += "and gc.Supplement like @Supplement\n";
        }

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("GroupClassName", "%" + txtGroupClassName.Text.Trim() + "%");
        dict.Add("Supplement", "%" + txtSupplement.Text.Trim() + "%");
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("uid");
        npoGridView.DisableColumn.Add("uid");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
        npoGridView.EditLink = Util.RedirectByTime("GroupClass_Edit.aspx", "GroupClassUID=");
        lblGridList.Text = npoGridView.Render();
    }
    //----------------------------------------------------------------------
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("GroupClass_Edit.aspx?Mode=ADD&", ""));
    }
    //----------------------------------------------------------------------
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    //----------------------------------------------------------------------
}
