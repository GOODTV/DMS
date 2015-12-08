using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class DonateMgr_DonateSaleQry : BasePage
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
        Session["ProgID"] = "DonateSaleQry";
        //權控處理
        AuthrityControl();
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            LoadDropDownListData();
            // 2014/4/10 修正進入預設不載入資料
            //LoadFormData();
        }
    }
    //--------------------------------------------------------------------------
    
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_Query", btnQuery);
        Authrity.CheckButtonRight("_Query", btnHistoryQuery);
        Authrity.CheckButtonRight("_AddNew", btnAdd);
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.Items.Insert(0, new ListItem("", ""));
        ddlDept.SelectedIndex = 1;
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = "select Ser_No, Sale_Subject as [標題], CONVERT(VARCHAR(10) , Sale_BeginDate, 111 ) as [列印起日], CONVERT(VARCHAR(10) , Sale_EndDate, 111 ) as [列印迄日]\n";
        strSql += " from Donate_Sale where 1=1 ";
        if (ddlDept.SelectedIndex!= 0 )
        {
            strSql += " and Dept_Id='" + ddlDept.SelectedValue + "'";
        }
        if (tbxSale_Subject.Text.Trim() != "")
        {
            strSql += " and Sale_Subject like '%" + tbxSale_Subject.Text.Trim() + "%'";
        }
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dt = NpoDB.GetDataTableS(strSql, dict);

        //Grid initial
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("Ser_No");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
        npoGridView.DisableColumn.Add("Ser_No");
        npoGridView.EditLink = Util.RedirectByTime("DonateSale_Edit.aspx", "Ser_No=");
        lblGridList.Text = npoGridView.Render();

    }
    //---------------------------------------------------------------------------
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    //---------------------------
    //確認用途後再寫
    protected void btnHistoryQuery_Click(object sender, EventArgs e)
    {

    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("DonateSale_Add.aspx"));
    }
}