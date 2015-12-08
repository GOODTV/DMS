using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class SysMgr_Log : BasePage
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
        Session["ProgID"] = "Log";
        //權控處理
        AuthrityControl();
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
        if (!IsPostBack)
        {
            LoadDropDownListData();
            //若是不要一開始帶入的話,刪除後必須用ajax帶回lbl
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
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.Items.Insert(0, new ListItem("", ""));
        ddlDept.SelectedIndex = 1;

        //帳號
        string strSql = "Select UserID, UserID + ' ' + UserName as UserName From AdminUser";
        Util.FillDropDownList(ddlUserID, strSql, "UserName", "UserID", false);
        ddlUserID.Items.Insert(0, new ListItem("", ""));
        ddlUserID.SelectedIndex = 0;

        //Log類別
        ddlLog_Type.Items.Insert(0, new ListItem("", ""));
        ddlLog_Type.Items.Insert(1, new ListItem("系統登入", "系統登入"));
        ddlLog_Type.Items.Insert(2, new ListItem("管理密碼", "管理密碼"));
        ddlLog_Type.Items.Insert(3, new ListItem("使用者異動", "使用者異動"));
        ddlLog_Type.Items.Insert(4, new ListItem("刪除資料", "刪除資料"));
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = @"select L.uid, CONVERT(VARCHAR, L.LogTime, 120 ) as [LOG時間] ,D.DeptShortName as [機構], A.UserID as [帳號], 
                       L.Log_Type as [LOG內容], A.UserName as [帳號名稱], L.LogIP as [登入IP]
                  from Log L 
                  left join AdminUser A on L.UserID = A.UserID 
                  left join Dept D on A.DeptID = D.DeptID
                  where 1=1";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        //LOG日期
        if (txtDateS.Text.Trim() != "")
        {
            strSql += " and L.LogTime >= '" + Util.DateTime2String(txtDateS.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "' ";
        }
        if (txtDateE.Text.Trim() != "")
        {
            strSql += " and L.LogTime <= '" + Util.DateTime2String(txtDateE.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "' ";
        }
        //機構
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and D.DeptId = '" + ddlDept.SelectedValue + "'";
        }
        //LOG類別
        if (ddlLog_Type.SelectedIndex != 0)
        {
            strSql += " and L.Log_Type = '" + ddlLog_Type.SelectedValue + "'";
        }
        //帳號
        if (ddlUserID.SelectedIndex != 0)
        {
            strSql += " and L.UserID = '" + ddlUserID.SelectedValue + "'";
        }
        strSql += " order by LogTime desc";
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
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            npoGridView.DisableColumn.Add("uid");
            //-------------------------------------------------------------------------
            NPOGridViewColumn col = new NPOGridViewColumn("LOG時間");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("LOG時間");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("機構");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("機構");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("帳號");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("帳號");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("帳號名稱");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("帳號名稱");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("LOG內容");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("LOG內容");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("登入IP");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("登入IP");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("刪除");
            //button
            //col.ColumnType = NPOColumnType.Button;
            //col.ControlText.Add("刪除");

            //link+icon
            col.ColumnType = NPOColumnType.Link;
            col.LinkText = "";
            col.IconPath = "../images/x1.gif";
            col.ShowConfirmDialog = true;
            col.ConfirmDialogMsg = "是否確定要刪除?";
            col.ControlKeyColumn.Add("uid");
            col.Link = "Log_Delete.aspx?uid=";
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            lblGridList.Text = npoGridView.Render();
        }
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
}