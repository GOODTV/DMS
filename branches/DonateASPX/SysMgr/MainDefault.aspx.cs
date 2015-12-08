using System;
using System.Web.UI;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Web.UI.WebControls;

public partial class MainDefault : BasePage
{
    #region NpoGridView 處理換頁相關程式碼
    Button btnNextPage;
    Button btnPreviousPage;
    Button btnGoPage;
    HiddenField HFD_CurrentPage;

    Button btnNextPage2;
    Button btnPreviousPage2;
    Button btnGoPage2;
    HiddenField HFD_CurrentPage2;

    Button btnNextPage3;
    Button btnPreviousPage3;
    Button btnGoPage3;
    HiddenField HFD_CurrentPage3;

    Button btnNextPage4;
    Button btnPreviousPage4;
    Button btnGoPage4;
    HiddenField HFD_CurrentPage4;

    Button btnNextPage5;
    Button btnPreviousPage5;
    Button btnGoPage5;
    HiddenField HFD_CurrentPage5;

    override protected void OnInit(EventArgs e)
    {
        CreatePageControl();
        base.OnInit(e);
    }
    
    private void CreatePageControl()
    {
        //Create dynamic controls here.
        btnNextPage = new Button();
        btnNextPage.ID = "btnNextPage";
        form1.Controls.Add(btnNextPage);
        btnNextPage.Click += new System.EventHandler(btnNextPage_Click);

        btnPreviousPage = new Button();
        btnPreviousPage.ID = "btnPreviousPage";
        form1.Controls.Add(btnPreviousPage);
        btnPreviousPage.Click += new System.EventHandler(btnPreviousPage_Click);

        btnGoPage = new Button();
        btnGoPage.ID = "btnGoPage";
        form1.Controls.Add(btnGoPage);
        btnGoPage.Click += new System.EventHandler(btnGoPage_Click);

        HFD_CurrentPage = new HiddenField();
        HFD_CurrentPage.Value = "1";
        HFD_CurrentPage.ID = "HFD_CurrentPage";
        form1.Controls.Add(HFD_CurrentPage);

        //第2個 GridView
        btnNextPage2 = new Button();
        btnNextPage2.ID = "btnNextPage2";
        form1.Controls.Add(btnNextPage2);
        btnNextPage2.Click += new System.EventHandler(btnNextPage2_Click);

        btnPreviousPage2 = new Button();
        btnPreviousPage2.ID = "btnPreviousPage2";
        form1.Controls.Add(btnPreviousPage2);
        btnPreviousPage2.Click += new System.EventHandler(btnPreviousPage2_Click);

        btnGoPage2 = new Button();
        btnGoPage2.ID = "btnGoPage2";
        form1.Controls.Add(btnGoPage2);
        btnGoPage2.Click += new System.EventHandler(btnGoPage2_Click);

        HFD_CurrentPage2 = new HiddenField();
        HFD_CurrentPage2.Value = "1";
        HFD_CurrentPage2.ID = "HFD_CurrentPage2";
        form1.Controls.Add(HFD_CurrentPage2);

        //第3個 GridView
        btnNextPage3 = new Button();
        btnNextPage3.ID = "btnNextPage3";
        form1.Controls.Add(btnNextPage3);
        btnNextPage3.Click += new System.EventHandler(btnNextPage3_Click);

        btnPreviousPage3 = new Button();
        btnPreviousPage3.ID = "btnPreviousPage3";
        form1.Controls.Add(btnPreviousPage3);
        btnPreviousPage3.Click += new System.EventHandler(btnPreviousPage3_Click);

        btnGoPage3 = new Button();
        btnGoPage3.ID = "btnGoPage3";
        form1.Controls.Add(btnGoPage3);
        btnGoPage3.Click += new System.EventHandler(btnGoPage3_Click);

        HFD_CurrentPage3 = new HiddenField();
        HFD_CurrentPage3.Value = "1";
        HFD_CurrentPage3.ID = "HFD_CurrentPage3";
        form1.Controls.Add(HFD_CurrentPage3);

        //第4個 GridView
        btnNextPage4 = new Button();
        btnNextPage4.ID = "btnNextPage4";
        form1.Controls.Add(btnNextPage4);
        btnNextPage4.Click += new System.EventHandler(btnNextPage4_Click);

        btnPreviousPage4 = new Button();
        btnPreviousPage4.ID = "btnPreviousPage4";
        form1.Controls.Add(btnPreviousPage4);
        btnPreviousPage4.Click += new System.EventHandler(btnPreviousPage4_Click);

        btnGoPage4 = new Button();
        btnGoPage4.ID = "btnGoPage4";
        form1.Controls.Add(btnGoPage4);
        btnGoPage4.Click += new System.EventHandler(btnGoPage4_Click);

        HFD_CurrentPage4 = new HiddenField();
        HFD_CurrentPage4.Value = "1";
        HFD_CurrentPage4.ID = "HFD_CurrentPage4";
        form1.Controls.Add(HFD_CurrentPage4);

        //第5個 GridView
        btnNextPage5 = new Button();
        btnNextPage5.ID = "btnNextPage5";
        form1.Controls.Add(btnNextPage5);
        btnNextPage5.Click += new System.EventHandler(btnNextPage5_Click);

        btnPreviousPage5 = new Button();
        btnPreviousPage5.ID = "btnPreviousPage5";
        form1.Controls.Add(btnPreviousPage5);
        btnPreviousPage5.Click += new System.EventHandler(btnPreviousPage5_Click);

        btnGoPage5 = new Button();
        btnGoPage5.ID = "btnGoPage5";
        form1.Controls.Add(btnGoPage5);
        btnGoPage5.Click += new System.EventHandler(btnGoPage5_Click);

        HFD_CurrentPage5 = new HiddenField();
        HFD_CurrentPage5.Value = "1";
        HFD_CurrentPage5.ID = "HFD_CurrentPage5";
        form1.Controls.Add(HFD_CurrentPage5);
    }

    //第1個 GridView
    protected void btnPreviousPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.MinusStringNumber(HFD_CurrentPage.Value);
        lblStreetList.Text = GetStreetList();
    }
    protected void btnNextPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.AddStringNumber(HFD_CurrentPage.Value);
        lblStreetList.Text = GetStreetList();
    }
    protected void btnGoPage_Click(object sender, EventArgs e)
    {
        lblStreetList.Text = GetStreetList();
    }

    //第2個 GridView
    protected void btnPreviousPage2_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage2.Value = Util.MinusStringNumber(HFD_CurrentPage2.Value);
        lblWomans.Text = GetWomanList();
    }
    protected void btnNextPage2_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage2.Value = Util.AddStringNumber(HFD_CurrentPage2.Value);
        lblWomans.Text = GetWomanList();
    }
    protected void btnGoPage2_Click(object sender, EventArgs e)
    {
        lblWomans.Text = GetWomanList();
    }

    //第3個 GridView
    protected void btnPreviousPage3_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage3.Value = Util.MinusStringNumber(HFD_CurrentPage3.Value);
        lblSpecialEvent.Text = GetSpecialEventList();
    }
    protected void btnNextPage3_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage3.Value = Util.AddStringNumber(HFD_CurrentPage3.Value);
        lblSpecialEvent.Text = GetSpecialEventList();
    }
    protected void btnGoPage3_Click(object sender, EventArgs e)
    {
        lblSpecialEvent.Text = GetSpecialEventList();
    }

    //第4個 GridView
    protected void btnPreviousPage4_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage4.Value = Util.MinusStringNumber(HFD_CurrentPage4.Value);
        lblNightSleep.Text = GetNightSleepList();
    }
    protected void btnNextPage4_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage4.Value = Util.AddStringNumber(HFD_CurrentPage4.Value);
        lblNightSleep.Text = GetNightSleepList();
    }
    protected void btnGoPage4_Click(object sender, EventArgs e)
    {
        lblNightSleep.Text = GetNightSleepList();
    }

    //第5個 GridView
    protected void btnPreviousPage5_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage5.Value = Util.MinusStringNumber(HFD_CurrentPage5.Value);
        lblMoneyApply.Text = GetMoneyApplyList();
    }
    protected void btnNextPage5_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage5.Value = Util.AddStringNumber(HFD_CurrentPage5.Value);
        lblMoneyApply.Text = GetMoneyApplyList();
    }
    protected void btnGoPage5_Click(object sender, EventArgs e)
    {
        lblMoneyApply.Text = GetMoneyApplyList();
    }
    #endregion NpoGridView 處理換頁相關程式碼
    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        string JavascriptCode = "$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();";
        JavascriptCode += "$('#btnPreviousPage2').hide();$('#btnNextPage2').hide();$('#btnGoPage2').hide();";
        JavascriptCode += "$('#btnPreviousPage3').hide();$('#btnNextPage3').hide();$('#btnGoPage3').hide();";
        JavascriptCode += "$('#btnPreviousPage4').hide();$('#btnNextPage4').hide();$('#btnGoPage4').hide();";
        JavascriptCode += "$('#btnPreviousPage5').hide();$('#btnNextPage5').hide();$('#btnGoPage5').hide();";
        Util.ShowSysMsgWithScript(JavascriptCode);

        if (!IsPostBack)
        {
            //公告訊息
            lblNews.Text = GetNewData();
            //寒士待處理案件列表
            lblStreetList.Text = GetStreetList();
            //單媽待處理案件列表
            lblWomans.Text = GetWomanList();
            //特殊事件處理列表
            lblSpecialEvent.Text = GetSpecialEventList();
            //夜宿申請表列表
            lblNightSleep.Text = GetNightSleepList();
            //急難救助金暨醫療費申請表列表
            lblMoneyApply.Text = GetMoneyApplyList();
        }
    }
    //------------------------------------------------------------------------------
    public string GetStreetList()
    {
        string strSql = "";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        strSql = @" 
                    select s.uid, 
                    d.DeptName as 站別,
                    s.CaseName as 姓名
                    from StreetData s
                    inner join Dept d on d.DeptID=s.DeptID
                    where isnull(s.IsDelete, 'N') = 'N'
                  ";
        string GroupName = SessionInfo.GroupName;
        switch (GroupName)
        {
            case "站長":
                strSql += @"
                            and isnull(s.CommitFlag, '') <> 'Y'
                            and s.DeptID=@DeptID -- //各平安站只能看到自己的
                           ";
                break;
            case "當地主管":
                strSql += @"
                            and isnull(s.CommitFlag, '') = 'Y'
                            and isnull(s.ExamineFlag, '') <> 'Y'
                            and s.DeptID=@DeptID -- //各平安站只能看到自己的
                           ";
                break;
            case "社工部社工":
                strSql += @"
                            and isnull(s.ExamineFlag, '') = 'Y'
                            and isnull(s.ReExamine1Flag, '') <> 'Y'
                           ";
                break;
            case "社工部主管":
                strSql += @"
                            and isnull(s.ReExamine1Flag, '') = 'Y'
                            and isnull(s.ReExamine2Flag, '') <> 'Y'
                           ";
                break;
            case "副秘書長":
                strSql += @"
                            and isnull(s.ReExamine2Flag, '') = 'Y'
                            and isnull(s.ApproveFlag, '') <> 'Y'
                           ";
                break;
            default: //其他的群組都不能審
                strSql += " and 1=0\n";
                break;
        }
        dict.Add("DeptID", SessionInfo.DeptID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        //需自訂欄位
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("uid");
        npoGridView.DisableColumn.Add("uid");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
        npoGridView.EditLink = Util.RedirectByTime("../CaseMgr/Street_Edit.aspx", "StreetUID=");
        npoGridView.PageSize = 5;
        npoGridView.PreviousPageLinkID = "btnPreviousPage";
        npoGridView.NextPageLinkID = "btnNextPage";
        npoGridView.GoPageLinkID = "btnGoPage";
        npoGridView.GoPageControlName = "GoPage";

        NPOGridViewColumn col;
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("流水號");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("uid");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("站別");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("站別");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("姓名");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("姓名");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        return npoGridView.Render();
    }
    //------------------------------------------------------------------------------
    public string GetWomanList()
    {
        string strSql = "";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        strSql = @" 
                    select w.uid, 
                    d.DeptName as 站別,
                    w.CaseName as 姓名
                    from Woman w
                    inner join Dept d on d.DeptID=w.DeptID
                    where isnull(w.IsDelete, 'N') = 'N'
                  ";
        string GroupName = SessionInfo.GroupName;
        switch (GroupName)
        {
            case "站長":
                strSql += @"
                            and isnull(w.CommitFlag, '') <> 'Y'
                            and w.DeptID=@DeptID -- //各平安站只能看到自己的
                           ";
                break;
            case "當地主管":
                strSql += @"
                            and isnull(w.CommitFlag, '') = 'Y'
                            and isnull(w.ExamineFlag, '') <> 'Y'
                            and w.DeptID=@DeptID -- //各平安站只能看到自己的
                           ";
                break;
            case "社工部社工":
                strSql += @"
                            and isnull(w.ExamineFlag, '') = 'Y'
                            and isnull(w.ReExamine1Flag, '') <> 'Y'
                           ";
                break;
            case "社工部主管":
                strSql += @"
                            and isnull(w.ReExamine1Flag, '') = 'Y'
                            and isnull(w.ReExamine2Flag, '') <> 'Y'
                           ";
                break;
            case "副秘書長":
                strSql += @"
                            and isnull(w.ReExamine2Flag, '') = 'Y'
                            and isnull(w.ApproveFlag, '') <> 'Y'
                           ";
                break;
            default: //其他的群組都不能審
                strSql += " and 1=0\n";
                break;
        }
        dict.Add("DeptID", SessionInfo.DeptID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        //需自訂欄位
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("uid");
        npoGridView.DisableColumn.Add("uid");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage2.Value);
        npoGridView.EditLink = Util.RedirectByTime("../CaseMgr/Woman_Edit2.aspx", "WomanUID=");
        npoGridView.PageSize = 5;
        npoGridView.CurrentPageLinkID = "HFD_CurrentPage2";
        npoGridView.PreviousPageLinkID = "btnPreviousPage2";
        npoGridView.NextPageLinkID = "btnNextPage2";
        npoGridView.GoPageLinkID = "btnGoPage2";
        npoGridView.GoPageControlName = "GoPage2";

        NPOGridViewColumn col;
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("流水號");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("uid");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("站別");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("站別");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("姓名");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("姓名");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        return npoGridView.Render();
    }
    //------------------------------------------------------------------------------
    public string GetSpecialEventList()
    {
        string strSql = "";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        strSql = @" 
                    select se.uid, se.StreetUID,
                    d.DeptName as 站別,
                    s.CaseName as 姓名,
                    convert(char(10), se.EventWhen, 111) as 時間,
                    se.EventWhere as 地點
                    from SpecialEvent se
                    inner join StreetData s on se.StreetUID=s.uid
                    inner join Dept d on d.DeptID=s.DeptID
                    where isnull(se.IsDelete, 'N') = 'N'
                  ";
        string GroupName = SessionInfo.GroupName;
        switch (GroupName)
        {
            case "站長":
                strSql += @"
                            and isnull(se.CommitFlag, '') <> 'Y'
                            and s.DeptID=@DeptID -- //各平安站只能看到自己的
                           ";
                break;
            case "當地主管":
                strSql += @"
                            and isnull(se.CommitFlag, '') = 'Y'
                            and isnull(se.ExamineFlag, '') <> 'Y'
                            and s.DeptID=@DeptID -- //各平安站只能看到自己的
                           ";
                break;
            case "社工部社工":
                strSql += @"
                            and isnull(se.ExamineFlag, '') = 'Y'
                            and isnull(se.ReExamine1Flag, '') <> 'Y'
                           ";
                break;
            case "社工部主管":
                strSql += @"
                            and isnull(se.ReExamine1Flag, '') = 'Y'
                            and isnull(se.ReExamine2Flag, '') <> 'Y'
                           ";
                break;
            case "副秘書長":
                strSql += @"
                            and isnull(se.ReExamine2Flag, '') = 'Y'
                            and isnull(se.ApproveFlag, '') <> 'Y'
                           ";
                break;
            default: //其他的群組都不能審
                strSql += " and 1=0\n";
                break;
        }
        dict.Add("DeptID", SessionInfo.DeptID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        //需自訂欄位
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("uid");
        npoGridView.DisableColumn.Add("uid");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage3.Value);
        if (dt.Rows.Count > 0)
        {
            npoGridView.EditLink = Util.RedirectByTime("../CaseMgr/SpecialEvent_Edit.aspx", "StreetUID=" + dt.Rows[0]["StreetUID"].ToString() + "&SpecialEventUID=");
        }
        npoGridView.PageSize = 5;
        npoGridView.CurrentPageLinkID = "HFD_CurrentPage3";
        npoGridView.PreviousPageLinkID = "btnPreviousPage3";
        npoGridView.NextPageLinkID = "btnNextPage3";
        npoGridView.GoPageLinkID = "btnGoPage3";
        npoGridView.GoPageControlName = "GoPage3";

        NPOGridViewColumn col;
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("流水號");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("uid");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("站別");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("站別");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("姓名");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("姓名");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("時間");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("時間");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("地點");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("地點");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        return npoGridView.Render();
    }
    //------------------------------------------------------------------------------
    public string GetNightSleepList()
    {
        string strSql = "";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        strSql = @" 
                    select ns.uid, ns.StreetUID,
                    d.DeptName as 站別,
                    s.CaseName as 姓名,
                    convert(char(10), ns.ApplyDate, 111) as 申請日期
                    from NightSleep ns
                    inner join StreetData s on ns.StreetUID=s.uid
                    inner join Dept d on d.DeptID=s.DeptID
                    where isnull(ns.IsDelete, 'N') = 'N'
                  ";
        string GroupName = SessionInfo.GroupName;
        switch (GroupName)
        {
            case "站長":
                strSql += @"
                            and isnull(ns.CommitFlag, '') <> 'Y'
                            and s.DeptID=@DeptID -- //各平安站只能看到自己的
                           ";
                break;
            case "當地主管":
                strSql += @"
                            and isnull(ns.CommitFlag, '') = 'Y'
                            and isnull(ns.ExamineFlag, '') <> 'Y'
                            and s.DeptID=@DeptID -- //各平安站只能看到自己的
                           ";
                break;
            case "社工部社工":
                strSql += @"
                            and isnull(ns.ExamineFlag, '') = 'Y'
                            and isnull(ns.ReExamine1Flag, '') <> 'Y'
                           ";
                break;
            case "社工部主管":
                strSql += @"
                            and isnull(ns.ReExamine1Flag, '') = 'Y'
                            and isnull(ns.ReExamine2Flag, '') <> 'Y'
                           ";
                break;
            case "副秘書長":
                strSql += @"
                            and isnull(ns.ReExamine2Flag, '') = 'Y'
                            and isnull(ns.ApproveFlag, '') <> 'Y'
                           ";
                break;
            default: //其他的群組都不能審
                strSql += " and 1=0\n";
                break;
        }
        dict.Add("DeptID", SessionInfo.DeptID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        //需自訂欄位
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("uid");
        npoGridView.DisableColumn.Add("uid");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage3.Value);
        if (dt.Rows.Count > 0)
        {
            npoGridView.EditLink = Util.RedirectByTime("../CaseMgr/NightSleep_Edit.aspx", "StreetUID=" + dt.Rows[0]["StreetUID"].ToString() + "&NightSleepUID=");
        }
        npoGridView.PageSize = 5;
        npoGridView.CurrentPageLinkID = "HFD_CurrentPage4";
        npoGridView.PreviousPageLinkID = "btnPreviousPage4";
        npoGridView.NextPageLinkID = "btnNextPage4";
        npoGridView.GoPageLinkID = "btnGoPage4";
        npoGridView.GoPageControlName = "GoPage4";

        NPOGridViewColumn col;
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("流水號");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("uid");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("站別");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("站別");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("姓名");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("姓名");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("申請日期");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("申請日期");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        return npoGridView.Render();
    }
    //------------------------------------------------------------------------------
    public string GetMoneyApplyList()
    {
        string strSql = "";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        strSql = @" 
                    select m.uid, m.StreetUID,
                    d.DeptName as 站別,
                    s.CaseName as 姓名,
                    convert(char(10), m.ApplyDate, 111) as 申請日期
                    from MoneyApply m
                    inner join StreetData s on m.StreetUID=s.uid
                    inner join Dept d on d.DeptID=s.DeptID
                    where isnull(m.IsDelete, 'N') = 'N'
                  ";
        string GroupName = SessionInfo.GroupName;
        switch (GroupName)
        {
            case "站長":
                strSql += @"
                            and isnull(m.CommitFlag, '') <> 'Y'
                            and s.DeptID=@DeptID -- //各平安站只能看到自己的
                           ";
                break;
            case "當地主管":
                strSql += @"
                            and isnull(m.CommitFlag, '') = 'Y'
                            and isnull(m.ExamineFlag, '') <> 'Y'
                            and s.DeptID=@DeptID -- //各平安站只能看到自己的
                           ";
                break;
            case "社工部社工":
                strSql += @"
                            and isnull(m.ExamineFlag, '') = 'Y'
                            and isnull(m.ReExamine1Flag, '') <> 'Y'
                           ";
                break;
            case "社工部主管":
                strSql += @"
                            and isnull(m.ReExamine1Flag, '') = 'Y'
                            and isnull(m.ReExamine2Flag, '') <> 'Y'
                           ";
                break;
            case "副秘書長":
                strSql += @"
                            and isnull(m.ReExamine2Flag, '') = 'Y'
                            and isnull(m.ApproveFlag, '') <> 'Y'
                           ";
                break;
            default: //其他的群組都不能審
                strSql += " and 1=0\n";
                break;
        }
        dict.Add("DeptID", SessionInfo.DeptID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        //需自訂欄位
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("uid");
        npoGridView.DisableColumn.Add("uid");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage3.Value);
        if (dt.Rows.Count > 0)
        {
            npoGridView.EditLink = Util.RedirectByTime("../CaseMgr/MoneyApply_Edit.aspx", "StreetUID=" + dt.Rows[0]["StreetUID"].ToString() + "&MoneyApplyUID=");
        }
        npoGridView.PageSize = 5;
        npoGridView.CurrentPageLinkID = "HFD_CurrentPage5";
        npoGridView.PreviousPageLinkID = "btnPreviousPage5";
        npoGridView.NextPageLinkID = "btnNextPage5";
        npoGridView.GoPageLinkID = "btnGoPage5";
        npoGridView.GoPageControlName = "GoPage5";

        NPOGridViewColumn col;
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("流水號");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("uid");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("站別");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("站別");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("姓名");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("姓名");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        col = new NPOGridViewColumn("申請日期");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("申請日期");
        npoGridView.Columns.Add(col);
        //-----------------------------------------------------------------------
        return npoGridView.Render();
    }
    //------------------------------------------------------------------------------
    public string GetNewData()
    {
        string strSql = "";
        DataTable dt;
        string SysDate =  Util.DateTime2String(DateTime.Now, DateType.yyyyMMdd, EmptyType.ReturnEmpty);
        strSql = "select top 2 News.* from news\n";
        strSql += "where @SysDate between NewsBeginDate and NewsEndDate\n";
        strSql += "and DeptName like @DeptName\n";
        strSql += "order by NewsRegDate Desc";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("SysDate", SysDate);
        dict.Add("DeptName", "%" + SessionInfo.DeptName + "%");
        dt = NpoDB.GetDataTableS(strSql, dict);

        StringBuilder sb = new StringBuilder();
        sb.AppendLine("<table width='98%' id='AutoNumber1' class='table_v' >");
        sb.AppendLine("<tr>");
        sb.AppendLine("<td width='100%' colspan='2'> </td>");
        sb.AppendLine("</tr>");
        foreach (DataRow dr in dt.Rows)
        {
            sb.AppendLine("<tr>");
            sb.AppendLine(" <td width='4%' align='right'><img border='0' src='../images/DIR_tri.gif' /></td>");
            sb.AppendLine(@" <td width='96%' style='color:#660000' align='left'>");
            //sb.AppendLine("<a href='../filemgr/News_Show.aspx?NewsUID=" + dr["uid"].ToString() + "' class='news'>" + dr["NewsSubject"].ToString() + "</a> ");
            sb.AppendLine("<a href='../filemgr/News_Show.aspx?NewsUID=" + dr["uid"].ToString() + "' class='news'>" +"【"+ dr["NewsType"].ToString() +"】"+ dr["NewsSubject"].ToString() + "</a> ");
            //sb.AppendLine("(" + dr["NewsRegDate"] + ")");
            sb.AppendLine("(" + Convert.ToDateTime(dr["NewsBeginDate"].ToString()).ToString("yyyy/MM/dd") + "～" + Convert.ToDateTime(dr["NewsEndDate"].ToString()).ToString("yyyy/MM/dd") + ")");
            sb.AppendLine("</td>");
            sb.AppendLine("</tr>");
            sb.AppendLine("<tr>");
            sb.AppendLine("<td width='4%' height='5'></td>");
            sb.AppendLine("<td width='96%' style='color:#4B4B4B' height='5'  align='left'> ");
            sb.AppendLine(dr["NewsBrief"].ToString());
            sb.AppendLine("</td>");
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</table>");
        return sb.ToString();
    }
}
