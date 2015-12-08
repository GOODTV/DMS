using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class EmailMgr_Email_Add : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            LoadDropDownListData();
            Banner_Bind();
        }
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //郵件類別
        Util.FillDropDownList(ddlEmailMgr_Type, Util.GetDataTable("CaseCode", "GroupName", "郵件類別", "", ""), "CaseName", "CaseName", false);
        ddlEmailMgr_Type.Items.Insert(0, new ListItem("", ""));
        ddlEmailMgr_Type.SelectedIndex = 0;

        //郵件類別
        Util.FillDropDownList(ddlEmailMgr_Type_Banner, Util.GetDataTable("CaseCode", "GroupName", "郵件類別", "", ""), "CaseName", "CaseName", false);
        ddlEmailMgr_Type_Banner.Items.Insert(0, new ListItem("", ""));
        ddlEmailMgr_Type_Banner.SelectedIndex = 0;
    }
    //----------------------------------------------------------------------
    public void Banner_Bind()
    {
        if (ddlEmailMgr_Type_Banner.SelectedIndex != 0)
        {
            string strSql;
            DataTable dt;
            strSql = "select Ser_No, EmailMgr_Subject as [郵件標題]\n";
            strSql += " from EmailMgr where 1=1 ";
            if (ddlEmailMgr_Type_Banner.SelectedIndex != 0)
            {
                strSql += " and EmailMgr_Type ='" + ddlEmailMgr_Type_Banner.SelectedItem.Text + "'";
            }
            Dictionary<string, object> dict = new Dictionary<string, object>();
            dt = NpoDB.GetDataTableS(strSql, dict);

            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.ShowPage = false;
            npoGridView.Keys.Add("Ser_No");
            npoGridView.DisableColumn.Add("Ser_No");
            npoGridView.EditLink = Util.RedirectByTime("Email_Edit.aspx", "Ser_No=");
            lblGridList.Text = npoGridView.Render();
        }
        string Item = "Content";
        lblGridList2.Text = "<iframe name=\"left2\" src=\"image_left_list.aspx?dept_id=" + SessionInfo.DeptID + "&item=" + Item + "&ser_no=" + Util.GetQueryString("ser_no") + "&id=" + Util.GetQueryString("id") + "\" height=\"420\" width=\"100%\" frameborder=\"1\" scrolling=\"auto\" target=\"_self\"></iframe>";
    }
    protected void ddlEmailMgr_Type_Banner_SelectedIndexChanged(object sender, EventArgs e)
    {
        Banner_Bind();
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            EmailMgr_AddNew();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("資料新增成功！");
            //******勸募活動編號******//
            string strSql2 = @"select Ser_No from EmailMgr
                    where EmailMgr_Subject ='" + tbxEmailMgr_Subject.Text.Trim() + "' and EmailMgr_Desc ='" + tbxEmailMgr_Desc.Text.Trim() + "'";
            //****執行查詢勸募活動****//
            DataTable dt2 = NpoDB.QueryGetTable(strSql2);
            DataRow dr2 = dt2.Rows[0];
            string Ser_No = dr2["Ser_No"].ToString().Trim();

            Response.Redirect(Util.RedirectByTime("Email_Edit.aspx", "Ser_No=" + Ser_No));
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("EmailList.aspx"));
    }
    public void EmailMgr_AddNew()
    {
        string strSql = "insert into EmailMgr\n";
        strSql += "( EmailMgr_Type, EmailMgr_Subject, EmailMgr_Desc, EmailMgr_RegDate) values\n";
        strSql += "( @EmailMgr_Type,@EmailMgr_Subject,@EmailMgr_Desc,@EmailMgr_RegDate) ";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("EmailMgr_Type", ddlEmailMgr_Type.SelectedItem.Text);
        dict.Add("EmailMgr_Subject", tbxEmailMgr_Subject.Text.Trim());
        dict.Add("EmailMgr_Desc", tbxEmailMgr_Desc.Text.Trim());
        dict.Add("EmailMgr_RegDate", DateTime.Now.ToString("yyyy-MM-dd"));
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}