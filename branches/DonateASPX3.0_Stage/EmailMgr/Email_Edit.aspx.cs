using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class EmailMgr_Email_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Ser_No");
            LoadDropDownListData();
            Form_DataBind();
            Banner_Bind();
        }
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //郵件類別
        Util.FillDropDownList(ddlEmailMgr_Type, Util.GetDataTable("CaseCode", "GroupName", "郵件類別", "", ""), "CaseName", "CaseName", false);
        ddlEmailMgr_Type.Items.Insert(0, new ListItem("", ""));

        //郵件類別
        Util.FillDropDownList(ddlEmailMgr_Type_Banner, Util.GetDataTable("CaseCode", "GroupName", "郵件類別", "", ""), "CaseName", "CaseName", false);
        ddlEmailMgr_Type_Banner.Items.Insert(0, new ListItem("", ""));
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
        strSql = " select *  from EmailMgr where Ser_No='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("EmailList.aspx");

        DataRow dr = dt.Rows[0];

        //郵件類別
        ddlEmailMgr_Type.Text = dr["EmailMgr_Type"].ToString().Trim();
        ddlEmailMgr_Type_Banner.Text = dr["EmailMgr_Type"].ToString().Trim();
        //郵件標題
        tbxEmailMgr_Subject.Text = dr["EmailMgr_Subject"].ToString().Trim();
        //郵件內容
        tbxEmailMgr_Desc.Text = dr["EmailMgr_Desc"].ToString().Trim();
        //Label1.Text = HttpUtility.HtmlEncode(Me.ckdata.Text)
    }
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
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            EmailMgr_Edit();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("資料修改成功！");
            Response.Redirect(Util.RedirectByTime("Email_Edit.aspx", "Ser_No=" + HFD_Uid.Value));
        }
    }
    protected void btnDel_Click(object sender, EventArgs e)
    {
        string strSql = "delete from EmailMgr where Ser_No=@Ser_No";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Ser_No", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);

        SetSysMsg("資料刪除成功！");
        Response.Redirect(Util.RedirectByTime("EmailList.aspx"));
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("EmailList.aspx"));
    }
    public void EmailMgr_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update EmailMgr set ";

        strSql += "  EmailMgr_Type = @EmailMgr_Type";
        strSql += ", EmailMgr_Subject = @EmailMgr_Subject";
        strSql += ", EmailMgr_Desc = @EmailMgr_Desc";
        strSql += " where Ser_No = @Ser_No";

        dict.Add("EmailMgr_Type", ddlEmailMgr_Type.SelectedItem.Text);
        dict.Add("EmailMgr_Subject", tbxEmailMgr_Subject.Text.Trim());
        dict.Add("EmailMgr_Desc", tbxEmailMgr_Desc.Text.Trim());
        dict.Add("Ser_No", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}