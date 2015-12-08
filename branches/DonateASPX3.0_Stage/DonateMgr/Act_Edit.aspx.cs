using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class DonateMgr_Act_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Act_Id");
            LoadDropDownListData();
            Form_DataBind();
        }
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.SelectedIndex = 0;
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
        strSql = " select *  from Act where Act_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("ActQry.aspx");

        DataRow dr = dt.Rows[0];

        //機構
        ddlDept.Text = dr["Dept_Id"].ToString().Trim();
        //活動名稱
        tbxAct_Name.Text = dr["Act_Name"].ToString().Trim();
        //活動簡稱
        tbxAct_ShortName.Text = dr["Act_ShortName"].ToString().Trim();
        //主辦單位
        tbxAct_OrgName.Text = dr["Act_OrgName"].ToString().Trim();
        //協辦單位
        tbxAct_OrgName2.Text = dr["Act_OrgName2"].ToString().Trim();
        //活動主題
        tbxAct_Subject.Text = dr["Act_Subject"].ToString().Trim();
        //活動期間
        tbxAct_BeginDate.Text = DateTime.Parse(dr["Act_BeginDate"].ToString()).ToString("yyyy/MM/dd");
        tbxAct_EndDate.Text = DateTime.Parse(dr["Act_EndDate"].ToString()).ToString("yyyy/MM/dd"); 
        //勸募許可文號
        tbxAct_Licence.Text = dr["Act_Licence"].ToString().Trim();
        //備註
        tbxRemark.Text = dr["Remark"].ToString().Trim();
    }
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            Edit();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("資料修改成功！");
            Response.Redirect(Util.RedirectByTime("Act_Edit.aspx", "Act_Id=" + HFD_Uid.Value));
        }
    }
    protected void btnDel_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            Del();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("資料刪除成功！");
            Response.Redirect(Util.RedirectByTime("ActQry.aspx"));
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("ActQry.aspx"));
    }
    public void Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update Act set ";

        strSql += "  Dept_Id = @Dept_Id";
        strSql += ", Act_Name = @Act_Name";
        strSql += ", Act_ShortName = @Act_ShortName";
        strSql += ", Act_OrgName = @Act_OrgName";
        strSql += ", Act_OrgName2 = @Act_OrgName2";
        strSql += ", Act_Subject = @Act_Subject";
        strSql += ", Act_Licence = @Act_Licence";
        strSql += ", Act_BeginDate = @Act_BeginDate";
        strSql += ", Act_EndDate = @Act_EndDate";
        strSql += ", Remark = @Remark";
        strSql += " where Act_Id = @Act_Id";

        dict.Add("Dept_Id", ddlDept.SelectedValue);
        dict.Add("Act_Name", tbxAct_Name.Text.Trim());
        dict.Add("Act_ShortName", tbxAct_ShortName.Text.Trim());
        dict.Add("Act_OrgName", tbxAct_OrgName.Text.Trim());
        dict.Add("Act_OrgName2", tbxAct_OrgName2.Text.Trim());
        dict.Add("Act_Subject", tbxAct_Subject.Text.Trim());
        dict.Add("Act_Licence", tbxAct_Licence.Text.Trim());
        dict.Add("Act_BeginDate", tbxAct_BeginDate.Text.Trim());
        dict.Add("Act_EndDate", tbxAct_EndDate.Text.Trim());
        dict.Add("Remark", tbxRemark.Text.Trim());
        dict.Add("Act_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void Del()
    {
        string strSql = "delete from Act where Act_Id=@Act_Id";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Act_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}