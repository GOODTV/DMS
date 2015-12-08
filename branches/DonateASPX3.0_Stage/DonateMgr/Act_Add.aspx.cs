using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class DonateMgr_Act_Add : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            LoadDropDownListData();
            tbxAct_BeginDate.Text = DateTime.Now.ToString("yyyy/MM/dd"); 
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
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            AddNew();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("勸募活動新增成功！");
            //******勸募活動編號******//
            string strSql2 = @"select Act_Id from Act
                    where Act_Name ='" + tbxAct_Name.Text.Trim() + "' and Act_ShortName ='" + tbxAct_ShortName.Text.Trim() + "'";
            //****執行查詢勸募活動****//
            DataTable dt2 = NpoDB.QueryGetTable(strSql2);
            DataRow dr2 = dt2.Rows[0];
            string Act_Id = dr2["Act_Id"].ToString().Trim();

            Response.Redirect(Util.RedirectByTime("Act_Edit.aspx", "Act_Id=" + Act_Id));
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("ActQry.aspx"));
    }
    public void AddNew()
    {
        string strSql = "insert into  Act\n";
        strSql += "( Dept_Id, Act_Name, Act_ShortName, Act_OrgName, Act_OrgName2, Act_Subject, Act_Licence,\n";
        strSql += "  Act_BeginDate,Act_EndDate, Remark) values\n";
        strSql += "( @Dept_Id,@Act_Name,@Act_ShortName,@Act_OrgName,@Act_OrgName2,@Act_Subject,@Act_Licence,\n";
        strSql += "  @Act_BeginDate,@Act_EndDate,@Remark) ";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
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
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}