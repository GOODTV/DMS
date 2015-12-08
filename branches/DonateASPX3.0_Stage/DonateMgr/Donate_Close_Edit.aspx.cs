using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class DonateMgr_Donate_Close_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Dept_Id");
            LoadDropDownListData();
            Form_DataBind();
        }
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        if (SessionInfo.UserID != "nopis")
        {
            string strSql = "Select DeptId, UserName, GroupID From AdminUser Where UserID<>'npois' Order By DeptId,GroupID";
            Util.FillDropDownList(ddlDonate_Open_User, strSql, "UserName", "UserName", true);
            Util.FillDropDownList(ddlContribute_Open_User, strSql, "UserName", "UserName", true);
        }
        else
        {
            string strSql = "Select DeptId, UserName, GroupID From AdminUser Order By DeptId,GroupID";
            Util.FillDropDownList(ddlDonate_Open_User, strSql, "UserName", "UserName", true);
            Util.FillDropDownList(ddlContribute_Open_User, strSql, "UserName", "UserName", true);
        }
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
        strSql = " select *  from Dept where DeptId='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("Donate_Close.aspx");

        DataRow dr = dt.Rows[0];

        //機構
        tbxDept.Text = dr["DeptShortName"].ToString().Trim();
        //捐款關帳日
        tbxDonate_Close.Text = Util.DateTime2String(dr["Donate_Close"], DateType.yyyyMMdd, EmptyType.ReturnEmpty);
        //開放權限修改人員
        ddlDonate_Open_User.Text = dr["Donate_Open_User"].ToString().Trim();
        //捐物資關帳日
        tbxContribute_Close.Text = Util.DateTime2String(dr["Contribute_Close"], DateType.yyyyMMdd, EmptyType.ReturnEmpty);
        //開放權限修改人員
        ddlContribute_Open_User.Text = dr["Contribute_Open_User"].ToString().Trim();
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
            Response.Redirect(Util.RedirectByTime("Donate_Close_Edit.aspx", "Dept_Id=" + HFD_Uid.Value));
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("Donate_Close.aspx"));
    }
    public void Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update Dept set ";

        strSql += " Donate_Close = @Donate_Close";
        strSql += ", Donate_Open_User = @Donate_Open_User";
        strSql += ", Donate_Open_LastDate = @Donate_Open_LastDate";
        strSql += ", Contribute_Close = @Contribute_Close";
        strSql += ", Contribute_Open_User = @Contribute_Open_User";
        strSql += ", Contribute_Open_LastDate = @Contribute_Open_LastDate";
        strSql += " where DeptId = @DeptId";

        dict.Add("Donate_Close", Util.DateTime2String(tbxDonate_Close.Text, DateType.yyyyMMdd, EmptyType.ReturnNull)); 
        if (ddlDonate_Open_User.SelectedIndex != 0)
        {
            dict.Add("Donate_Open_User", ddlDonate_Open_User.SelectedItem.Text);
            dict.Add("Donate_Open_LastDate", DateTime.Now.ToString("yyyy-MM-dd"));
        }
        else
        {
            dict.Add("Donate_Open_User", "");
            dict.Add("Donate_Open_LastDate", null);
        }
        dict.Add("Contribute_Close",  Util.DateTime2String(tbxContribute_Close.Text, DateType.yyyyMMdd, EmptyType.ReturnNull)); 
        if (ddlContribute_Open_User.SelectedIndex != 0)
        {
            dict.Add("Contribute_Open_User", ddlContribute_Open_User.SelectedItem.Text);
            dict.Add("Contribute_Open_LastDate", DateTime.Now.ToString("yyyy-MM-dd"));
        }
        else
        {
            dict.Add("Contribute_Open_User", "");
            dict.Add("Contribute_Open_LastDate", null);
        }
        dict.Add("DeptId", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}