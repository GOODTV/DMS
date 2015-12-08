using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonateMgr_Fdc : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "Fdc";
        //權控處理
        AuthrityControl();
        if (!IsPostBack)
        {
            LoadDropDownListData();
            OrgData();
        }
    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_Print", btnTxt);

    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.SelectedIndex = 0;

        //捐款年度
        for (int i = 0; i < 3; i++)
        {
            ddlDonate_Date_Year.Items.Insert(i, new ListItem((int.Parse(DateTime.Now.Year.ToString()) - i).ToString(), (int.Parse(DateTime.Now.Year.ToString()) - i).ToString()));
        }
        //募款活動
        Util.FillDropDownList(ddlAct_Id, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        ddlAct_Id.Items.Insert(0, new ListItem("", ""));
        ddlAct_Id.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    protected void btnTxt_Click(object sender, EventArgs e)
    {
        string strSql = @"select Donor_Id
                          from Donor
                          where IsFdc = 'Y' and IDNo != '' and Dept_Id= @Dept_Id";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Dept_Id",ddlDept.SelectedValue);
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);
        DataRow dr;
        string Donor_Id = "";
        for(int i = 0; i < dt.Rows.Count;i++)
        {
            dr = dt.Rows[i];
            Donor_Id += dr["Donor_Id"].ToString() + " ";
        }

        Session["Donor_Id"] = Donor_Id;
        Session["Donate_Date_Year"] = ddlDonate_Date_Year.SelectedItem.Text;
        Session["Act_Id"] = ddlAct_Id.SelectedValue;
        Session["Uniform_No"] = tbxUniform_No.Text.Trim();
        Session["Licence"] = tbxLicence.Text.Trim();

        Response.Redirect(Util.RedirectByTime("Fdc_Print.aspx"));
    }
    protected void ddlDept_SelectedIndexChanged(object sender, EventArgs e)
    {
        OrgData();
    }
    public void OrgData()
    {
        string strSql = @"select * from Organization where OrgID=@OrgID";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("OrgID", ddlDept.SelectedValue);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        DataRow dr = dt.Rows[0];

        Session["OrgName"] = dr["OrgName"].ToString();
        tbxUniform_No.Text = dr["Uniform_No"].ToString();
        tbxLicence.Text = dr["Licence"].ToString();
    }
}