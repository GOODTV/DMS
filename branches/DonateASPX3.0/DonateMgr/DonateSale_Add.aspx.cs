using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;


public partial class DonateMgr_DonateSaleQry_Add : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            LoadDropDownListData();
            tbxSale_BeginDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
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
            Donate_Sale_AddNew();
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
            string strSql = @"select Ser_No from Donate_Sale
                    where Sale_Subject ='" + tbxSale_Subject.Text.Trim() + "' and convert(nvarchar(255),Sale_Content) ='" + tbxSale_Content.Text.Trim() + "'";
            //****執行查詢勸募活動****//
            DataTable dt = NpoDB.QueryGetTable(strSql);
            DataRow dr = dt.Rows[0];
            string Ser_No = dr["Ser_No"].ToString().Trim();

            Response.Redirect(Util.RedirectByTime("DonateSale_Edit.aspx", "Ser_No=" + Ser_No));
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("DonateSaleQry.aspx"));
    }
    public void Donate_Sale_AddNew()
    {
        string strSql = "insert into  Donate_Sale\n";
        strSql += "( Dept_Id, Sale_Subject, Sale_BeginDate, Sale_EndDate, Sale_Content) values\n";
        strSql += "( @Dept_Id,@Sale_Subject,@Sale_BeginDate,@Sale_EndDate,@Sale_Content)\n";
        strSql += "select @@IDENTITY";


        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Dept_Id", ddlDept.SelectedValue);
        dict.Add("Sale_Subject", tbxSale_Subject.Text.Trim());
        dict.Add("Sale_BeginDate", tbxSale_BeginDate.Text.Trim());
        dict.Add("Sale_EndDate", tbxSale_EndDate.Text.Trim());
        dict.Add("Sale_Content", tbxSale_Content.Text.Trim());
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}