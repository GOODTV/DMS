using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class DonateMgr_DonateSaleQry_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Ser_No");
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
        strSql = " select *  from Donate_Sale where Ser_No='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("DonateSaleQry.aspx");

        DataRow dr = dt.Rows[0];

        //機構
        ddlDept.Text = dr["Dept_Id"].ToString().Trim();
        //文宣標題
        tbxSale_Subject.Text = dr["Sale_Subject"].ToString().Trim();
        //活動期間
        tbxSale_BeginDate.Text = DateTime.Parse(dr["Sale_BeginDate"].ToString()).ToString("yyyy/MM/dd");
        tbxSale_EndDate.Text = DateTime.Parse(dr["Sale_EndDate"].ToString()).ToString("yyyy/MM/dd");
        //文宣內容
        tbxSale_Content.Text = dr["Sale_Content"].ToString().Trim();
    }
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            Donate_Sale_Edit();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("資料修改成功！");
            Response.Redirect(Util.RedirectByTime("DonateSale_Edit.aspx", "Ser_No=" + HFD_Uid.Value));
        }
    }
    protected void btnDel_Click(object sender, EventArgs e)
    {
        string strSql = "delete from Donate_Sale where Ser_No=@Ser_No";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Ser_No", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);

        SetSysMsg("資料刪除成功！");
        Response.Redirect(Util.RedirectByTime("DonateSaleQry.aspx"));
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("DonateSaleQry.aspx"));
    }
    public void Donate_Sale_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update Donate_Sale set ";

        strSql += "  Dept_Id = @Dept_Id";
        strSql += ", Sale_Subject = @Sale_Subject";
        strSql += ", Sale_BeginDate = @Sale_BeginDate";
        strSql += ", Sale_EndDate = @Sale_EndDate";
        strSql += ", Sale_Content = @Sale_Content";
        strSql += " where Ser_No = @Ser_No";

        dict.Add("Dept_Id", ddlDept.SelectedValue);
        dict.Add("Sale_Subject", tbxSale_Subject.Text.Trim());
        dict.Add("Sale_BeginDate", tbxSale_BeginDate.Text.Trim());
        dict.Add("Sale_EndDate", tbxSale_EndDate.Text.Trim());
        dict.Add("Sale_Content", tbxSale_Content.Text.Trim());
        dict.Add("Ser_No", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}