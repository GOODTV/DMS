using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateSingle: System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            //顯示奉獻成功訊息
            if (Session["Msg"] != null)
            {
                Util.ShowMsg(Session["Msg"].ToString());
                Session["Msg"] = null;
            }
            LoadFormData();
            if (Session["DonorName"] != null)
            {
                lblTitle.Text = Session["DonorName"].ToString() + lblTitle.Text;
            }
            LoadDropDownList();
        }
    }
    //-------------------------------------------------------------------------------------------------------------
    private void LoadDropDownList()
    {
    }
    //---------------------------------------------------------------------------
    private void LoadFormData()
    {
        string strSql = @"
                            select * from Donor where Donor_Id=@Donor_Id
                        ";
    }
    //---------------------------------------------------------------------------
    protected void btnBackDefault_Click(object sender, EventArgs e)
    {
        Session["ItemOnce"] = null;
        Response.Redirect(Util.RedirectByTime("DonateDone.aspx"));
    }
    //---------------------------------------------------------------------------
}