using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateDone: System.Web.UI.Page
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
            Session["DonorID"] = null;
            Session["DonorName"] = null;
            Session["InsertPeriod"] = null;
        }
    }
    //---------------------------------------------------------------------------
    private void LoadFormData()
    {

    }
    //---------------------------------------------------------------------------
    protected void btnBackDefault_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("donate_index.html"));
    }
    //---------------------------------------------------------------------------
}