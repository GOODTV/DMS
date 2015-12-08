using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateType : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            LoadFormData();
            if (Session["DonorName"] != null)
            {
                lblTitle.Text = Session["DonorName"].ToString() + lblTitle.Text;
            }
        }
    }
    //---------------------------------------------------------------------------
    private void LoadFormData()
    {
        //奉獻方式
        List<ControlData> list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_DonateType", "rdoDonateType1", "CreditCard", "線上刷卡立即奉獻", false, ""));
        lblType1.Text = HtmlUtil.RenderControl(list);
        //list = new List<ControlData>();
        //list.Add(new ControlData("Radio", "RDO_DonateType", "rdoDonateType2", "WebATM", "線上ATM立即轉帳奉獻", false, ""));
        //lblType2.Text = HtmlUtil.RenderControl(list);
    }
    //---------------------------------------------------------------------------
    protected void btnConfirm_Click(object sender, EventArgs e)
    {
        Response.Redirect("ShoppingCart.aspx?Mode=FromDefault");
    }
    //---------------------------------------------------------------------------
    protected void btnPrev_Click(object sender, EventArgs e)
    {
        //Response.Redirect("DonateOnlineDefault.aspx");
        Response.Redirect("ShoppingCart.aspx?Mode=FromDefault");
    }
}