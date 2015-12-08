using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateOnlineDefault : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {
            lblItem1.Text = "※ 經常奉獻";
            lblItem2.Text = "※ 優質媒體奉獻";
            lblItem3.Text = "※ 其他";
            if (Session["DonorName"] != null)
            {
                lblTitle.Text = Session["DonorName"].ToString() + lblTitle.Text;
            }
        }
    }
    //---------------------------------------------------------------------------
    protected void btnCheckOut_Click(object sender, EventArgs e)
    {
        Response.Redirect("ShoppingCart.aspx?Mode=FromDefault");
    }
}