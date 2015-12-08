using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Online_DonateMobile : System.Web.UI.Page
{
    string strIpayHttp = System.Configuration.ConfigurationManager.AppSettings["cathay_epos"];
    protected void Page_Load(object sender, EventArgs e)
    {
        strRqXML.Value = Request["strRqXML"].ToString();
        ClientScript.RegisterStartupScript(this.GetType(), "post", "<script>document.forms[0].action='" + strIpayHttp + "'; document.forms[0].target='my_iframe';document.forms[0].submit();</script>");
    }
}