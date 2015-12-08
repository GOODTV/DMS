using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateOnlineAll : System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {

            Item.Items.Add(new ListItem("請選擇...", ""));
            Item.Items.Add(new ListItem("單筆奉獻", "單筆奉獻"));
            Item.Items.Add(new ListItem("定期定額奉獻", "定期定額奉獻"));

            HFD_chkItem.Value = "";
            //奉獻金額
            HFD_Amount.Value = "";
            //繳費方式

        }

    }

}
