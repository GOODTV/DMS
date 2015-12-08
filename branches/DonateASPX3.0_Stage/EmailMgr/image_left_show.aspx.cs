using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class EmailMgr_image_left_show : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        ltlGridView.Text = "<td bgcolor='#FFFFFF' align='center'><img src='../upload/" + Util.GetQueryString("imgfile") + "' ></td>";
    }
}