using System;
using System.Web.UI;

public partial class MainTop : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            btnLogout.ImageUrl = "../images/title_logout_bt.gif";  //預設值
            btnLogout.Attributes["onmouseover"] = "EvImageOverChange(this, 'in');";
            btnLogout.Attributes["onmouseout"] = "EvImageOverChange(this, 'out');";
        }

        try
        {
            txtDeptName.Text = SessionInfo.DeptName;
            txtUserGroup.Text = SessionInfo.GroupName + " (" + SessionInfo.UserName + ")";
        }
        catch (Exception ex)
        {
            Response.Write(@"<script>window.parent.location.href='../Default.aspx';</script>");
        }
    }
    //------------------------------------------------------------------------
    protected void btnLogout_Click(object sender, ImageClickEventArgs e)
    {
        string UserID = "";
        if (SessionInfo != null)
        {
            UserID = SessionInfo.UserID;
        }
        Response.Write(@"<script>window.parent.location.href='../Default.aspx?logout=true&UserID=" + UserID + "';</script>");
    }
    //------------------------------------------------------------------------
}
