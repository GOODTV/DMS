using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class MemberMgr_MemberStatQry : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "MemberStatQry";
        //權控處理
        AuthrityControl();
    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_Print", btnPrint);
        Authrity.CheckButtonRight("_Print", btnToxls);
    }
    //---------------------------------------------------------------------------
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Response.Redirect("MemberStatQry_Print_Excel.aspx");
    }
}