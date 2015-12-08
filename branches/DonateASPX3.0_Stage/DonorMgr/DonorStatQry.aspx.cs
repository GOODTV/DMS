using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonorMgr_Donor_Stat_Qry : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "DonorStatQry";
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
        ResponseHelper.Redirect(Util.RedirectByTime("DonorStatQry_Print_Excel.aspx?Is_Abroad=" + cbxIs_Abroad.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress.Checked + "&Sex=" + cbxSex.Checked + "&") ,"_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
    }
    public static class ResponseHelper
    {
        public static void Redirect(string url, string target, string windowFeatures)
        {

            Page page = (Page)HttpContext.Current.Handler;

            if (page == null)
            {
                throw new InvalidOperationException("Cannot redirect to new window outside Page context.");
            }
            url = page.ResolveClientUrl(url);

            string script;
            if (!String.IsNullOrEmpty(windowFeatures))
            {
                script = @"window.open(""{0}"", ""{1}"", ""{2}"");";
            }
            else
            {
                script = @"window.open(""{0}"", ""{1}"");";
            }
            script = String.Format(script, url, target, windowFeatures);
            ScriptManager.RegisterStartupScript(page, typeof(Page), "Redirect", script, true);
        }
    }
}