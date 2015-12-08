using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class SysMgr_Log_Delete : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (Util.GetQueryString("uid") == "")
            {
                Response.Redirect("Log.aspx");
            }
            string Uid = Request.QueryString["uid"].ToString();
            Del(Uid);
        }
    }
    public void Del(string strUid)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = "delete from Log where uid=@uid";
        dict.Add("uid", strUid);
        NpoDB.ExecuteSQLS(strSql, dict);

        SetSysMsg("刪除成功！");
        Response.Redirect("Log.aspx");
    }
}