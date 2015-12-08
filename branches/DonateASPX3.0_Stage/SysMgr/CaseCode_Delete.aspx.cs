using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class SysMgr_CaseCode_Delete : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (Util.GetQueryString("Uid") == "")
            {
                Response.Redirect("CaseCode.aspx");
            }
            string Uid = Request.QueryString["Uid"].ToString();
            Del(Uid);
        }

    }
    public void Del(string strUid)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = "delete from CaseCode where Uid=@Uid";
        dict.Add("Uid", strUid);
        NpoDB.ExecuteSQLS(strSql, dict);

        SetSysMsg("代碼選項刪除成功！");
        Response.Redirect("CaseCode.aspx");
    }
}