using System;
using System.Data;
using System.Collections.Generic;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class ContributeMgr_LinkedItem_Delete : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (Util.GetQueryString("Ser_No") == "")
            {
                Response.Redirect("LinkedList.aspx");
            }
            string Ser_No = Request.QueryString["Ser_No"].ToString();
            string Linked_Id = Request.QueryString["Linked_Id"].ToString();
            Del(Ser_No, Linked_Id);
        }
    }
    public void Del(string Ser_No, string Linked_Id)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = "delete from Linked2 where Ser_No=@Ser_No";
        dict.Add("Ser_No", Ser_No);
        NpoDB.ExecuteSQLS(strSql, dict);

        SetSysMsg("物品類別刪除成功！");
        Response.Redirect("LinkedList.aspx?Linked_Id=" + Linked_Id); 
    }
}