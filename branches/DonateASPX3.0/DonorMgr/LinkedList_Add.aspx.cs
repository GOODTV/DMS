using System;
using System.Data;
using System.Configuration;
using System.Collections.Generic;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

public partial class DonorMgr_Linked_List_Add : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            tbxLinked_Seq.Text = CaseUtil.GetNewMaxLinked_Seq();
        }
        if (Session["Linked_Id"] == "0")
        {
            btnCancle.Visible = false;
        }
    }
    protected void btnSave_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            Linked_AddNew();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("新增贈品品項成功，請繼續新增「公關贈品品項類別」！");
            string Linked_Id = "";

            string strSql = @"select Linked_Id from Linked ";
            strSql += " where Linked_Name='" + tbxLinked_Name.Text.Trim() + "'";
            DataTable dt = NpoDB.GetDataTableS(strSql, null);
            if (dt.Rows.Count > 0)
            {
                Linked_Id = dt.Rows[0]["Linked_Id"].ToString();
            }
            if (Session["Linked_Id"] == "0")
            {
                Session["Linked_Id"] = "";
                Response.Redirect(Util.RedirectByTime("LinkedList.aspx"));
            }
            Response.Write("<script>opener.location.href='LinkedList.aspx?Linked_Id=" + Linked_Id + "'</script>");
            Response.Write("<script language='javascript'>window.close();</script>");
            }
    }
    public void Linked_AddNew()
    {
        string strSql = "insert into  Linked\n";
        strSql += "( Linked_Type, Linked_Name, Linked_Seq ) values\n";
        strSql += "(@Linked_Type ,@Linked_Name, @Linked_Seq)";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        // 2014/4/8 修正新增時存入錯誤的公關贈品品項的識別
        //dict.Add("Linked_Type", "Contribute");
        dict.Add("Linked_Type", "gift");
        dict.Add("Linked_Name", tbxLinked_Name.Text.Trim());
        dict.Add("Linked_Seq", tbxLinked_Seq.Text.Trim());
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}