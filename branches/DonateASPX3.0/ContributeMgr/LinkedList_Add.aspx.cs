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

public partial class ContributeMgr_LinkedList_Add : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            tbxLinked_Seq.Text = CaseUtil.GetNewMaxLinked2_Seq();
        }
        if (Session["Linked_Id"] == "0")
        {
            btnCancle.Visible = false;
        }
    }
    protected void btnSave_Click(object sender, EventArgs e)
    {
        string strRet = CheckFieldMustFillBasic();
        if (strRet != "")
        {
            ShowSysMsg(strRet);
            return;
        }
        else
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
                SetSysMsg("新增物品類別成功，請繼續新增「物品類別」！");
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
    }
    //-------------------------------------------------------------------------
    //檢核基本資料必填欄位
    public string CheckFieldMustFillBasic()
    {
        string strRet = "";

        if (tbxLinked_Name.Text.Trim() == "")
        {
            strRet += "物品類別 ";
        }
        if (tbxLinked_Seq.Text.Trim() == "")
        {
            strRet += "排序 ";
        }
        if (tbxLinked_Name.Text.Trim() == "" || tbxLinked_Seq.Text.Trim() == "")
        {
            strRet += "欄位不可為空白！";
        }
        return strRet;
    }
    public void Linked_AddNew()
    {
        string strSql = "insert into  Linked\n";
        strSql += "( Linked_Type, Linked_Name, Linked_Seq ) values\n";
        strSql += "(@Linked_Type ,@Linked_Name, @Linked_Seq)";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Linked_Type", "Contribute");
        dict.Add("Linked_Name", tbxLinked_Name.Text.Trim());
        dict.Add("Linked_Seq", tbxLinked_Seq.Text.Trim());
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}