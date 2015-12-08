using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;
using System.Web.UI.HtmlControls;

public partial class ContributeMgr_LinkedList_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Linked_Id.Value = Util.GetQueryString("Linked_Id");
            Form_DataBind();
        }
    }
    //----------------------------------------------------------------------
    //帶入資料
    public void Form_DataBind()
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;

        //****變數設定****//
        uid = HFD_Linked_Id.Value;

        //****設定查詢****//
        strSql = " select *  from Linked where Linked_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);

        DataRow dr = dt.Rows[0];

        tbxLinked_Name.Text = dr["Linked_Name"].ToString();
        tbxLinked_Seq.Text = dr["Linked_Seq"].ToString();
    }
    protected void btnEdit_Click(object sender, EventArgs e)
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
                Linked_Edit();
                flag = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (flag == true)
            {
                SetSysMsg("物品類別修改成功！");
                Response.Write("<script>opener.location.href='LinkedList.aspx?Linked_Id=" + HFD_Linked_Id.Value + "'</script>");
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
            strRet += "欄位不得為空白！";
        }
        return strRet;
    }
    public void Linked_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Linked set ";
        strSql += "  Linked_Name = @Linked_Name";
        strSql += ", Linked_Seq = @Linked_Seq";
        strSql += " where Linked_Id = @Linked_Id";

        dict.Add("Linked_Name", tbxLinked_Name.Text.Trim());
        dict.Add("Linked_Seq", tbxLinked_Seq.Text.Trim());
        dict.Add("Linked_Id", HFD_Linked_Id.Value);

        NpoDB.ExecuteSQLS(strSql, dict);
    }
}