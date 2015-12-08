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

public partial class ContributeMgr_LinkedItem_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Linked_Id.Value = Util.GetQueryString("Linked_Id");
            HFD_Ser_No.Value = Util.GetQueryString("Ser_No");
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
        uid = HFD_Ser_No.Value;

        //****設定查詢****//
        strSql = " select *  from Linked2 where Ser_No='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);

        DataRow dr = dt.Rows[0];

        tbxLinked2_Name.Text = dr["Linked2_Name"].ToString();
        tbxLinked2_Seq.Text = dr["Linked2_Seq"].ToString();
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
                Linked2_Edit();
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

        if (tbxLinked2_Name.Text.Trim() == "")
        {
            strRet += "物品類別 ";
        }
        if (tbxLinked2_Seq.Text.Trim() == "")
        {
            strRet += "排序 ";
        }
        if (tbxLinked2_Name.Text.Trim() == "" || tbxLinked2_Seq.Text.Trim() == "")
        {
            strRet += "欄位不得為空白！";
        }
        return strRet;
    }
    public void Linked2_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Linked2 set ";
        strSql += "  Linked2_Name = @Linked2_Name";
        strSql += ", Linked2_Seq = @Linked2_Seq";
        strSql += " where Ser_No = @Ser_No";

        dict.Add("Linked2_Name", tbxLinked2_Name.Text.Trim());
        dict.Add("Linked2_Seq", tbxLinked2_Seq.Text.Trim());
        dict.Add("Ser_No", HFD_Ser_No.Value);

        NpoDB.ExecuteSQLS(strSql, dict);
    }
}