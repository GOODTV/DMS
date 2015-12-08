using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;

public partial class SysMgr_CaseCode_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            HFD_Uid.Value = Util.GetQueryString("Uid");
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
        uid = HFD_Uid.Value;

        //****設定查詢****//
        strSql = " select * from CaseCode where Uid='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);

        DataRow dr = dt.Rows[0];

        tbxCaseFolder.Text = dr["GroupName"].ToString();
        tbxCaseItems.Text = dr["CaseName"].ToString();
        tbxCaseID.Text = dr["CaseID"].ToString();
    }
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        string strRet = CheckData();
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
                CaseCode_Edit();
                flag = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (flag == true)
            {
                SetSysMsg("代碼選項修改成功！");
                Response.Write("<script>opener.location.href='CaseCode.aspx?Uid=" + HFD_Uid.Value + "'</script>");
                Response.Write("<script language='javascript'>window.close();</script>");
            }
        }
    }
    //檢核欄位有沒有重複
    private string CheckData()
    {
        string strRet = "";
        string strSql = "select * from CaseCode where CaseName = @CaseName and GroupName = @GroupName and Uid != @Uid";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("CaseName", tbxCaseItems.Text);
        dict.Add("GroupName", tbxCaseFolder.Text);
        dict.Add("Uid",HFD_Uid.Value);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count > 0)
        {
            strRet += "選項 ";
        }
        strSql = "select * from CaseCode where CaseID = @CaseID and GroupName = @GroupName and Uid != @Uid";
        dict.Clear();
        dict.Add("CaseID", tbxCaseID.Text);
        dict.Add("GroupName", tbxCaseFolder.Text);
        dict.Add("Uid", HFD_Uid.Value);
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count > 0)
        {
            strRet += "排序 ";
        }
        if (strRet != "")
        {
            strRet += "欄位不可重複！";
        }
        return strRet;
    }
    public void CaseCode_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update CaseCode set ";
        strSql += "  CaseName = @CaseName";
        strSql += ", CaseID = @CaseID";
        strSql += " where Uid = @Uid";

        dict.Add("CaseName", tbxCaseItems.Text);
        dict.Add("CaseID", tbxCaseID.Text);
        dict.Add("Uid", HFD_Uid.Value);

        NpoDB.ExecuteSQLS(strSql, dict);
    }
}