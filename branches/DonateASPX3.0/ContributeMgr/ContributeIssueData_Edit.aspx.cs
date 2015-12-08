using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;

public partial class ContributeMgr_ContributeIssueData_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
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
        strSql = " select *  from Contribute_IssueData where Ser_No='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);

        DataRow dr = dt.Rows[0];
        if (dr["Goods_Id"].ToString() != "")
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "ReadOnly();", true);
        }
        else
        {
            lblWarm.Visible = false;
        }
        HFD_Issue_Id.Value = dr["Issue_Id"].ToString();
        HFD_Goods_Id.Value = dr["Goods_Id"].ToString();
        tbxGoods_Name.Text = dr["Goods_Name"].ToString();
        tbxGoods_Qty.Text = dr["Goods_Qty"].ToString();
        HFD_Goods_Qty.Value = dr["Goods_Qty"].ToString();
        tbxGoods_Unit.Text = dr["Goods_Unit"].ToString();
        tbxGoods_Comment.Text = dr["Goods_Comment"].ToString();
    }
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        string strRet = CheckStock();
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
                Contribute_IssueData_Edit();
                flag = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (flag == true)
            {
                SetSysMsg("修改成功！");
                Response.Write("<script>opener.location.href='ContributeIssue_Edit.aspx?Issue_Id=" + HFD_Issue_Id.Value + "'</script>");
                Response.Write("<script language='javascript'>window.close();</script>");
            }
        }
    }
    public void Contribute_IssueData_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Contribute_IssueData set ";
        strSql += " Goods_Name = @Goods_Name";
        strSql += " ,Goods_Qty = @Goods_Qty";
        strSql += " ,Goods_Unit = @Goods_Unit";
        strSql += " ,Goods_Comment = @Goods_Comment";
        strSql += " where Ser_No = @Ser_No";

        dict.Add("Goods_Name", tbxGoods_Name.Text.Trim());
        dict.Add("Goods_Qty", tbxGoods_Qty.Text.Trim());
        dict.Add("Goods_Unit", tbxGoods_Unit.Text.Trim());
        dict.Add("Goods_Comment", tbxGoods_Comment.Text.Trim());
        dict.Add("Ser_No", HFD_Ser_No.Value);

        NpoDB.ExecuteSQLS(strSql, dict);

        //******修改Goods的資料******//
        Dictionary<string, object> dict_Goods = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql_Goods = " update Goods set ";
        if (int.Parse(tbxGoods_Qty.Text.Trim()) > int.Parse(HFD_Goods_Qty.Value))
        {
            int Goods_Qty = int.Parse(tbxGoods_Qty.Text.Trim()) - int.Parse(HFD_Goods_Qty.Value);
            strSql_Goods += "  Goods_Qty = Goods_Qty -'" + Goods_Qty + "'";
        }
        else if (int.Parse(tbxGoods_Qty.Text.Trim()) < int.Parse(HFD_Goods_Qty.Value))
        {
            int Goods_Qty = int.Parse(HFD_Goods_Qty.Value) - int.Parse(tbxGoods_Qty.Text.Trim());
            strSql_Goods += "  Goods_Qty = Goods_Qty +'" + Goods_Qty + "'";
        }
        strSql_Goods += " where Goods_Id = @Goods_Id";
        dict_Goods.Add("Goods_Id", HFD_Goods_Id.Value);
        NpoDB.ExecuteSQLS(strSql_Goods, dict_Goods);
    }
    public string CheckStock()
    {
        string strRet = "";
        strRet = Check(HFD_Goods_Id.Value, int.Parse(tbxGoods_Qty.Text.Trim()));
        if (strRet != "")
        {
            strRet += "的物品庫存量不足無法領用";
        }
        return strRet;
    }
    public string Check(string Goods_Id, int tbxGoods_Qty)
    {
        string strRet = "";
        string strSql = " select *  from Goods where Goods_Id='" + Goods_Id + "'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        int Goods_Qty = int.Parse(dr["Goods_Qty"].ToString());//庫存量
        string Goods_Name = dr["Goods_Name"].ToString();
        if (Goods_Qty < tbxGoods_Qty)
        {
            strRet += "『" + Goods_Name + "』 ";
        }
        return strRet;
    }
}