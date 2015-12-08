using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;

public partial class ContributeMgr_ContributeData_Edit : BasePage
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
        strSql = " select *  from ContributeData where Ser_No='" + uid + "'";

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
        HFD_Donor_Id.Value = dr["Donor_Id"].ToString();
        HFD_Contribute_Id.Value = dr["Contribute_Id"].ToString();
        HFD_Goods_Id.Value = dr["Goods_Id"].ToString();
        tbxGoods_Name.Text = dr["Goods_Name"].ToString();
        tbxGoods_Qty.Text = dr["Goods_Qty"].ToString();
        HFD_Goods_Qty.Value = dr["Goods_Qty"].ToString();
        tbxGoods_Unit.Text = dr["Goods_Unit"].ToString();
        tbxGoods_Amt.Text = (Convert.ToInt64(dr["Goods_Amt"])).ToString();
        HFD_Goods_Amt.Value = (Convert.ToInt64(dr["Goods_Amt"])).ToString();
        if (dr["Goods_DueDate"].ToString() != "" && DateTime.Parse(dr["Goods_DueDate"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01" )
        {
            tbxGoods_DueDate.Text = DateTime.Parse(dr["Goods_DueDate"].ToString()).ToString("yyyy/MM/dd");
        }
        tbxGoods_Comment.Text = dr["Goods_Comment"].ToString();
    }
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            //******修改ContributeData的資料******//
            ContributeData_Edit();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            //Donor中的捐款資料修改
            Donor_Edit();
            SetSysMsg("物品類別修改成功！");
            Response.Write("<script>opener.location.href='Contribute_Edit.aspx?Contribute_Id=" + HFD_Contribute_Id.Value + "'</script>");
            Response.Write("<script language='javascript'>window.close();</script>");
        }
    }
    public void ContributeData_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update ContributeData set ";
        strSql += " Goods_Name = @Goods_Name";
        strSql += " ,Goods_Qty = @Goods_Qty";
        strSql += " ,Goods_Unit = @Goods_Unit";
        strSql += " ,Goods_Amt = @Goods_Amt";
        strSql += " ,Goods_DueDate = @Goods_DueDate";
        strSql += " ,Goods_Comment = @Goods_Comment";
        strSql += " where Ser_No = @Ser_No";

        dict.Add("Goods_Name", tbxGoods_Name.Text.Trim());
        dict.Add("Goods_Qty", tbxGoods_Qty.Text.Trim());
        dict.Add("Goods_Unit", tbxGoods_Unit.Text.Trim());
        dict.Add("Goods_Amt", tbxGoods_Amt.Text.Trim());
        dict.Add("Goods_DueDate", Util.DateTime2String(tbxGoods_DueDate.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty));
        dict.Add("Goods_Comment", tbxGoods_Comment.Text.Trim());
        dict.Add("Ser_No", HFD_Ser_No.Value);
        NpoDB.ExecuteSQLS(strSql, dict);

        if (int.Parse(tbxGoods_Amt.Text.Trim()) > int.Parse(HFD_Goods_Amt.Value))//修改後的金額大於原本金額
        {
            //******修改Contribute的資料******//
            //****變數宣告****//
            Dictionary<string, object> dict2 = new Dictionary<string, object>();
            //****設定SQL指令****//
            int Donate_TotalC = int.Parse(tbxGoods_Amt.Text.Trim()) - int.Parse(HFD_Goods_Amt.Value);

            string strSql2 = " update Contribute set ";
            strSql2 += " Contribute_Amt = Contribute_Amt +'" + Donate_TotalC + "'";
            strSql2 += " where Contribute_Id = @Contribute_Id";
            dict2.Add("Contribute_Id", HFD_Contribute_Id.Value);
            NpoDB.ExecuteSQLS(strSql2, dict2);

        }
        else if (int.Parse(tbxGoods_Amt.Text.Trim()) < int.Parse(HFD_Goods_Amt.Value))//修改後的金額小於原本金額
        {
            //******修改Contribute的資料******//
            //****變數宣告****//
            Dictionary<string, object> dict2 = new Dictionary<string, object>();
            //****設定SQL指令****//
            int Donate_TotalC = int.Parse(HFD_Goods_Amt.Value) - int.Parse(tbxGoods_Amt.Text.Trim());

            string strSql2 = " update Contribute set ";
            strSql2 += "  Contribute_Amt = Contribute_Amt -'" + Donate_TotalC + "'";
            strSql2 += " where Contribute_Id = @Contribute_Id";
            dict2.Add("Contribute_Id", HFD_Contribute_Id.Value);
            NpoDB.ExecuteSQLS(strSql2, dict2);

        }

        if (int.Parse(tbxGoods_Qty.Text.Trim()) > int.Parse(HFD_Goods_Qty.Value))//修改後的數量大於原本金額
        {
            //******修改Goods的資料******//
            Dictionary<string, object> dict_Goods = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Goods = " update Goods set ";

            int Goods_Qty = int.Parse(tbxGoods_Qty.Text.Trim()) - int.Parse(HFD_Goods_Qty.Value);
            strSql_Goods += "  Goods_Qty = Goods_Qty +'" + Goods_Qty + "'";
            strSql_Goods += " where Goods_Id = @Goods_Id";
            dict_Goods.Add("Goods_Id", HFD_Goods_Id.Value);
            NpoDB.ExecuteSQLS(strSql_Goods, dict_Goods);
        }
        else if (int.Parse(tbxGoods_Qty.Text.Trim()) < int.Parse(HFD_Goods_Qty.Value))//修改後的數量小於原本金額
        {
            //******修改Goods的資料******//
            Dictionary<string, object> dict_Goods = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Goods = " update Goods set ";

            int Goods_Qty = int.Parse(HFD_Goods_Qty.Value) - int.Parse(tbxGoods_Qty.Text.Trim());
            strSql_Goods += "  Goods_Qty = Goods_Qty -'" + Goods_Qty + "'";
            strSql_Goods += " where Goods_Id = @Goods_Id";
            dict_Goods.Add("Goods_Id", HFD_Goods_Id.Value);
            NpoDB.ExecuteSQLS(strSql_Goods, dict_Goods);
        }
    }
    public void Donor_Edit()
    {
        if (int.Parse(tbxGoods_Amt.Text.Trim()) > int.Parse(HFD_Goods_Amt.Value))
        {
            //******修改Donor的資料******//
            Dictionary<string, object> dict_Donor = new Dictionary<string, object>();//修改後的金額大於原本金額
            //****設定SQL指令****//
            string strSql_Donor = " update Donor set ";

            int Donate_TotalC = int.Parse(tbxGoods_Amt.Text.Trim()) - int.Parse(HFD_Goods_Amt.Value);
            strSql_Donor += "  Donate_TotalC = Donate_TotalC +'" + Donate_TotalC + "'";
            strSql_Donor += " where Donor_Id = @Donor_Id";
            dict_Donor.Add("Donor_Id", HFD_Donor_Id.Value);
            NpoDB.ExecuteSQLS(strSql_Donor, dict_Donor);
        }
        else if (int.Parse(tbxGoods_Amt.Text.Trim()) < int.Parse(HFD_Goods_Amt.Value))//修改後的金額小於原本金額
        {
            //******修改Donor的資料******//
            Dictionary<string, object> dict_Donor = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Donor = " update Donor set ";

            int Donate_TotalC = int.Parse(HFD_Goods_Amt.Value) - int.Parse(tbxGoods_Amt.Text.Trim());
            strSql_Donor += "  Donate_TotalC = Donate_TotalC -'" + Donate_TotalC + "'";
            strSql_Donor += " where Donor_Id = @Donor_Id";
            dict_Donor.Add("Donor_Id", HFD_Donor_Id.Value);
            NpoDB.ExecuteSQLS(strSql_Donor, dict_Donor);
        }
    }
}