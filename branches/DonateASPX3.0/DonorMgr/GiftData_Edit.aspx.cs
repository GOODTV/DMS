using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;

public partial class DonorMgr_GiftData_Edit : BasePage
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
        strSql = @" select Donor_Id,Gift_Id,Goods_Id,L.Linked2_Name,Goods_Qty,Goods_Comment  from GiftData G 
                    right join Linked2 L on G.Goods_Id = L.Ser_No where G.Ser_No='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);

        DataRow dr = dt.Rows[0];
        HFD_Donor_Id.Value = dr["Donor_Id"].ToString();
        HFD_Gift_Id.Value = dr["Gift_Id"].ToString();
        HFD_Goods_Id.Value = dr["Goods_Id"].ToString();
        tbxGoods_Name.Text = dr["Linked2_Name"].ToString();
        tbxGoods_Qty.Text = dr["Goods_Qty"].ToString();
        HFD_Goods_Qty.Value = dr["Goods_Qty"].ToString();
        tbxGoods_Comment.Text = dr["Goods_Comment"].ToString();
    }
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            //******修改GiftData的資料******//
            GiftData_Edit();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("物品類別修改成功！");
            Response.Write("<script>opener.location.href='Gift_Edit.aspx?Gift_Id=" + HFD_Gift_Id.Value + "'</script>");
            Response.Write("<script language='javascript'>window.close();</script>");
        }
    }
    public void GiftData_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update GiftData set ";
        strSql += " Goods_Name = @Goods_Name";
        strSql += " ,Goods_Qty = @Goods_Qty";
        strSql += " ,Goods_Comment = @Goods_Comment";
        strSql += " where Ser_No = @Ser_No";

        dict.Add("Goods_Name", tbxGoods_Name.Text.Trim());
        dict.Add("Goods_Qty", tbxGoods_Qty.Text.Trim());
        dict.Add("Goods_Comment", tbxGoods_Comment.Text.Trim());
        dict.Add("Ser_No", HFD_Ser_No.Value);
        NpoDB.ExecuteSQLS(strSql, dict);

    }
    public string Check(string Goods_Name, int tbxGoods_Qty)
    {
        string strRet = "";
        string strSql = " select *  from Linked2 where Linked2_Name='" + Goods_Name + "'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        int Goods_Qty = int.Parse(dr["Goods_Qty"].ToString());//庫存量
        if (Goods_Qty - tbxGoods_Qty < 0)
        {
            strRet += "『" + Goods_Name + "』 ";
        }
        return strRet;
    }
}