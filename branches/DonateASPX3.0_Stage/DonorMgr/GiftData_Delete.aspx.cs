using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonorMgr_GiftData_Delete : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (Util.GetQueryString("Ser_No") == "")
            {
                Response.Redirect("../ContributeMgr/ContributeList.aspx");
            }
            string Ser_No = Request.QueryString["Ser_No"].ToString();
            Del(Ser_No);
        }
    }
    public void Del(string Ser_No)
    {
        //****設定查詢****//
        string strSql = " select *  from GiftData where Ser_No='" + Ser_No + "'";
        //****執行語法****//
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Gift_Id = dr["Gift_Id"].ToString();
        string Donor_Id = dr["Donor_Id"].ToString();
        string Goods_Name = dr["Goods_Name"].ToString();
        string Goods_Id = dr["Goods_Id"].ToString();
        int Goods_Qty = int.Parse(dr["Goods_Qty"].ToString());

        //******刪除GiftData的資料******//
        GiftData_Del(Ser_No);

        //******修改Goods的資料******//
        Goods_Edit(Goods_Qty, Goods_Id);

        SetSysMsg("『" + Goods_Name + "』刪除成功！");
        Response.Redirect("Gift_Edit.aspx?Gift_Id=" + Gift_Id);
    }
    public void GiftData_Del(string Ser_No)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = "delete from GiftData where Ser_No=@Ser_No";
        dict.Add("Ser_No", Ser_No);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void Goods_Edit(int Goods_Qty, string Goods_Id)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Goods set ";
        strSql += "  Goods_Qty = Goods_Qty -'" + Goods_Qty + "'";
        strSql += " where Goods_Id = @Goods_Id";
        dict.Add("Goods_Id", Goods_Id);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}