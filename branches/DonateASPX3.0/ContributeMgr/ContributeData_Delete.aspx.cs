using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class ContributeMgr_ContributeData_Delete : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (Util.GetQueryString("Ser_No") == "")
            {
                Response.Redirect("ContributeList.aspx");
            }
            string Ser_No = Request.QueryString["Ser_No"].ToString();
            Del(Ser_No);
        }
    }
    public void Del(string Ser_No)
    {
        //****設定查詢****//
        string strSql = " select *  from ContributeData where Ser_No='" + Ser_No + "'";
        //****執行語法****//
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Contribute_Id = dr["Contribute_Id"].ToString();
        string Donor_Id = dr["Donor_Id"].ToString();
        string Goods_Name = dr["Goods_Name"].ToString();
        string Goods_Id = dr["Goods_Id"].ToString();
        string Goods_Amt = (Convert.ToInt64(dr["Goods_Amt"])).ToString();
        int Goods_Qty = int.Parse(dr["Goods_Qty"].ToString());

        //******刪除ContributeData的資料******//
        ContributeData_Del(Ser_No);

        //******修改Goods的資料******//
        Goods_Edit(Goods_Qty, Goods_Id);

        //******修改Donor的資料******//
        Donor_Edit(Goods_Amt, Donor_Id);

        //******修改Contribute的資料******//
        Contribute_Edit(Goods_Amt, Contribute_Id);

        SetSysMsg("『" + Goods_Name + "』刪除成功！");
        Response.Redirect("Contribute_Edit.aspx?Contribute_Id=" + Contribute_Id);
    }
    public void ContributeData_Del(string Ser_No)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = "delete from ContributeData where Ser_No=@Ser_No";
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
    public void Donor_Edit(string Goods_Amt, string Donor_Id)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Donor set ";
        strSql += "  Donate_TotalC = Donate_TotalC -'" + int.Parse(Goods_Amt) + "'";
        strSql += " where Donor_Id = @Donor_Id";
        dict.Add("Donor_Id", Donor_Id);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void Contribute_Edit(string Goods_Amt, string Contribute_Id)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Contribute set ";
        strSql += " Contribute_Amt = Contribute_Amt -'" + int.Parse(Goods_Amt) + "'";
        strSql += " where Contribute_Id = @Contribute_Id";
        dict.Add("Contribute_Id", Contribute_Id);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}