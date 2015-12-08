using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class ContributeMgr_GoodsList_Add : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            LoadDropDownListData();
        }
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //物品性質
        Util.FillDropDownList(ddlGoods_Property, Util.GetDataTable("Linked", "Linked_Type", "contribute", "", ""), "Linked_Name", "Linked_Id", true);
        ddlGoods_Property.SelectedIndex = 0;
    }
    //dropdownlist選物品性質自動帶入物品類別
    protected void ddlGoods_Property_SelectedIndexChanged(object sender, EventArgs e)
    {
        //物品類別
        Util.FillDropDownList(ddlGoods_Type, Util.GetDataTable("Linked2", "Linked_Id", ddlGoods_Property.SelectedValue, "", ""), "Linked2_Name", "Linked2_Name", true);
        ddlGoods_Type.SelectedIndex = 0;
    }
    //----------------------------------------------------------------------
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            Goods_AddNew();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("物品主檔資料新增成功！");
            Response.Redirect(Util.RedirectByTime("GoodsList.aspx"));
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("GoodsList.aspx"));
    }
    public void Goods_AddNew()
    {
        string strSql = "insert into  Goods\n";
        strSql += "( Goods_Id, Goods_Name, Goods_Property, Goods_Type, Goods_IsStock, Goods_Qty, Goods_Unit,Remark) values\n";
        strSql += "( @Goods_Id,@Goods_Name,@Goods_Property,@Goods_Type,@Goods_IsStock,@Goods_Qty,@Goods_Unit,@Remark) ";
        strSql += "\n";
        strSql += "select @@IDENTITY";


        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Goods_Id", tbxGoods_Id.Text.Trim());
        dict.Add("Goods_Name", tbxGoods_Name.Text.Trim());
        dict.Add("Goods_Property", ddlGoods_Property.SelectedItem.Text);
        dict.Add("Goods_Type", ddlGoods_Type.SelectedItem.Text);
        dict.Add("Goods_IsStock", rblGoods_IsStock.SelectedValue);
        dict.Add("Goods_Qty", "0");
        dict.Add("Goods_Unit", tbxGoods_Unit.Text.Trim());
        dict.Add("Remark", tbxRemark.Text.Trim());


        //******勸募活動編號******//
        string strSql2 = @"select Goods_Id from Goods";
        //****執行查詢勸募活動****//
        DataTable dt2 = NpoDB.QueryGetTable(strSql2);
        DataRow dr;
        if (dt2.Rows.Count > 0)
        {
            int count = dt2.Rows.Count;
            for (int i = 0; i < count; i++)
            {
                dr = dt2.Rows[i];
                if (tbxGoods_Id.Text.Trim() == dr["Goods_Id"].ToString())
                {
                    ShowSysMsg("您輸入的物品代號已經存在！");
                    return;
                }
            }
            NpoDB.ExecuteSQLS(strSql, dict);
        }
        else
        {
            NpoDB.ExecuteSQLS(strSql, dict);
        }
    }
}