using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class ContributeMgr_GoodsList_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Goods_Id");
            LoadDropDownListData();
            Form_DataBind();
        }
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //物品性質
        Util.FillDropDownList(ddlGoods_Property, Util.GetDataTable("Linked", "Linked_Type", "contribute", "", ""), "Linked_Name", "Linked_Id", true);
        //ddlGoods_Property.SelectedIndex = 0;
    }
    //dropdownlist選物品性質自動帶入物品類別
    protected void ddlGoods_Property_SelectedIndexChanged(object sender, EventArgs e)
    {
        PropertyToType();
    }
    private void PropertyToType()
    {
        //物品類別
        Util.FillDropDownList(ddlGoods_Type, Util.GetDataTable("Linked2", "Linked_Id", ddlGoods_Property.SelectedValue, "", ""), "Linked2_Name", "Linked2_Name", true);
        //ddlGoods_Type.SelectedIndex = 0;
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
        strSql = " select *  from Goods where Goods_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("GoodsInfo.aspx");

        DataRow dr = dt.Rows[0];

        //物品代號
        tbxGoods_Id.Text = dr["Goods_Id"].ToString().Trim();
        //物品名稱
        tbxGoods_Name.Text = dr["Goods_Name"].ToString().Trim();

        //物品性質
        ddlGoods_Property.Items.FindByText("").Selected = false;
        ddlGoods_Property.Items.FindByText(dr["Goods_Property"].ToString().Trim()).Selected = true;
        //物品類別
        PropertyToType();
        ddlGoods_Type.Text = dr["Goods_Type"].ToString().Trim();
        //庫存單位
        tbxGoods_Unit.Text = dr["Goods_Unit"].ToString().Trim();
        //是否庫存管理
        if (dr["Goods_IsStock"].ToString().Trim() == "Y")
        {
            rblGoods_IsStock.SelectedValue = "Y";
        }
        else if (dr["Goods_IsStock"].ToString().Trim() == "N")
        {
            rblGoods_IsStock.SelectedValue = "N";
        }
        //現有庫存量
        tbxGoods_Qity.Text = dr["Goods_Qty"].ToString().Trim();
        //備註
        tbxRemark.Text = dr["Remark"].ToString().Trim();
    }
    //----------------------------------------------------------------------
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            Goods_Edit();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("物品主檔資料修改成功！");
            Response.Redirect(Util.RedirectByTime("GoodsList.aspx"));
        }
    }
    protected void btnDel_Click(object sender, EventArgs e)
    {
        string strSql = "delete from Goods where Goods_Id=@Goods_Id";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Goods_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);

        SetSysMsg("資料刪除成功！");
        Response.Redirect(Util.RedirectByTime("GoodsList.aspx"));
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("GoodsList.aspx"));
    }
    public void Goods_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update Goods set ";

        strSql += "  Goods_Name = @Goods_Name";
        strSql += ", Goods_Property = @Goods_Property";
        strSql += ", Goods_Type = @Goods_Type";
        strSql += ", Goods_Unit = @Goods_Unit";
        strSql += ", Goods_IsStock = @Goods_IsStock";
        strSql += ", Remark = @Remark";
        strSql += " where Goods_Id = @Goods_Id";

        dict.Add("Goods_Name", tbxGoods_Name.Text.Trim());
        dict.Add("Goods_Property", ddlGoods_Property.SelectedItem.Text);
        dict.Add("Goods_Type", ddlGoods_Type.SelectedItem.Text);
        dict.Add("Goods_Unit", tbxGoods_Unit.Text.Trim());
        dict.Add("Goods_IsStock", rblGoods_IsStock.SelectedValue);
        dict.Add("Remark", tbxRemark.Text.Trim());
        dict.Add("Goods_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}