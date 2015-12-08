using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class ContributeMgr_Goods_Show : BasePage
{
    string srcCtl = "";
    string srcForm = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            srcCtl = Request.Params["SrcCtl"];
            srcForm = Request.Params["SrcForm"];
            LoadDropDownListData();
            LoadFormData(); 
        }
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //物品性質
        Util.FillDropDownList(ddlGoods_Property, Util.GetDataTable("Linked", "Linked_Type", "contribute", "", ""), "Linked_Name", "Linked_Id", true);
        ddlGoods_Property.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = "Select Goods_Id as [物品代號], Goods_Name as [物品名稱], Goods_Property as [物品性質], Goods_Type as [物品類別],\n";
        strSql += " Goods_Qty as [現有庫存量], Goods_Unit as [庫存單位], (Case When Goods_IsStock='Y' Then 'V' Else '' End) as [庫存管理]\n";
        strSql += " From GOODS G Left Join LINKED L On G.Goods_Property=L.Linked_Name \n";
        strSql += " Left Join LINKED2 L2 On G.Goods_Type=L2.Linked2_Name \n";
        strSql += " Where L.Linked_Type='contribute' ";
        
        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (ddlGoods_Property.SelectedIndex != 0)
        {
            strSql += " and L.Linked_Name = '" + ddlGoods_Property.SelectedItem.Text + "'";
        }
        if (ddlGoods_Type.SelectedIndex != 0 && ddlGoods_Property.SelectedIndex != 0)
        {
            strSql += " and L2.Linked2_Name = '" + ddlGoods_Type.SelectedValue + "'";
        }
        dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.ShowPage = false;
            npoGridView.Keys.Add("物品代號");
            npoGridView.Keys.Add("物品名稱");
            npoGridView.Keys.Add("庫存單位");
            //在前端設置JS語法並傳入KeyValue並執行
            npoGridView.EditClick = "ReturnOpener";
            lblGridList.Text = npoGridView.Render();
        }
    }
    protected void ddlGoods_Property_SelectedIndexChanged(object sender, EventArgs e)
    {
        //物品類別
        Util.FillDropDownList(ddlGoods_Type, Util.GetDataTable("Linked2", "Linked_Id", ddlGoods_Property.SelectedValue, "", ""), "Linked2_Name", "Linked2_Name", true);
        ddlGoods_Type.SelectedIndex = 0;
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
}