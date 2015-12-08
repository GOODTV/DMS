using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
/*
 * 群組名稱設定
 * 資料庫: Goods
 * List頁面: GoodsList
 * 新增頁面: GoodsList_Add.aspx 
 * 修改頁面: GoodsList_Edit.aspx
 */
public partial class ContributeMgr_GoodsList :BasePage
{
    #region NpoGridView 處理換頁相關程式碼
    Button btnNextPage, btnPreviousPage, btnGoPage;
    HiddenField HFD_CurrentPage, HFD_CurrentQuerye;

    override protected void OnInit(EventArgs e)
    {
        CreatePageControl();
        base.OnInit(e);
    }
    private void CreatePageControl()
    {
        // Create dynamic controls here.
        btnNextPage = new Button();
        btnNextPage.ID = "btnNextPage";
        Form1.Controls.Add(btnNextPage);
        btnNextPage.Click += new System.EventHandler(btnNextPage_Click);

        btnPreviousPage = new Button();
        btnPreviousPage.ID = "btnPreviousPage";
        Form1.Controls.Add(btnPreviousPage);
        btnPreviousPage.Click += new System.EventHandler(btnPreviousPage_Click);

        btnGoPage = new Button();
        btnGoPage.ID = "btnGoPage";
        Form1.Controls.Add(btnGoPage);
        btnGoPage.Click += new System.EventHandler(btnGoPage_Click);

        HFD_CurrentPage = new HiddenField();
        HFD_CurrentPage.Value = "1";
        HFD_CurrentPage.ID = "HFD_CurrentPage";
        Form1.Controls.Add(HFD_CurrentPage);

        HFD_CurrentQuerye = new HiddenField();
        HFD_CurrentQuerye.Value = "Query";
        HFD_CurrentQuerye.ID = "HFHFD_CurrentQuerye";
        Form1.Controls.Add(HFD_CurrentQuerye);
    }
    protected void btnPreviousPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.MinusStringNumber(HFD_CurrentPage.Value);
        LoadFormData();
    }
    protected void btnNextPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.AddStringNumber(HFD_CurrentPage.Value);
        LoadFormData();
    }
    protected void btnGoPage_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    #endregion NpoGridView 處理換頁相關程式碼
    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "GoodsList";
        //權控處理
        AuthrityControl();
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
        if (!IsPostBack)
        {
            LoadDropDownListData();
            LoadFormData();          
        }
    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_Query", btnQuery);
        Authrity.CheckButtonRight("_Print", btnPrint);
        Authrity.CheckButtonRight("_Print", btnToxls);
        Authrity.CheckButtonRight("_AddNew", btnAdd);
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //庫存管理
        ddlGoods_IsStock.Items.Insert(0, new ListItem("", ""));
        ddlGoods_IsStock.Items.Insert(1, new ListItem("是", "是"));
        ddlGoods_IsStock.Items.Insert(2, new ListItem("否", "否"));
        ddlGoods_IsStock.SelectedIndex = 0;

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
        if (tbxGoods_Id.Text.Trim() != "")
        {
            strSql += " and G.Goods_Id like '%" + tbxGoods_Id.Text.Trim() + "%'";
        }
        if (tbxGoods_Name.Text.Trim() != "")
        {
            strSql += " and G.Goods_Name like '%" + tbxGoods_Name.Text.Trim() + "%'";
        }

        if (ddlGoods_IsStock.SelectedIndex == 1)
        {
            strSql += " and G.Goods_IsStock = 'Y'";
        }
        else if(ddlGoods_IsStock.SelectedIndex == 2)
        {
            strSql += " and G.Goods_IsStock = 'N'";
        }
        if (ddlGoods_Property.SelectedIndex != 0)
        {
            strSql += " and L.Linked_Name = '" + ddlGoods_Property.SelectedItem.Text + "'";
        }
        if (ddlGoods_Type.SelectedIndex != 0 && ddlGoods_Property.SelectedIndex != 0)
        {
            strSql += " and L2.Linked2_Name = '" + ddlGoods_Type.SelectedValue + "'";
        }
        if (cbxGoods_Qty.Checked == true)
        {
            strSql += " and G.Goods_Qty > 0";
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
            npoGridView.Keys.Add("物品代號");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            npoGridView.EditLink = Util.RedirectByTime("GoodsList_Edit.aspx", "Goods_Id=");
            lblGridList.Text = npoGridView.Render();
        }

        Session["strSql"] = strSql;
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Response.Redirect("GoodsList_Print_Excel.aspx");
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("GoodsList_Add.aspx"));
    }
    //dropdownlist選物品性質自動帶入物品類別
    protected void ddlGoods_Property_SelectedIndexChanged(object sender, EventArgs e)
    {
        //物品類別
        Util.FillDropDownList(ddlGoods_Type, Util.GetDataTable("Linked2", "Linked_Id", ddlGoods_Property.SelectedValue, "", ""), "Linked2_Name", "Linked2_Name", true);
        ddlGoods_Type.SelectedIndex = 0;
    }
}