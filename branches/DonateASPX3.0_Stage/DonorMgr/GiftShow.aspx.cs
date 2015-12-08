using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonorMgr_Gift_Show : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            LoadDropDownListData();
            LoadFormData();
        }
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //物品性質
        Util.FillDropDownList(ddlGifts_Property, Util.GetDataTable("Linked", "Linked_Type", "Gift", "", ""), "Linked_Name", "Linked_Id", true);
        ddlGifts_Property.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = " Select Linked.Linked_Name as [公關贈品品項] ,Linked2_Name as [物品名稱] , Linked2.Ser_No as [物品編號]";
        strSql += " From LINKED2 Join Linked On Linked.Linked_Id=LINKED2.Linked_Id Where Linked.Linked_Type='gift' ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (ddlGifts_Property.SelectedIndex != 0)
        {
            strSql += " and Linked.Linked_Name = '" + ddlGifts_Property.SelectedItem.Text + "'";
        }
        strSql += " Order By Linked2.Linked_Id,Linked2_Seq ";

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
            npoGridView.Keys.Add("物品編號");
            npoGridView.Keys.Add("物品名稱");
            npoGridView.DisableColumn.Add("物品編號");

            //在前端設置JS語法並傳入KeyValue並執行
            npoGridView.EditClick = "ReturnOpener";
            lblGridList.Text = npoGridView.Render();
        }
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
}