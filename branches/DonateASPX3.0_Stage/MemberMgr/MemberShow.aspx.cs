using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class MemberMgr_MemberShow : BasePage
{
    string srcCtl = "";
    string srcForm = "";
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
        //身分別
        Util.FillDropDownList(ddlDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "", ""), "CaseName", "CaseName", true);
        ddlDonor_Type.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = "Select Donor_Id as [編號], Donor_Name as [會員姓名], Donor_Type as [身份別], Cellular_Phone as [聯絡電話],\n";
        strSql += "  (Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) as [通訊地址]\n";
        strSql += " From Donor Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode \n";
        strSql += " Where DeleteDate is null ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (tbxDonor_Name.Text.Trim()!="")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%'";
            strSql += " and Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%'";
            
        }
        if (tbxDonor_Id.Text.Trim() != "")
        {
            strSql += " and Donor_Id = @Donor_Id";
            dict.Add("Donor_Id", tbxDonor_Id.Text.Trim());
        } 
        if (ddlDonor_Type.SelectedIndex!=0)
        {
            strSql += " and Donor_Type = '" + ddlDonor_Type.SelectedItem.Text + "'";
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
            npoGridView.Keys.Add("編號");
            npoGridView.Keys.Add("會員姓名");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
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