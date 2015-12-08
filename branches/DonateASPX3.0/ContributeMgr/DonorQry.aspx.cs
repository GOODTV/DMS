using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class ContributeMgr_DonorQry : BasePage
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

    protected void Page_Load(object sender, EventArgs e)
    {
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            LoadDropDownListData();
            lblGridList.Text = "** 請先輸入查詢條件 **";
        }
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        Util.FillDropDownList(ddlDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "", ""), "CaseName", "CaseID", false);
        ddlDonor_Type.Items.Insert(0, new ListItem("請選擇", "請選擇"));
        ddlDonor_Type.SelectedIndex = 0;

        Util.FillDropDownList(ddlCity, Util.GetDataTable("CodeCity", "ParentCityID", "0", "", ""), "Name", "ZipCode", false);
        ddlCity.Items.Insert(0, new ListItem("縣 市", "縣 市"));
        ddlCity.SelectedIndex = 0;

        ddlArea.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    //Dropdownlist選城市自動帶入地區
    protected void ddlCity_SelectedIndexChanged(object sender, EventArgs e)
    {
        Util.FillDropDownList(ddlArea, Util.GetDataTable("CodeCity", "ParentCityID", ddlCity.SelectedValue, "", ""), "Name", "ZipCode", false);
        ddlArea.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql;
        strSql = "select Donor_Id as [編號], Donor_Name + (case when NickName <> '' then '('+NickName+')' else '' end) as [捐款人],\n";
        strSql += " Category as [類別], Donor_Type as [身分別], Cellular_Phone as [手機], Tel_Office as [電話日], Email as [Email], IDNo as [身分證], \n";
        strSql += " Contactor as [聯絡人], (Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) as [通訊地址], CONVERT(VARCHAR(10) , Last_DonateDate, 111 ) as [最近捐款日期]\n";
        strSql += " from Donor Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode\n";
        strSql += " where DeleteDate is null";

        Dictionary<string, object> dict = new Dictionary<string, object>();

        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and (Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%' or NickName like '%" + tbxDonor_Name.Text.Trim() + "%')";
            strSql += " and (Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%' or NickName like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%')";
        }
        if (tbxDonor_Id.Text.Trim() != "")
        {
            strSql += " and Donor_Id like '%" + tbxDonor_Id.Text.Trim() + "%' ";
        }
        if (tbxIDNo.Text.Trim() != "")
        {
            strSql += " and IDNo like '%" + tbxIDNo.Text.Trim() + "%' ";
        }
        if (tbxCellular_Phone.Text.Trim() != "")
        {
            strSql += " and (Tel_Office Like '%" + tbxCellular_Phone.Text.Trim() + "%' Or Tel_Home Like '%" + tbxCellular_Phone.Text.Trim() + "%' Or Cellular_Phone like '%" + tbxCellular_Phone.Text.Trim() + "%') ";
        }
        if (ddlDonor_Type.SelectedIndex != 0)
        {
            strSql += " and Donor_Type ='" + ddlDonor_Type.SelectedItem.Text + "'";
        }
        if (ddlCity.SelectedIndex != 0)
        {
            strSql += " and City ='" + ddlCity.SelectedValue + "'";
        }
        if (ddlArea.SelectedIndex != 0)
        {
            strSql += " and Area = '" + ddlArea.SelectedValue + "'";
        }
        if (tbxAddress.Text.Trim() != "")
        {
            strSql += " and Address like '%" + tbxAddress.Text.Trim() + "%' ";
        }

        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count == 0)
        {
            SetSysMsg("查無相關資料，請先新增捐款人資料！");
            Response.Redirect(Util.RedirectByTime("../DonorMgr/DonorInfo_Add.aspx"));
        }

        else
        {
            //Grid initial
            //20140415Modify by GoodTV Tanya:因查詢結果資料量大時，未分頁會查很久出不來，因此加上分頁
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            //npoGridView.ShowPage = false;
            npoGridView.Keys.Add("編號");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            npoGridView.EditLink = Util.RedirectByTime("Contribute_Add.aspx", "Donor_Id=");
            lblGridList.Text = "** 請選擇您要輸入的捐款人 ** " + npoGridView.Render();
        }
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        if (tbxDonor_Name.Text.Trim() == "" && tbxDonor_Id.Text.Trim() == "" && tbxIDNo.Text.Trim() == "" && tbxCellular_Phone.Text.Trim() == "" && ddlDonor_Type.SelectedIndex == 0 && ddlCity.SelectedIndex == 0 && ddlArea.SelectedIndex == 0 && tbxAddress.Text.Trim() == "")
        {
            ShowSysMsg("請輸入查詢條件！");
            return;
        }
        LoadFormData();
    }
    protected void btnDonor_Add_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("../DonorMgr/DonorInfo_Add.aspx"));
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("ContributeList.aspx"));
    }
}