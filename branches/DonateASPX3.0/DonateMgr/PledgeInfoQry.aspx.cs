using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class DonateMgr_PledgeQry : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
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
        strSql += " Category as [類別], Donor_Type as [身分別], Cellular_Phone as [手機], Tel_Office as [電話日],  Email as [Email], IDNo as [身分證],\n";
        strSql += " Contactor as [聯絡人],  (Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) as [通訊地址], \n";
        strSql += " CONVERT(VARCHAR(10) , Last_DonateDate, 111 ) as [最近捐款日期]\n";
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
        if (tbxMemberNo.Text.Trim() != "")
        {
            ////該查詢哪個欄位?
            //strSql += " and   like '%" + tbxMemberNo.Text.Trim() + "%' ";
        }
        if (tbxIDNo.Text.Trim() != "")
        {
            strSql += " and IDNo like '%" + tbxIDNo.Text.Trim() + "%' ";
        }
        if (tbxCellular_Phone.Text.Trim() != "")
        {
            strSql += " and Cellular_Phone like '%" + tbxCellular_Phone.Text.Trim() + "%' ";
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
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.ShowPage = false;
            npoGridView.Keys.Add("編號");
            npoGridView.EditLink = Util.RedirectByTime("Pledge_Add.aspx", "Donor_Id=");
            lblGridList.Text = "** 請選擇您要輸入的捐款人 **\n" + npoGridView.Render();
        }
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        if (tbxDonor_Name.Text.Trim() == "" && tbxDonor_Id.Text.Trim() == "" && tbxMemberNo.Text.Trim() == "" && tbxIDNo.Text.Trim() == "" && tbxCellular_Phone.Text.Trim() == "" && ddlDonor_Type.SelectedIndex == 0 && ddlCity.SelectedIndex == 0 && ddlArea.SelectedIndex == 0 && tbxAddress.Text.Trim() == "")
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
        Response.Redirect(Util.RedirectByTime("PledgeInfo.aspx"));
    }
}