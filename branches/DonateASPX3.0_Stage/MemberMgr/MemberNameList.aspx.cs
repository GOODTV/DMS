using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class MemberMgr_MemberNameList : BasePage
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
        Session["ProgID"] = "MemberNameList";
        //權控處理
        AuthrityControl();
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            LoadDropDownListData();
            LoadCheckBoxListData();
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
    }
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.SelectedIndex = 1;

        //狀態
        Util.FillDropDownList(ddlMember_Status, Util.GetDataTable("CaseCode", "GroupName", "會員狀態", "", ""), "CaseName", "CaseID", false);
        ddlMember_Status.Items.Insert(0, new ListItem("", ""));
        ddlMember_Status.SelectedIndex = 0;

        //通訊地址-縣市
        Util.FillDropDownList(ddlCity, Util.GetDataTable("CodeCity", "ParentCityID", "0", "", ""), "Name", "ZipCode", false);
        ddlCity.Items.Insert(0, new ListItem("縣　市", "縣　市"));
        ddlCity.SelectedIndex = 0;

        //生日月份
        Util.FillDropDownList(ddlBirthMonth, Util.GetDataTable("CaseCode", "GroupName", "生日月份", "", ""), "CaseName", "CaseName", false);
        ddlBirthMonth.Items.Insert(0, new ListItem("請選擇", "請選擇"));
        ddlBirthMonth.SelectedIndex = 0;
    }
    public void LoadCheckBoxListData()
    {
        //身份別
        Util.FillCheckBoxList(cblDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "", ""), "CaseName", "CaseID", false);
        cblDonor_Type.Items[0].Selected = false;
        //文宣品
        cblPropaganda.Items.Add("紙本月刊");
        cblPropaganda.Items.Add("電子文宣");
    }
    //---------------------------------------------------------------------------
    //dropdownlist選縣市自動帶入鄉鎮市區
    protected void ddlCity_SelectedIndexChanged(object sender, EventArgs e)
    {
        Util.FillDropDownList(ddlArea, Util.GetDataTable("CodeCity", "ParentCityID", ddlCity.SelectedValue, "", ""), "Name", "ZipCode", false);
        ddlArea.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    //查詢
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = @"select
                        Distinct Donor_Id, 
                        Donor_Id as 編號,
	                    (Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End) as 讀者姓名,
                        Member_Status as 狀態, 
	                    Sex as 性別,
	                    Donor_Type as 身分別,
	                    Tel_Office as 聯絡電話日,
	                    Cellular_Phone as 手機,
	                    Tel_Home as 聯絡電話夜,
	                    Email as 電子信箱,
	                    (Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) as 通訊地址
	                From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode
                    where DeleteDate is null  and IsMember ='Y' ";

        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and Dept_Id='" + ddlDept.SelectedValue + "' ";
        }
        
        //身分別-----------------------------------------------------------------------------------//
        bool first = true; int cnt = 0;
        for (int i = 0; i < cblDonor_Type.Items.Count; i++)
        {
            if (cblDonor_Type.Items[i].Selected && first)
            {
                strSql += "and (Donor_Type like '%" + cblDonor_Type.Items[i].ToString() + "%' "; first = false; cnt++;
            }
            else if (cblDonor_Type.Items[i].Selected)
            {
                strSql += "or Donor_Type like '%" + cblDonor_Type.Items[i].ToString() + "%' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //----------------------------------------------------------------------------------------//
        if (ddlMember_Status.SelectedIndex != 0)
        {
            strSql += "and Member_Status= '" + ddlMember_Status.SelectedItem.Text + "'";
        }
        
        if (ddlCity.SelectedIndex != 0)
        {
            strSql += "and City= '" + ddlCity.SelectedValue + "'";
        }
        if (ddlArea.SelectedIndex != 0 && ddlArea.Items.Count != 0)
        {
            strSql += "and Area like '%" + ddlArea.SelectedValue + "'";
        }
        //----------------------------------------------------------------------------------------//
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += "and Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%' ";
            strSql += "and Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%' ";
        }
        if (tbxMember_NoS.Text.Trim() != "")
        {
            strSql += @"and Donor_Id >='" + tbxMember_NoS.Text.Trim() + "'";
        }
        if (tbxMember_NoE.Text.Trim() != "")
        {
            strSql += @"and Donor_Id <='" + tbxMember_NoE.Text.Trim() + "'";
        }
        if (ddlBirthMonth.SelectedIndex != 0)
        {
            strSql += "and cast(MONTH([Birthday])as nvarchar)+'月'='" + ddlBirthMonth.SelectedItem.Text + "'";
        }
        //文宣品-----------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblPropaganda.Items.Count; i++)
        {
            if (cblPropaganda.Items[i].Selected && first)
            {
                strSql += "and ("; first = false;
                switch (i)
                {
                    case 0: strSql += " (IsSendNewsNum<>'' or IsSendNewsNum<>'0')"; break;
                    case 1: strSql += " IsSendEpaper='Y' "; break;
                }
                cnt++;
            }
            else if (cblPropaganda.Items[i].Selected)
            {
                strSql += "and";
                switch (i)
                {
                    case 0: strSql += " (IsSendNewsNum<>'' or IsSendNewsNum<>'0') "; break;
                    case 1: strSql += " IsSendEpaper='Y' "; break;
                }
                cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //----------------------------------------------------------------------------------------//
        if (cbxIsAbroad.Checked)
        {
            strSql += "and IsAbroad='Y'";
        }

        Session["strSql"] = strSql;
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dt = NpoDB.GetDataTableS(strSql, dict);

        //Grid initial
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.DisableColumn.Add("Donor_Id");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
        lblGridList.Text = npoGridView.Render();
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Response.Redirect("MemberNameList_Print_Excel.aspx");
    }
}