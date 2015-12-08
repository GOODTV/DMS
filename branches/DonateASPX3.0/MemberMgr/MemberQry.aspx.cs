using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Member_MemberQry : BasePage
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
        Session["ProgID"] = "MemberQry";
        //權控處理
        AuthrityControl();
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            LoadDropDownListData();
            lblGridList.Text = "** 請先輸入查詢條件 **";
            Session["strSql"] = "";
        }
    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_AddNew", btnAdd);
        Authrity.CheckButtonRight("_Query", btnQuery);
        Authrity.CheckButtonRight("_Print", btnPrint);
        Authrity.CheckButtonRight("_Print", btnToxls);
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        Util.FillDropDownList(ddlDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "", ""), "CaseName", "CaseID", false);
        ddlDonor_Type.Items.Insert(0, new ListItem("請選擇", ""));
        ddlDonor_Type.SelectedIndex = 0;

        Util.FillDropDownList(ddlCity, Util.GetDataTable("CodeCity", "ParentCityID", "0", "", ""), "Name", "ZipCode", false);
        ddlCity.Items.Insert(0, new ListItem("縣 市", ""));
        ddlCity.SelectedIndex = 0;

        ddlArea.Items.Insert(0, new ListItem("鄉鎮市區", ""));
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
        strSql = "select Distinct Donor_Id ,CONVERT(nvarchar(20),Donor_Id) as [編號],\n";
        strSql += " (Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End) as [讀者姓名],\n";
        strSql += " (Case When Cellular_Phone<>'' Then Cellular_Phone Else Case When Tel_Office<>'' Then Tel_Office Else Tel_Home End End) as [連絡電話],\n";
        strSql += " (Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) as [通訊地址],\n";
        strSql += " (Case When IsSendEpaper='Y' Then 'V' Else '' End) as [電子文宣]\n";
        strSql += " From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where IsMember='Y' and DeleteDate is null ";

        Dictionary<string, object> dict = new Dictionary<string, object>();

        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and (Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%' or NickName like '%" + tbxDonor_Name.Text.Trim() + "%')";
            strSql += " and (Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%' or NickName like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%')";
        }
        if (tbxMember_No.Text.Trim() != "")
        {
            strSql += " and Donor_Id Like '%" + tbxMember_No.Text.Trim() + "%' ";
        }
        if (tbxTel_Office.Text.Trim() != "")
        {
            strSql += " and (Tel_Office like '%" + tbxTel_Office.Text.Trim() + "%' or Tel_Home like '%" + tbxTel_Office.Text.Trim() + "%' or Cellular_Phone like '%" + tbxTel_Office.Text.Trim() + "%')";
        }
        if (ddlDonor_Type.SelectedIndex != 0)
        {
            strSql += " and Donor_Type ='" + ddlDonor_Type.SelectedItem.Text + "'";
        }
        if (cbxIsSendEpaper.Checked)
        {
            strSql += " and IsSendEpaper ='Y' ";
        }
        if (cbxIsAbroad.Checked)
        {
            strSql += " and IsAbroad='Y' ";
        }
        if (ddlCity.SelectedIndex != 0)
        {
            strSql += " and (City='" + ddlCity.SelectedValue + "' or Invoice_City ='" + ddlCity.SelectedValue + "')";
        }
        if (ddlArea.SelectedIndex != 0)
        {
            strSql += " and (Area='" + ddlArea.SelectedValue + "' or Invoice_Area ='" + ddlArea.SelectedValue + "')";
        }
        if (tbxAddress.Text.Trim() != "")
        {
            strSql += " and (Address like '%" + tbxAddress.Text.Trim() + "%' or Invoice_Address like '%" + tbxAddress.Text.Trim() + "%')";
        }

        DataTable dt;
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
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            npoGridView.Keys.Add("Donor_Id");
            npoGridView.DisableColumn.Add("Donor_Id");
            npoGridView.EditLink = Util.RedirectByTime("Member_Edit.aspx", "Donor_Id=");
            lblGridList.Text = npoGridView.Render();

            Session["strSql"] = strSql;
        }
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        // 2014/8/13 修改若輸入「讀者姓名」查無此讀者資料，按「新增」按鈕後，應可直接帶入「讀者姓名」欄位。
        //Response.Redirect(Util.RedirectByTime("Member_Add.aspx"));
        Response.Redirect(Util.RedirectByTime("Member_Add.aspx", "Member_Name=" + tbxDonor_Name.Text));
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Response.Redirect("MemberQry_Print_Excel.aspx");
    }
}