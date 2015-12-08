using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;

public partial class ContributeMgr_ContributeDataList : BasePage
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

        //權控處理
        AuthrityControl();
        if (!IsPostBack)
        {
            Session["cType"] = "";
            HFD_Uid.Value = Util.GetQueryString("Donor_Id");
            MemberMenu.Text = CaseUtil.MakeMenu(HFD_Uid.Value, 4); //傳入目前選擇的 tab
            Form_DataBind();
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
        Authrity.CheckButtonRight("_AddNew", btnAdd);
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
        strSql = " select *  from DONOR where Donor_id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("DonorInfo.aspx");

        DataRow dr = dt.Rows[0];

        //捐款人
        txtDonor_Name.Text = dr["Donor_Name"].ToString() + " "  + dr["Title"].ToString();
        //捐款人編號
        tbxDonor_Id.Text = dr["Donor_Id"].ToString();
    }
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = @"select C.Gift_Id , CONVERT(nvarchar,C.Gift_Date,111) as [贈送日期], C.Create_User as [經手人]
                    ,(SELECT Linked2_Name + ',' 
                    FROM GiftData CD 
                    right join Linked2 L on CD.Goods_Id = L.Ser_No where C.Gift_Id=CD.Gift_Id For XML PATH('')) as [物品名稱]
                    from Gift C
                    where Donor_Id ='" + HFD_Uid.Value + "'";

        dt = NpoDB.QueryGetTable(strSql);

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
            npoGridView.Keys.Add("Gift_Id");
            npoGridView.DisableColumn.Add("Gift_Id");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            Session["cType"] = "ContributeDataList";
            npoGridView.EditLink = Util.RedirectByTime("../DonorMgr/Gift_Detail.aspx", "Gift_Id=");
            lblGridList.Text = npoGridView.Render();
        }
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("../DonorMgr/Gift_Add.aspx", "Donor_Id=" + HFD_Uid.Value));
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("../DonorMgr/Gift_Print_Excel.aspx", "Donor_Id=" + HFD_Uid.Value));
    }
}