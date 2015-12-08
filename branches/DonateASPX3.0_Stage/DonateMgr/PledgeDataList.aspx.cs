using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;

public partial class DonateMgr_PledgeDataList : BasePage
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
            MemberMenu.Text = CaseUtil.MakeMenu(HFD_Uid.Value, 3); //傳入目前選擇的 tab
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
        strSql = @" select *  , (Case When ISNULL(DONOR.Invoice_City,'')='' Then Invoice_Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.Invoice_ZipCode+B.mValue+Invoice_Address Else A.mValue+DONOR.Invoice_ZipCode+Invoice_Address End End) as [地址]
                    from DONOR 
                        Left Join CODECITY As A On Donor.Invoice_City=A.mCode Left Join CODECITY As B On Donor.Invoice_Area=B.mCode 
                    where Donor_id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("DonorInfo.aspx");

        DataRow dr = dt.Rows[0];

        //捐款人
        txtDonor_Name.Text = dr["Donor_Name"].ToString() + " " + dr["Title"].ToString(); ;
        //捐款人編號
        tbxDonor_Id.Text = dr["Donor_Id"].ToString();
        //收據開立
        tbxInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString();
        //收據地址
        tbxAddress.Text = dr["地址"].ToString();
    }
    public void LoadFormData()
    {
        DataTable dt;
        string strSql = "";
        strSql = "select P.Pledge_Id as [授權編號], P.Donate_Payment as [授權方式], P.Donate_Purpose as [指定用途],\n";
        strSql += " REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,P.Donate_Amt),1),'.00','') as [扣款金額],\n";
        strSql += " CONVERT(NVARCHAR, P.Donate_FromDate, 111) as [授權起日], CONVERT(NVARCHAR, P.Donate_ToDate, 111) as [授權迄日],\n";
        strSql += " P.Donate_Period as [轉帳週期], CONVERT(NVARCHAR, P.Next_DonateDate, 111) as [下次扣款日期],\n";
        strSql += " substring(P.Valid_Date,1,2) +'/'+ substring(P.Valid_Date,3,2) as [信用卡有效月年],\n";
        strSql += " P.Status as [授權狀態], P.Create_User as [經手人]\n";
        strSql += " from Pledge P where Donor_Id ='" + HFD_Uid.Value + "'";
        strSql += " order by Pledge_Id desc";

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
            npoGridView.Keys.Add("授權編號");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            Session["cType"] = "PledgeDataList";
            npoGridView.EditLink = Util.RedirectByTime("Pledge_Edit.aspx", "Pledge_Id=");
            lblGridList.Text = npoGridView.Render();
        }
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Session["cType"] = "PledgeDataList";
        Response.Redirect(Util.RedirectByTime("Pledge_Add.aspx", "Donor_Id=" + HFD_Uid.Value));
    }
}