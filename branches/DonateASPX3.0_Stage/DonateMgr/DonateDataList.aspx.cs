using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;

public partial class DonateMgr_DonateDataList : BasePage
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
            MemberMenu.Text = CaseUtil.MakeMenu(HFD_Uid.Value, 2); //傳入目前選擇的 tab
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
        strSql = @" select * , (Case When ISNULL(DONOR.Invoice_City,'')='' Then Invoice_Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.Invoice_ZipCode+B.mValue+Invoice_Address Else A.mValue+DONOR.Invoice_ZipCode+Invoice_Address End End) as [收據地址]
                    from Donor 
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
        tbxInvoice_Type.Text = dr["Invoice_Type"].ToString();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString();
        //收據地址
        tbxAddress.Text = dr["收據地址"].ToString();

    }
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = "select do.Donate_Id , CONVERT(VARCHAR(10) , do.Donate_Date, 111 ) as 捐款日期, REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,case when ISNULL(do.Issue_Type,'') !='D' then do.Donate_Amt else '0' end),1),'.00','') as 捐款金額,\n";
        strSql += " do.Donate_Payment as 捐款方式, do.Donate_Purpose as 捐款用途, do.Invoice_Type as 收據開立, ISNULL(do.Invoice_Pre,'') + do.Invoice_No as 收據編號,do.Invoice_Title as 收據抬頭,\n";
        strSql += " case when (do.Invoice_Print = '1' or do.Invoice_Print_Yearly = '1') then 'V' else '' end as 列印, \n";
        strSql += " case when do.Issue_Type = 'D' then '作廢' when do.Issue_Type = 'M' then '手開' else '' end as 狀態, do.Create_User as 經手人\n";
        strSql += " from donate do\n";
        strSql += " where Donor_Id ='" + HFD_Uid.Value + "'";
        //20140416 GoodTV Tanya:新增排序同ASP2.0
        strSql += " order by convert(varchar,do.Donate_Date,111) desc,do.Donate_Id desc";

        dt = NpoDB.QueryGetTable(strSql);

        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
            lblAmt.Text = "0";
        }
        else
        {
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.Keys.Add("Donate_Id");
            npoGridView.DisableColumn.Add("Donate_Id");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            Session["cType"] = "DonateDataList";
            npoGridView.EditLink = Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=");
            lblGridList.Text = npoGridView.Render();

            //捐款金額合計
            lblAmt.Text = Donate_Amt();
        }
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Session["cType"] = "DonateDataList";
        Response.Redirect(Util.RedirectByTime("Donate_Add.aspx", "Donor_Id=" + HFD_Uid.Value));
    }
    private string Donate_Amt()
    {
        string strSql = "select REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when ISNULL(Issue_Type,'') != 'D' then Donate_Amt else 0 end)),1),'.00','') as Donate_Amt \n";
        strSql += " from donate do\n";
        strSql += " where Donor_Id ='" + HFD_Uid.Value + "'";
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Donate_Amt = dr["Donate_Amt"].ToString();
        return Donate_Amt;
    }
}