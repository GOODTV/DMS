using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
/*
 * 群組名稱設定
 * 資料庫: Donate
 * List頁面: DonateInfo
 * 新增頁面: DonateInfo_Add.aspx 
 * 修改頁面: DonateInfo_Edit.aspx
 */
public partial class DonateMgr_DonateInfo : BasePage
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
        Session["ProgID"] = "DonateInfo";
        //權控處理
        AuthrityControl();

        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            Session["cType"] = "";
            LoadDropDownListData();
            //2014/4/9 修改成進入時先不查詢
            //LoadFormData();
            lblGridList.Text = "** 請先輸入查詢條件 **";
            lblGridList.ForeColor = System.Drawing.Color.Red; //"Red";
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
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.Items.Insert(0, new ListItem("", ""));
        ddlDept.SelectedIndex = 1;

        //捐款方式
        Util.FillDropDownList(ddlDonate_Payment, Util.GetDataTable("CaseCode", "GroupName", "捐款方式", "ABS(CaseID)", ""), "CaseName", "CaseName", false);
        ddlDonate_Payment.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Payment.SelectedIndex = 0;

        //捐款用途
        Util.FillDropDownList(ddlDonate_Purpose, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseName", false);
        ddlDonate_Purpose.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Purpose.SelectedIndex = 0;

        //經手人
        Util.FillDropDownList(ddlCreate_User, Util.GetDataTable("AdminUser", "1", "1", "", ""), "UserName", "UserName", false);
        ddlCreate_User.Items.Insert(0, new ListItem("", ""));
        ddlCreate_User.SelectedIndex = 0;

        //收據開立
        Util.FillDropDownList(ddlInvoice_Type, Util.GetDataTable("CaseCode", "GroupName", "收據開立", "", ""), "CaseName", "CaseName", false);
        ddlInvoice_Type.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Type.SelectedIndex = 0;

        //入帳銀行
        Util.FillDropDownList(ddlAccoun_Bank, Util.GetDataTable("CaseCode", "GroupName", "入帳銀行", "", ""), "CaseName", "CaseName", false);
        ddlAccoun_Bank.Items.Insert(0, new ListItem("", ""));
        ddlAccoun_Bank.SelectedIndex = 0;

        //會計科目
        Util.FillDropDownList(ddlAccounting_Title, Util.GetDataTable("CaseCode", "GroupName", "款項會計科目", "", ""), "CaseName", "CaseName", false);
        ddlAccounting_Title.Items.Insert(0, new ListItem("", ""));
        ddlAccounting_Title.SelectedIndex = 0;

        //募款活動
        Util.FillDropDownList(ddlAct_Id, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        ddlAct_Id.Items.Insert(0, new ListItem("", ""));
        ddlAct_Id.SelectedIndex = 0;

    }
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        //20140416 Modify by GoodTV Tanya:增加判斷是否為「手開收據」
        //20140425 Modify by GoodTV Tanya:修改「作廢收據」判斷為Issue_Type = 'D'
        strSql = "select do.Donate_Id ,Case When dr.Member_No<>'' Then dr.Member_No Else do.Donor_Id End as 編號, dr.Donor_Name as 捐款人, do.Invoice_Title as '收據抬頭' ,CONVERT(VARCHAR(10) , do.Donate_Date, 111 ) as 捐款日期,\n";
        strSql += " REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,case when ISNULL(do.Issue_Type,'') !='D' then do.Donate_Amt else '0' end),1),'.00','') as 捐款金額, do.Donate_Payment as 捐款方式, do.Donate_Purpose as 捐款用途, a.Act_Name as 專案活動,\n";
        strSql += " do.Invoice_Type as 收據開立, ISNULL(do.Invoice_Pre,'') + do.Invoice_No as 收據編號, do.Accoun_Date as 沖帳日期, Case when (do.Invoice_Print = '1' or do.Invoice_Print_Yearly = '1') then 'V' else '' end as 列印,\n";
        strSql += " case when do.Issue_Type = 'D' then '作廢' when do.Issue_Type = 'M' then '手開' else '' end as 狀態, do.Create_User as 經手人, dr.Donor_Id_Old as 舊編號\n";
        strSql += " from donate do\n";
        strSql += " left join donor dr on do.Donor_Id = dr. Donor_Id\n";
        strSql += " left join Act A on do.Act_Id = A.Act_Id\n";
        strSql += " where 1=1 and dr.DeleteDate is null and do.Donor_Id <> '' ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and do.Dept_Id = '" + ddlDept.SelectedValue + "'";
        }
        
        if (tbxDonor_Name.Text.Trim()!="")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%'"; 
            strSql += " and dr.Donor_Name like N'%" + tbxDonor_Name.Text.Trim().Replace("'","''") + "%'"; 
        }
        if (tbxDonor_Id.Text.Trim() != "")
        {
            //20140425 Modify by GoodTV Tanya
            strSql += " and do.Donor_Id= '" + tbxDonor_Id.Text.Trim() + "'";
            //dict.Add("Donor_Id", tbxDonor_Id.Text.Trim());
        }
        if (tbxDonor_Id_Old.Text.Trim() != "")
        {
            //20140425 Modify by GoodTV Tanya
            strSql += " and do.Donor_Id_Old= '" + tbxDonor_Id_Old.Text.Trim() + "'";
            //dict.Add("Donor_Id_Old", tbxDonor_Id_Old.Text.Trim());
        }
        if (ddlDonate_Payment.SelectedIndex != 0)
        {
            strSql += " and do.Donate_Payment = '" + ddlDonate_Payment.SelectedItem.Text + "'";
        }
        if (ddlDonate_Purpose.SelectedIndex != 0)
        {
            strSql += " and do.Donate_Purpose = '" + ddlDonate_Purpose.SelectedItem.Text + "'";
        }
        if (txtDonateDateS.Text.Trim() != "")
        {
            strSql += " and do.Donate_Date >= '" + txtDonateDateS.Text.Trim() + "' ";
        }
        if (txtDonateDateE.Text.Trim() != "")
        {
            strSql += " and do.Donate_Date <= '" + txtDonateDateE.Text.Trim() + "' ";
        }
        if (ddlInvoice_Type.SelectedIndex != 0)
        {
            strSql += " and do.Invoice_Type = '" + ddlInvoice_Type.SelectedItem.Text + "'";
        }
        if (tbxInvoice_NoS.Text.Trim() != "" && tbxInvoice_NoE.Text.Trim() == "")
        {
            strSql += " and do.Invoice_No between '" + tbxInvoice_NoS.Text.Trim() + "' and CONVERT(VARCHAR(8) , GETDATE(), 112 ) ";
        }
        if (tbxInvoice_NoS.Text.Trim() == "" && tbxInvoice_NoE.Text.Trim() != "")
        {
            strSql += " and do.Invoice_No between '19000101' and'" + tbxInvoice_NoE.Text.Trim() + "' ";
        }
        if (tbxInvoice_NoS.Text.Trim() != "" && tbxInvoice_NoE.Text.Trim() != "")
        {
            strSql += " and do.Invoice_No between '" + tbxInvoice_NoS.Text.Trim() + "' and '" + tbxInvoice_NoE.Text.Trim() + "' ";
        }
        if (ddlCreate_User.SelectedIndex != 0)
        {
            strSql += " and do.Create_User = '" + ddlCreate_User.SelectedItem.Text + "'";
        }
        //20140425 Modify by GoodTV Tanya：「手開收據」判斷為Issue_Type = 'M'
        if (cbxInvoice_Pre.Checked == true)
        {
            strSql += " and do.Issue_Type = 'M'";
        }
        //20140425 Modify by GoodTV Tanya：「作廢收據」判斷為Issue_Type = 'D'
        if (cbxExport.Checked == true)
        {
            strSql += " and do.Issue_Type = 'D'";
        }

        if (tbxDonation_NumberNo.Text.Trim() != "")
        {
            //20140425 Modify by GoodTV Tanya
            strSql += " and do.Donation_NumberNo= '" + tbxDonation_NumberNo.Text.Trim() + "'";
            //dict.Add("Donation_NumberNo", tbxDonation_NumberNo.Text.Trim());
        }
        if (tbxDonation_SubPoenaNo.Text.Trim() != "")
        {
            //20140425 Modify by GoodTV Tanya
            strSql += " and do.Donation_SubPoenaNo= '" + tbxDonation_SubPoenaNo.Text.Trim() + "'";
            //dict.Add("Donation_SubPoenaNo", tbxDonation_SubPoenaNo.Text.Trim());
        }
        if (ddlAct_Id.SelectedIndex != 0)
        {
            strSql += " and do.Act_Id = '" + ddlAct_Id.SelectedValue + "'";
        }
        if (tbxAccoun_DateS.Text.Trim() != "" && tbxAccoun_DateE.Text.Trim() == "")
        {
            strSql += " and do.Accoun_Date >= '" + tbxAccoun_DateS.Text.Trim() + "' ";
        }
        if (tbxAccoun_DateS.Text.Trim() == "" && tbxAccoun_DateE.Text.Trim() != "")
        {
            strSql += " and do.Accoun_Date <= '" + tbxAccoun_DateE.Text.Trim() + "' ";
        }
        if (ddlAccoun_Bank.SelectedIndex != 0)
        {
            strSql += " and do.Accoun_Bank = '" + ddlAccoun_Bank.SelectedItem.Text + "'";
        }
        if (ddlAccounting_Title.SelectedIndex != 0)
        {
            strSql += " and do.Accounting_Title = '" + ddlAccounting_Title.SelectedItem.Text + "'";
        }
        //20150331新增收據抬頭模糊查詢
        if (tbxInvoice_Title.Text.Trim() != "")
        {
            strSql += " and do.Invoice_Title like N'%" + tbxInvoice_Title.Text.Trim() + "%'";
        }
        strSql += " order by do.Donate_Date desc ";
        dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
            // 2014/4/9 有顏色區別
            lblGridList.ForeColor = System.Drawing.Color.Red;
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
            npoGridView.DisableColumn.Add("沖帳日期");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            Session["cType"] = "DonateInfo";
            npoGridView.EditLink = Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=");
            lblGridList.Text = npoGridView.Render();
            lblGridList.ForeColor = System.Drawing.Color.Black;

            //捐款金額合計
            lblAmt.Text = Donate_Amt();
            Session["Amt"] = lblAmt.Text;
        }

        Session["strSql"] = strSql;
    }
    //---------------------------------------------------------------------------
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Response.Redirect("DonateInfo_Print_Excel.aspx");
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Session["cType"] = "DonateInfo";
        Response.Redirect(Util.RedirectByTime("DonorQry.aspx"));
    }
    private string Donate_Amt()
    {
        string strSql = "select REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when ISNULL(Issue_Type,'') != 'D' then Donate_Amt else 0 end)),1),'.00','') as Donate_Amt \n";
        strSql += " from donate do\n";
        strSql += " left join donor dr on do.Donor_Id = dr. Donor_Id\n";
        strSql += " WhereClause1";
        strSql += " where dr.DeleteDate is null and 1=1 ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and do.Dept_Id = '" + ddlDept.SelectedValue + "'";
        }
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%'"; 
            strSql += " and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%'"; 
        }
        if (tbxDonor_Id.Text.Trim() != "")
        {
            strSql += " and do.Donor_Id=@Donor_Id ";
            dict.Add("Donor_Id", tbxDonor_Id.Text.Trim());
        }
        if (tbxDonor_Id_Old.Text.Trim() != "")
        {
            strSql += " and do.Donor_Id_Old=@Donor_Id_Old ";
            dict.Add("Donor_Id_Old", tbxDonor_Id_Old.Text.Trim());
        }
        if (ddlDonate_Payment.SelectedIndex != 0)
        {
            strSql += " and do.Donate_Payment = '" + ddlDonate_Payment.SelectedItem.Text + "'";
        }
        if (ddlDonate_Purpose.SelectedIndex != 0)
        {
            strSql += " and do.Donate_Purpose = '" + ddlDonate_Purpose.SelectedItem.Text + "'";
        }
        if (txtDonateDateS.Text.Trim() != "")
        {
            strSql += " and do.Donate_Date >= '" + txtDonateDateS.Text.Trim() + "' ";
        }
        if (txtDonateDateE.Text.Trim() != "")
        {
            strSql += " and do.Donate_Date <= '" + txtDonateDateE.Text.Trim() + "' ";
        }
        if (ddlInvoice_Type.SelectedIndex != 0)
        {
            strSql += " and do.Invoice_Type = '" + ddlInvoice_Type.SelectedItem.Text + "'";
        }
        if (tbxInvoice_NoS.Text.Trim() != "" && tbxInvoice_NoE.Text.Trim() == "")
        {
            strSql += " and do.Invoice_No between '" + tbxInvoice_NoS.Text.Trim() + "' and CONVERT(VARCHAR(8) , GETDATE(), 112 ) ";
        }
        if (tbxInvoice_NoS.Text.Trim() == "" && tbxInvoice_NoE.Text.Trim() != "")
        {
            strSql += " and do.Invoice_No between '19000101' and'" + tbxInvoice_NoE.Text.Trim() + "' ";
        }
        if (tbxInvoice_NoS.Text.Trim() != "" && tbxInvoice_NoE.Text.Trim() != "")
        {
            strSql += " and do.Invoice_No between '" + tbxInvoice_NoS.Text.Trim() + "' and '" + tbxInvoice_NoE.Text.Trim() + "' ";
        }
        if (ddlCreate_User.SelectedIndex != 0)
        {
            strSql += " and do.Create_User = '" + ddlCreate_User.SelectedItem.Text + "'";
        }
        //「手開收據」判斷為Issue_Type = 'M'
        if (cbxInvoice_Pre.Checked == true)
        {
            //strSql += " and do.Invoice_Pre = 'B'";
            strSql += " and do.Issue_Type = 'M'";
        }
        //「作廢收據」判斷為Issue_Type = 'D'
        if (cbxExport.Checked == true)
        {
            //strSql += " and do.Export = 'Y'";
            strSql += " and do.Issue_Type = 'D'";
        }

        if (tbxDonation_NumberNo.Text.Trim() != "")
        {
            strSql += " and do.Donation_NumberNo=@Donation_NumberNo ";
            dict.Add("Donation_NumberNo", tbxDonation_NumberNo.Text.Trim());
        }
        if (tbxDonation_SubPoenaNo.Text.Trim() != "")
        {
            strSql += " and do.Donation_SubPoenaNo=@Donation_SubPoenaNo ";
            dict.Add("Donation_SubPoenaNo", tbxDonation_SubPoenaNo.Text.Trim());
        }
        //20150331新增收據抬頭模糊查詢
        if (tbxInvoice_Title.Text.Trim() != "")
        {
            strSql += " and do.Invoice_Title like '%" + tbxInvoice_Title.Text.Trim() + "%'";
        }
        if (ddlAct_Id.SelectedIndex != 0)
        {
            strSql += " and do.Act_Id = '" + ddlAct_Id.SelectedValue + "'";
        }
        if (tbxAccoun_DateS.Text.Trim() != "" && tbxAccoun_DateE.Text.Trim() == "")
        {
            strSql += " and do.Accoun_Date >= '" + tbxAccoun_DateS.Text.Trim() + "' ";
        }
        if (tbxAccoun_DateS.Text.Trim() == "" && tbxAccoun_DateE.Text.Trim() != "")
        {
            strSql += " and do.Accoun_Date <= '" + tbxAccoun_DateE.Text.Trim() + "' ";
        }
        if (ddlAccoun_Bank.SelectedIndex != 0)
        {
            strSql += " and do.Accoun_Bank = '" + ddlAccoun_Bank.SelectedItem.Text + "'";
        }
        if (ddlAccounting_Title.SelectedIndex != 0)
        {
            strSql += " and do.Accounting_Title = '" + ddlAccounting_Title.SelectedItem.Text + "'";
        }
        else
        {
            strSql = strSql.Replace("WhereClause1", "");
        }
        
        DataTable dt;
        //20140425 Modify by GoodTV Tanya
        //dt = NpoDB.QueryGetScalar(strSql, dict);
        dt = NpoDB.GetDataTableS(strSql, dict);
        DataRow dr = dt.Rows[0];
        string Donate_Amt = dr["Donate_Amt"].ToString();
        return Donate_Amt;
    }
}