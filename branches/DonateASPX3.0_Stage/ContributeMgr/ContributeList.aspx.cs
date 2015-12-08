using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
/*
 * 群組名稱設定
 * 資料庫: Contribute
 * List頁面: ContributeList
 * 新增頁面: Contribute_Add.aspx 
 * 詳細頁面: Contribute_Detail.aspx
 * 修改頁面: Contribute_Edit.aspx
 */
public partial class ContributeMgr_ContributeList : BasePage
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
        Session["ProgID"] = "ContributeList";
        //權控處理
        AuthrityControl();
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
        if (!IsPostBack)
        {
            Session["cType"] = "";
            LoadDropDownListData();
            txtContribute_DateS.Text = DateTime.Now.Year.ToString() + "/1/1";
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
        Util.FillDropDownList(ddlContribute_Payment, Util.GetDataTable("CaseCode", "GroupName", "物品捐贈方式", "", ""), "CaseName", "CaseName", false);
        ddlContribute_Payment.Items.Insert(0, new ListItem("", ""));
        ddlContribute_Payment.SelectedIndex = 0;

        //捐贈用途
        Util.FillDropDownList(ddlContribute_Purpose, Util.GetDataTable("CaseCode", "GroupName", "物品捐贈用途", "", ""), "CaseName", "CaseName", false);
        ddlContribute_Purpose.Items.Insert(0, new ListItem("", ""));
        ddlContribute_Purpose.SelectedIndex = 0;

        //經手人
        Util.FillDropDownList(ddlCreate_User, Util.GetDataTable("AdminUser", "1", "1", "", ""), "UserName", "UserName", false);
        ddlCreate_User.Items.Insert(0, new ListItem("", ""));
        ddlCreate_User.SelectedIndex = 0;

        //收據開立
        ddlInvoice_Type.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Type.Items.Insert(1, new ListItem("不寄", "不寄"));
        ddlInvoice_Type.Items.Insert(2, new ListItem("單次收據", "單次收據"));
        ddlInvoice_Type.SelectedIndex = 0;

        //會計科目
        Util.FillDropDownList(ddlAccounting_Title, Util.GetDataTable("CaseCode", "GroupName", "款項會計科目", "", ""), "CaseName", "CaseName", false);
        ddlAccounting_Title.Items.Insert(0, new ListItem("", ""));
        ddlAccounting_Title.SelectedIndex = 0;

        //募款活動
        Util.FillDropDownList(ddlAct_Id, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        ddlAct_Id.Items.Insert(0, new ListItem("", ""));
        ddlAct_Id.SelectedIndex = 0;

    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = "select D.Donor_Id as [編號] ,C.Contribute_Id as [捐物編號], (Case When D.NickName<>'' Then D.Donor_Name+'('+D.NickName+')' Else D.Donor_Name End)as [捐贈人], \n";
        strSql += " CONVERT(nvarchar,C.Contribute_Date,111) as [捐贈日期], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,C.Contribute_Amt),1),'.00','') as [折合現金],C.Contribute_Payment as [捐款方式],\n";
        strSql += " C.Contribute_Purpose as [捐款用途], C.Invoice_Type as [收據開立],C.Invoice_Pre + C.Invoice_No as [收據編號],\n";
        strSql += " (Case When C.Invoice_Print='1' Then 'V' Else '' End) as [列印],Case When Export='Y' Then '作廢'  Else Case When Issue_Type='M' Then '手開' Else '' End End as [狀態],\n";
        strSql += " C.Create_User as [經手人]\n";
        strSql += " WhereClause";
        strSql += " from Contribute C left join Donor D on C.Donor_Id = D.Donor_Id ";
        strSql += " where 1=1";
        

        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and C.Dept_Id = '" + ddlDept.SelectedValue + "'";
        }
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and D.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%'";
            strSql += " and D.Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%'";
        }
        if (tbxDonor_Id.Text.Trim() != "")
        {
            strSql += " and D.Donor_Id like '%" + tbxDonor_Id.Text.Trim() + "%'";
        }
        if (ddlContribute_Payment.SelectedIndex != 0)
        {
            strSql += " and C.Contribute_Payment = '" + ddlContribute_Payment.SelectedItem.Text + "'";
        }
        if (ddlContribute_Purpose.SelectedIndex != 0)
        {
            strSql += " and C.Contribute_Purpose = '" + ddlContribute_Purpose.SelectedItem.Text + "'";
        }
        if (ddlCreate_User.SelectedIndex != 0)
        {
            strSql += " and C.Create_User = '" + ddlCreate_User.SelectedItem.Text + "'";
        }
        if (ddlInvoice_Type.SelectedIndex != 0)
        {
            strSql += " and C.Invoice_Type = '" + ddlInvoice_Type.SelectedItem.Text + "'";
        }
        if (txtContribute_DateS.Text.Trim() != "")
        {
            strSql += " and C.Contribute_Date >= '" + Util.DateTime2String(txtContribute_DateS.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty)  + "' ";
        }
        if (txtContribute_DateE.Text.Trim() != "")
        {
            strSql += " and C.Contribute_Date <= '" + Util.DateTime2String(txtContribute_DateE.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "' ";
        }
        //Invoice_No
        if (tbxInvoice_NoS.Text.Trim() != "")
        {
            strSql += " and C.Invoice_No >= '" + tbxInvoice_NoS.Text.Trim().PadLeft(12,'0') + "' ";
        }
        if (tbxInvoice_NoE.Text.Trim() != "")
        {
            strSql += " and C.Invoice_No <= '" + tbxInvoice_NoE.Text.Trim().PadLeft(12, '0') + "' ";
        }
        //
        if (tbxAccoun_DateS.Text.Trim() != "")
        {
            strSql += " and C.Accoun_Date >= '" + Util.DateTime2String(tbxAccoun_DateS.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "' ";
        }
        if (tbxAccoun_DateE.Text.Trim() != "")
        {
            strSql += " and C.Accoun_Date <= '" + Util.DateTime2String(tbxAccoun_DateE.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "' ";
        }

        if (ddlAccounting_Title.SelectedIndex != 0)
        {
            strSql += " and C.Accounting_Title = '" + ddlAccounting_Title.SelectedItem.Text + "'";
        }
        if (ddlAct_Id.SelectedIndex != 0)
        {
            strSql += " and C.Act_Id = '" + ddlAct_Id.SelectedValue + "'";
        }
        if (cbxInvoice_Pre.Checked == true)
        {
            strSql += " and C.Issue_Type = 'M'";
        }
        if (cbxExport.Checked == true)
        {
            strSql += " and C.Export = 'Y'";
        }
        Session["strSql"] = strSql;
        strSql = strSql.Replace("WhereClause", "");
        dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
            lblAmt.Text = "0";
            Session["Amt"] = lblAmt.Text;
            //Session["strSql"] = "";
        }
        else
        {
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.Keys.Add("捐物編號");
            npoGridView.DisableColumn.Add("捐物編號");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            Session["cType"] = "ContributeList";
            npoGridView.EditLink = Util.RedirectByTime("Contribute_Detail.aspx", "Contribute_Id=");
            lblGridList.Text = npoGridView.Render();

            //折合現金合計
            lblAmt.Text = Contribute_Amt();
            Session["Amt"] = lblAmt.Text;
        }
        
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Response.Redirect("ContributeList_Print_Excel.aspx");
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Session["cType"] = "ContributeList";
        Response.Redirect(Util.RedirectByTime("DonorQry.aspx"));
    }
    private string Contribute_Amt()
    {
        string strSql = "select REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when Export = 'N' then Contribute_Amt else 0 end)),1),'.00','') as Contribute_Amt \n";
        strSql += " from Contribute C\n";
        strSql += " left join Donor D on C.Donor_Id = D.Donor_Id\n";
        strSql += " where 1=1 ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and C.Dept_Id = '" + ddlDept.SelectedValue + "'";
        }
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and D.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%'";
            strSql += " and D.Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%'";
        }
        if (tbxDonor_Id.Text.Trim() != "")
        {
            strSql += " and D.Donor_Id like '%" + tbxDonor_Id.Text.Trim() + "%'";
        }
        if (ddlContribute_Payment.SelectedIndex != 0)
        {
            strSql += " and C.Contribute_Payment = '" + ddlContribute_Payment.SelectedItem.Text + "'";
        }
        if (ddlContribute_Purpose.SelectedIndex != 0)
        {
            strSql += " and C.Contribute_Purpose = '" + ddlContribute_Purpose.SelectedItem.Text + "'";
        }
        if (ddlCreate_User.SelectedIndex != 0)
        {
            strSql += " and C.Create_User = '" + ddlCreate_User.SelectedItem.Text + "'";
        }
        if (ddlInvoice_Type.SelectedIndex != 0)
        {
            strSql += " and C.Invoice_Type = '" + ddlInvoice_Type.SelectedItem.Text + "'";
        }
        if (txtContribute_DateS.Text.Trim() != "")
        {
            strSql += " and C.Contribute_Date >= '" + Util.DateTime2String(txtContribute_DateS.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "' ";
        }
        if (txtContribute_DateE.Text.Trim() != "")
        {
            strSql += " and C.Contribute_Date <= '" + Util.DateTime2String(txtContribute_DateE.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "' ";
        }
        //Invoice_No
        if (tbxInvoice_NoS.Text.Trim() != "")
        {
            strSql += " and C.Invoice_No >= '" + tbxInvoice_NoS.Text.Trim().PadLeft(12, '0') + "' ";
        }
        if (tbxInvoice_NoE.Text.Trim() != "")
        {
            strSql += " and C.Invoice_No <= '" + tbxInvoice_NoE.Text.Trim().PadLeft(12, '0') + "' ";
        }
        //
        if (tbxAccoun_DateS.Text.Trim() != "")
        {
            strSql += " and C.Accoun_Date >= '" + Util.DateTime2String(tbxAccoun_DateS.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "' ";
        }
        if (tbxAccoun_DateE.Text.Trim() != "")
        {
            strSql += " and C.Accoun_Date <= '" + Util.DateTime2String(tbxAccoun_DateE.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "' ";
        }
        if (ddlAccounting_Title.SelectedIndex != 0)
        {
            strSql += " and C.Accounting_Title = '" + ddlAccounting_Title.SelectedItem.Text + "'";
        }
        if (ddlAct_Id.SelectedIndex != 0)
        {
            strSql += " and C.Act_Id = '" + ddlAct_Id.SelectedValue + "'";
        }
        if (cbxInvoice_Pre.Checked == true)
        {
            strSql += " and C.Issue_Type = 'M'";
        }
        if (cbxExport.Checked == true)
        {
            strSql += " and C.Export = 'Y'";
        }

        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Contribute_Amt = dr["Contribute_Amt"].ToString();
        return Contribute_Amt;
    }
}