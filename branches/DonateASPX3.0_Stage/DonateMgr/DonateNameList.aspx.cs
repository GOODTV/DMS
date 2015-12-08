using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonateMgr_DonateNameList : BasePage
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
        Session["ProgID"] = "DonorQry";
        //權控處理
        AuthrityControl();
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            LoadDropDownListData();
            LoadCheckBoxListData();
            // 2014/4/10 修正進入預設不載入資料
            //LoadFormData();
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
        ddlDept.Items.Insert(0, new ListItem("", ""));
        ddlDept.SelectedIndex = 1;

        //通訊地址-縣市
        Util.FillDropDownList(ddlCity, Util.GetDataTable("CodeCity", "ParentCityID", "0", "", ""), "Name", "ZipCode", false);
        ddlCity.Items.Insert(0, new ListItem("縣　市", "縣　市"));
        ddlCity.SelectedIndex = 0;

        //收據地址-縣市
        Util.FillDropDownList(ddlCity2, Util.GetDataTable("CodeCity", "ParentCityID", "0", "", ""), "Name", "ZipCode", false);
        ddlCity2.Items.Insert(0, new ListItem("縣　市", "縣　市"));
        ddlCity2.SelectedIndex = 0;

        //募款活動
        //Util.FillDropDownList(ddlActName, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        //ddlActName.Items.Insert(0, new ListItem("", ""));
        //ddlActName.SelectedIndex = 0;

    }
    public void LoadCheckBoxListData()
    {
        //身份別
        Util.FillCheckBoxList(cblDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "CaseID", ""), "CaseName", "CaseID", false);
        cblDonor_Type.Items[0].Selected = false;
        //捐款方式
        Util.FillCheckBoxList(cblDonate_Payment, Util.GetDataTable("CaseCode", "GroupName", "捐款方式", "ABS(CaseID)", ""), "CaseName", "CaseID", false);
        cblDonate_Payment.Items[0].Selected = false;
        //捐款用途
        Util.FillCheckBoxList(cblDonate_Purpose, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseID", false);
        cblDonate_Purpose.Items[0].Selected = false;
        //捐款類別
        cblDonate_Type.Items.Insert(0, new ListItem("單次捐款", "單次捐款"));
        cblDonate_Type.Items.Insert(1, new ListItem("長期捐款", "長期捐款"));
        //入帳銀行
        //Util.FillCheckBoxList(cblAccoun_Bank, Util.GetDataTable("CaseCode", "GroupName", "入帳銀行", "", ""), "CaseName", "CaseID", false);
        //cblAccoun_Bank.Items[0].Selected = false;

        //會計科目
        //Util.FillCheckBoxList(cblAccounting_Title, Util.GetDataTable("CaseCode", "GroupName", "款項會計科目", "", ""), "CaseName", "CaseID", false);
        //cblAccounting_Title.Items[0].Selected = false;
        //其他
        //cblExp_Pre.Items.Insert(0, new ListItem("手開收據", "手開收據"));
        //cblExp_Pre.Items.Insert(1, new ListItem("作廢收據", "作廢收據"));

        //線上付款方式20150827新增
        Util.FillCheckBoxList(cblPayment_type, Util.GetDataTable("Donate_IePayType", "Display", "1", "", ""), "CodeName", "CodeID", false);
        cblPayment_type.Items[0].Selected = false;
    }
    //---------------------------------------------------------------------------
    //dropdownlist選縣市自動帶入鄉鎮市區
    protected void ddlCity_SelectedIndexChanged(object sender, EventArgs e)
    {
        Util.FillDropDownList(ddlArea, Util.GetDataTable("CodeCity", "ParentCityID", ddlCity.SelectedValue, "", ""), "Name", "ZipCode", false);
        ddlArea.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea.SelectedIndex = 0;
    }
    protected void ddlCity2_SelectedIndexChanged(object sender, EventArgs e)
    {
        Util.FillDropDownList(ddlArea2, Util.GetDataTable("CodeCity", "ParentCityID", ddlCity2.SelectedValue, "", ""), "Name", "ZipCode", false);
        ddlArea2.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea2.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    //查詢
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = @"select
	                    dr.Donor_Id as [捐款人編號],
                        dr.Donor_Name as [捐款人],
	                    CONVERT(VARCHAR(10) , do.Donate_Date, 111 ) as [捐款日期],
	                    REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,case when do.Issue_Type !='D' then do.Donate_Amt else '0' end),1),'.00','') as [捐款金額],
	                    do.Donate_Payment as [捐款方式],
	                    do.Donate_Purpose as [捐款用途],
	                    do.Invoice_Type as [收據開立],
	                    ISNULL(do.Invoice_Pre,'') + do.Invoice_No as [收據編號],
	                    --(case when CONVERT(VARCHAR(10) , do.Accoun_Date, 111 ) != '1900/01/01' then CONVERT(VARCHAR(10) , do.Accoun_Date, 111 ) else '' end ) as [沖帳日期],
                        do.Donate_Type as [捐款類別],
                        --do.Accoun_Bank as [入帳銀行],
                        --do.Accounting_Title as [會計科目],
                        --A.Act_ShortName as [活動名稱],
                        (Case when do.Issue_Type = 'D' then '作廢' when do.Issue_Type = 'M' then '手開' else '' end) as [狀態],Donate_Id
	                from Donate do 
                    join Dept dn on do.Dept_Id=dn.DeptId
                    left join Donor dr on dr.Donor_Id = do.Donor_Id
                    left join (select * from (select ROW_NUMBER() OVER(PARTITION by orderid ORDER BY Ser_No) AS ROWID,*
                         from DONATE_IEPAY) as P1 where ROWID = 1) as di on do.od_sob = di.orderid 
                    left join Donate_IePayType dp on di.paytype = dp.CodeID 
                    where dr.DeleteDate is null and do.Issue_Type !='D'";

        strSql += Condition();
        strSql += " order by do.Donate_Date desc";
        Session["strSql"] = strSql;
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
            lblAmt.Text = "0";
            return;
        }
        //Grid initial
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.DisableColumn.Add("Donate_Id");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
        npoGridView.Keys.Add("Donate_Id");
        npoGridView.EditLink = Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=");
        lblGridList.Text = npoGridView.Render();
        //捐款金額合計
        lblAmt.Text = Donate_Amt();
    }
    private string Condition()
    {
        string strSql = "";
        //if (ddlDept.SelectedIndex != 0)
        //{
        //    strSql += " and do.Dept_Id='" + ddlDept.SelectedValue + "' ";
        //}
        //身分別-----------------------------------------------------------------------------------//
        bool first = true; int cnt = 0;
        for (int i = 0; i < cblDonor_Type.Items.Count; i++)
        {
            if (cblDonor_Type.Items[i].Selected && first)
            {
                strSql += " and (dr.Donor_Type like '%" + cblDonor_Type.Items[i].ToString() + "%' "; first = false; cnt++;
            }
            else if (cblDonor_Type.Items[i].Selected)
            {
                strSql += " or dr.Donor_Type like '%" + cblDonor_Type.Items[i].ToString() + "%' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //通訊地址----------------------------------------------------------------------------------//
        if (ddlCity.SelectedIndex != 0)
        {
            strSql += " and dr.City= '" + ddlCity.SelectedValue + "'";
        }
        if (ddlArea.SelectedIndex != 0 && ddlArea.Items.Count != 0)
        {
            strSql += " and dr.Area like '%" + ddlArea.SelectedValue + "'";
        }
        //收據地址----------------------------------------------------------------------------------//
        if (ddlCity2.SelectedIndex != 0)
        {
            strSql += " and dr.invoice_City= '" + ddlCity2.SelectedValue + "'";
        }
        if (ddlArea2.SelectedIndex != 0 && ddlArea2.Items.Count != 0)
        {
            strSql += " and dr.invoice_Area like '%" + ddlArea2.SelectedValue + "'";
        }
        //捐款日期----------------------------------------------------------------------------------//
        if (tbxDonate_DateS.Text.Trim() != "")
        {
            strSql += " and do.Donate_Date >='" + tbxDonate_DateS.Text.Trim() + "'";
        }
        if (tbxDonate_DateE.Text.Trim() != "")
        {
            strSql += " and do.Donate_Date <='" + tbxDonate_DateE.Text.Trim() + "'";
        }
        //單筆金額----------------------------------------------------------------------------------//
        if (tbxDonate_AmtS.Text.Trim() != "")
        {
            strSql += " and cast(do.Donate_Amt as int) > '" + tbxDonate_AmtS.Text.Trim() + "'";
        }
        if (tbxDonate_AmtE.Text.Trim() != "")
        {
            strSql += " and cast(do.Donate_Amt as int) < '" + tbxDonate_AmtE.Text.Trim() + "'";
        }
        //捐款人編號
        if (txtDonor_IdS.Text.Trim() != "")
        {
            strSql += @" and do.Donor_Id >='" + txtDonor_IdS.Text.Trim() + "'";
        }
        if (txtDonor_IdE.Text.Trim() != "")
        {
            strSql += @" and do.Donor_Id <='" + txtDonor_IdE.Text.Trim() + "'";
        }
        //募款活動----------------------------------------------------------------------------------//
        //if (ddlActName.SelectedIndex != 0)
        //{
        //    strSql += " and do.Act_Id ='" + ddlActName.SelectedValue + "' ";
        //}
        //收據編號----------------------------------------------------------------------------------//
        if (tbxInvoice_NoS.Text.Trim() != "")
        {
            strSql += " and do.Invoice_No >='" + tbxInvoice_NoS.Text.Trim() + "'";
        }
        if (tbxInvoice_NoE.Text.Trim() != "")
        {
            strSql += " and do.Invoice_No <='" + tbxInvoice_NoE.Text.Trim() + "'";
        }
        //姓名--------------------------------------------------------------------------------------//
        //if (tbxDonor_Name.Text.Trim() != "")
        //{
        //    //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
        //    //strSql += " and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%' "; 
        //    strSql += " and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%' ";
        //}
        //捐款方式---------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblDonate_Payment.Items.Count; i++)
        {
            if (cblDonate_Payment.Items[i].Selected && first)
            {
                strSql += " and (do.Donate_Payment='" + cblDonate_Payment.Items[i].ToString() + "' "; first = false; cnt++;
            }
            else if (cblDonate_Payment.Items[i].Selected)
            {
                strSql += " or do.Donate_Payment='" + cblDonate_Payment.Items[i].ToString() + "' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //付款方式-----------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblPayment_type.Items.Count; i++)
        {
            if (cblPayment_type.Items[i].Selected && first)
            {
                strSql += "and (di.paytype='" + cblPayment_type.Items[i].Value.ToString() + "' "; first = false; cnt++;
            }
            else if (cblPayment_type.Items[i].Selected)
            {
                strSql += "or di.paytype='" + cblPayment_type.Items[i].Value.ToString() + "' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //捐款用途---------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblDonate_Purpose.Items.Count; i++)
        {
            if (cblDonate_Purpose.Items[i].Selected && first)
            {
                strSql += " and (do.Donate_Purpose='" + cblDonate_Purpose.Items[i].ToString() + "' "; first = false; cnt++;
            }
            else if (cblDonate_Purpose.Items[i].Selected)
            {
                strSql += " or do.Donate_Purpose='" + cblDonate_Purpose.Items[i].ToString() + "' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //捐款類別---------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblDonate_Type.Items.Count; i++)
        {
            if (cblDonate_Type.Items[i].Selected && first)
            {
                strSql += " and (do.Donate_Type='" + cblDonate_Type.Items[i].ToString() + "' "; first = false; cnt++;
            }
            else if (cblDonate_Type.Items[i].Selected)
            {
                strSql += " or do.Donate_Type='" + cblDonate_Type.Items[i].ToString() + "' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //入帳銀行---------------------------------------------------------------------------------//
        //first = true; cnt = 0;
        //for (int i = 0; i < cblAccoun_Bank.Items.Count; i++)
        //{
        //    if (cblAccoun_Bank.Items[i].Selected && first)
        //    {
        //        strSql += " and (do.Accoun_Bank='" + cblAccoun_Bank.Items[i].ToString() + "' "; first = false; cnt++;
        //    }
        //    else if (cblAccoun_Bank.Items[i].Selected)
        //    {
        //        strSql += " or do.Accoun_Bank='" + cblAccoun_Bank.Items[i].ToString() + "' "; cnt++;
        //    }
        //}
        //if (cnt != 0) strSql += ')';
        //沖帳日期----------------------------------------------------------------------------------//
        //if (tbxAccoun_DateS.Text.Trim() != "")
        //{
        //    strSql += " and do.Accoun_DateS >='" + tbxAccoun_DateS.Text.Trim() + "'";
        //}
        //if (tbxAccoun_DateE.Text.Trim() != "")
        //{
        //    strSql += " and do.Accoun_DateE <='" + tbxAccoun_DateE.Text.Trim() + "'";
        //}
        //會計科目---------------------------------------------------------------------------------//
        //first = true; cnt = 0;
        //for (int i = 0; i < cblAccounting_Title.Items.Count; i++)
        //{
        //    if (cblAccounting_Title.Items[i].Selected && first)
        //    {
        //        strSql += " and (do.Accounting_Title='" + cblAccounting_Title.Items[i].ToString() + "' "; first = false; cnt++;
        //    }
        //    else if (cblAccounting_Title.Items[i].Selected)
        //    {
        //        strSql += " or do.Accounting_Title='" + cblAccounting_Title.Items[i].ToString() + "' "; cnt++;
        //    }
        //}
        //if (cnt != 0) strSql += ')';
        //收據狀態-------------------------------------------------------------------------------------//
        //if (cblExp_Pre.Items[0].Selected)
        //{
        //    //20140514 修改「手開收據」判斷為Issue_Type = 'M'
        //    strSql += " and do.Issue_Type = 'M'";
        //}
        //else if (cblExp_Pre.Items[1].Selected)
        //{
        //    //20140514 修改「作廢收據」判斷為Issue_Type = 'D'
        //    strSql += " and do.Issue_Type = 'D'";
        //}
        //區域 20150827新增
        if (rblAddress.SelectedValue == "1")
        {
            strSql += " and  dr.IsAbroad = 'N'";
        }
        else if (rblAddress.SelectedValue == "2")
        {
            strSql += " and  dr.IsAbroad = 'Y'";
        }
        else
        {
            
        }
        //其他//20150827增加不含錯址、歿、不主動聯絡
        if (cbxIsErrAddress.Checked)
        {
            strSql += " and (IsErrAddress='N' or IsErrAddress is NULL)";
        }
        if (cbxNoDie.Checked)
        {
            strSql += " and IsNull(Sex,'') <> '歿'";
        }
        if (cbxIsContact.Checked)
        {
            strSql += " and IsContact='N'";
        }
        Session["Condition"] = strSql;
        return strSql;
    }
    private string Donate_Amt()
    {
        string strSql = "select REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(case when ISNULL(Issue_Type,'') != 'D' then Donate_Amt else 0 end)),1),'.00','') as Donate_Amt \n";
        strSql += " from donate do\n";
        strSql += " left join donor dr on do.Donor_Id = dr. Donor_Id\n";
        strSql += @" left join (select * from (select ROW_NUMBER() OVER(PARTITION by orderid ORDER BY Ser_No) AS ROWID,*
                         from DONATE_IEPAY) as P1 where ROWID = 1) as di on do.od_sob = di.orderid 
                     left join Donate_IePayType dp on di.paytype = dp.CodeID";
        strSql += " where dr.DeleteDate is null ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        strSql += Condition();

        DataTable dt;
        //20140425 Modify by GoodTV Tanya
        //dt = NpoDB.QueryGetScalar(strSql, dict);
        dt = NpoDB.GetDataTableS(strSql, dict);
        DataRow dr = dt.Rows[0];
        string Donate_Amt = dr["Donate_Amt"].ToString();
        return Donate_Amt;
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        HFD_Query_Flag.Value = "Y";
        LoadFormData();
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("../DonateMgr/DonateNameList_Print_Excel.aspx"));
    }
    protected void btnToFinancexls_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("../DonateMgr/DonateMonthQry_Print_Excel.aspx"));
    }
}