using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;

public partial class DonateMgr_Donate_New_InvoiceNo : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Donate_Id");
            LoadDropDownListData();
            Form_DataBind();
        }
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //捐款方式
        Util.FillDropDownList(ddlDonate_Payment, Util.GetDataTable("CaseCode", "GroupName", "捐款方式", "ABS(CaseID)", ""), "CaseName", "CaseName", false);
        ddlDonate_Payment.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Payment.SelectedIndex = 2;

        //捐款用途
        Util.FillDropDownList(ddlDonate_Purpose, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseName", false);
        ddlDonate_Purpose.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Purpose.SelectedIndex = 0;

        //入帳銀行
        Util.FillDropDownList(ddlAccoun_Bank, Util.GetDataTable("CaseCode", "GroupName", "入帳銀行", "", ""), "CaseName", "CaseName", false);
        ddlAccoun_Bank.Items.Insert(0, new ListItem("", ""));
        ddlAccoun_Bank.SelectedIndex = 0;

        //捐款類別
        ddlDonate_Type.Items.Insert(0, new ListItem("單次捐款", "單次捐款"));
        ddlDonate_Type.Items.Insert(1, new ListItem("長期捐款", "長期捐款"));
        ddlDonate_Type.SelectedIndex = 0;

        //會計科目
        Util.FillDropDownList(ddlAccounting_Title, Util.GetDataTable("CaseCode", "GroupName", "款項會計科目", "", ""), "CaseName", "CaseName", false);
        ddlAccounting_Title.Items.Insert(0, new ListItem("", ""));
        ddlAccounting_Title.SelectedIndex = 0;

        //募款活動
        Util.FillDropDownList(ddlAct_Id, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        ddlAct_Id.Items.Insert(0, new ListItem("", ""));
        ddlAct_Id.SelectedIndex = 0;

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
        strSql = @" select * , (Case When dr.Invoice_City='' Then Invoice_Address Else Case When B.mValue<>C.mValue Then B.mValue+dr.Invoice_ZipCode+C.mValue+Invoice_Address Else B.mValue+dr.Invoice_ZipCode+Invoice_Address End End) as [地址]
                    from Donate do 
                        Left Join Donor dr on do.Donor_Id = dr.Donor_ID 
                        Left Join Dept D on do.Dept_Id = D.DeptID left join Act A on do.Act_Id = A.Act_Id
                        Left Join CODECITY As B On dr.Invoice_City=B.mCode Left Join CODECITY As C On dr.Invoice_Area=C.mCode 
                    where dr.DeleteDate is null and Donate_Id='" + uid + "'";
        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("DonateInfo.aspx");

        DataRow dr = dt.Rows[0];
        
        //捐款人Id
        HFD_Donor_Id.Value = dr["Donor_Id"].ToString().Trim();
        //捐款人
        tbxDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
        //類別
        tbxCategory.Text = dr["Category"].ToString().Trim();
        //身份別
        tbxDonor_Type.Text = dr["Donor_Type"].ToString().Trim();
        //收據地址
        tbxAddress.Text = dr["地址"].ToString();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //提款日期
        tbxDonate_Date.Text = DateTime.Parse(dr["Donate_Date"].ToString()).ToString("yyyy/MM/dd");
        //捐款方式
        ddlDonate_Payment.Text = dr["Donate_Payment"].ToString().Trim();
        //捐款用途
        ddlDonate_Purpose.Text = dr["Donate_Purpose"].ToString().Trim();
        //收據開立
        tbxInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();
        //捐款金額
        tbxDonate_Amt.Text = (Convert.ToInt64(dr["Donate_Amt"])).ToString();
        //手續費
        tbxDonate_Fee.Text = (Convert.ToInt64(dr["Donate_Fee"])).ToString();
        //實收金額
        tbxDonate_Accou.Text = (Convert.ToInt64(dr["Donate_Accou"])).ToString();
        //外幣
        tbxDonate_Forign.Text = dr["Donate_Forign"].ToString();

        if (dr["Donate_Payment"].ToString().Trim() == "支票")
        {
            tbxCheck_No.Text = dr["Check_No"].ToString().Trim();
            tbxCheck_ExpireDate.Text = DateTime.Parse(dr["Check_ExpireDate"].ToString()).ToString("yyyy/MM/dd");
            PanelCheck.Visible = true;
            Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Check_ExpireDate();", true);
        }
        if (dr["Donate_Payment"].ToString().Trim() == "信用卡授權書(一般)" || dr["Donate_Payment"].ToString().Trim() == "信用卡授權書(聯信)")
        {
            //銀行別
            tbxCard_Bank.Text = dr["Card_Bank"].ToString().Trim();
            //信用卡別
            tbxCard_Type.Text = dr["Card_Type"].ToString().Trim();
            //信用卡號
            tbxAccount_No1.Text = dr["Account_No"].ToString().Substring(0, 4).Trim();
            tbxAccount_No2.Text = dr["Account_No"].ToString().Substring(4, 4).Trim();
            tbxAccount_No3.Text = dr["Account_No"].ToString().Substring(8, 4).Trim();
            if (dr["Donate_Payment"].ToString().Trim() == "信用卡授權書(聯信)")
            {
                tbxAccount_No4.Text = dr["Account_No"].ToString().Substring(12, 3).Trim();
            }
            else
            {
                tbxAccount_No4.Text = dr["Account_No"].ToString().Substring(12, 4).Trim();
            }
            //有效月年
            tbxValid_Date.Text = dr["Valid_Date"].ToString().Trim();
            //持卡人
            tbxCard_Owner.Text = dr["Card_Owner"].ToString().Trim();
            //持卡人身分證
            tbxOwner_IDNo.Text = dr["Owner_IDNo"].ToString().Trim();
            //與捐款人關係
            tbxRelation.Text = dr["Relation"].ToString().Trim();
            //授權碼
            tbxAuthorize.Text = dr["Authorize"].ToString().Trim();
            PanelCreditCard.Visible = true;
        }
        else if (dr["Donate_Payment"].ToString().Trim() == "郵局帳戶轉帳授權書")
        {
            //存簿戶名
            tbxPost_Name.Text = dr["Post_Name"].ToString().Trim();
            //持有人身分證
            tbxPost_IDNo.Text = dr["Post_IDNo"].ToString().Trim();
            //存簿局號
            tbxPost_SavingsNo.Text = dr["Post_SavingsNo"].ToString().Trim();
            //存簿帳號
            tbxPost_AccountNo.Text = dr["Post_AccountNo"].ToString().Trim();
            PanelAccount.Visible = true;
        }

        //機構
        tbxDept.Text = dr["DeptShortName"].ToString().Trim();
        //收據抬頭
        tbxInvoice_Title.Text = dr["Invoice_Title"].ToString().Trim();
        //收據編號
        // 2014/5/21 修正收據編號
        //tbxInvoice_No.Text = dr["Invoice_Pre"].ToString().Trim() + dr["Invoice_No"].ToString().Trim();
        tbxInvoice_No.Text = "重新取號";
        HFD_Invoice_No.Value = dr["Invoice_No"].ToString().Trim();
        //請款日期
        if (dr["Request_Date"].ToString() != "" && DateTime.Parse(dr["Request_Date"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxRequest_Date.Text = DateTime.Parse(dr["Request_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //入帳銀行
        ddlAccoun_Bank.Text = dr["Accoun_Bank"].ToString().Trim();
        //沖帳日期
        if (dr["Accoun_Date"].ToString() != "" && DateTime.Parse(dr["Accoun_Date"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxAccoun_Date.Text = DateTime.Parse(dr["Accoun_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //捐款類別
        ddlDonate_Type.Text = dr["Donate_Type"].ToString().Trim();
        //傳票號碼
        tbxDonation_NumberNo.Text = dr["Donation_NumberNo"].ToString().Trim();
        //劃撥 / 匯款單號
        tbxDonation_SubPoenaNo.Text = dr["Donation_SubPoenaNo"].ToString().Trim();
        //會計科目
        ddlAccounting_Title.Text = dr["Accounting_Title"].ToString().Trim();
        //收據寄送
        if (dr["InvoceSend_Date"].ToString() != "" && DateTime.Parse(dr["InvoceSend_Date"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxInvoceSend_Date.Text = tbxAccoun_Date.Text = DateTime.Parse(dr["InvoceSend_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //專案活動
        ddlAct_Id.SelectedValue = dr["Act_Id"].ToString().Trim();
        //捐款備註
        tbxComment.Text = dr["Comment"].ToString().Trim();
        //收據備註
        tbxInvoice_PrintComment.Text = dr["Invoice_PrintComment"].ToString().Trim();
    }
    protected void btn_Re_No_Click(object sender, EventArgs e)
    {
        bool check_close = Util.Get_Close("1", SessionInfo.DeptID, tbxDonate_Date.Text, SessionInfo.UserID);
        if (check_close == false)
        {
            //作廢舊的
            //****變數宣告****//
            Dictionary<string, object> dict = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql = " update Donate set ";
            //strSql += " Export= @Export";
            strSql += " Issue_Type= @Issue_Type";
            strSql += " where Donate_Id = @Donate_Id";
            //dict.Add("Export", "Y");
            dict.Add("Issue_Type", "D");
            dict.Add("Donate_Id", HFD_Uid.Value);
            NpoDB.ExecuteSQLS(strSql, dict);

            //新增新的
            bool flag = false;
            try
            {
                Donate_AddNew();
                flag = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (flag == true)
            {
                Donor_Edit();
                SetSysMsg("捐款資料重新取號成功！");
                // 2014/4/10 修改成 導向捐款紀錄頁面
                ////******設定收據編號******//
                //string strSql2 = @"select Donate_Id from Donate
                //    where Invoice_No ='" + HFD_Invoice_No.Value + "'";
                ////****執行查詢流水號語法****//
                //DataTable dt2 = NpoDB.QueryGetTable(strSql2);
                //DataRow dr2 = dt2.Rows[0];
                //string Donate_Id = dr2["Donate_Id"].ToString().Trim();

                //Response.Redirect(Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=" + Donate_Id));
                Response.Redirect(Util.RedirectByTime("DonateDataList.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
            }
        }
        else
        {
            AjaxShowMessage("您輸入的捐款日期『" + tbxDonate_Date.Text + "』 已關帳無法重新取號 ！");
        }
    }
    protected void btnExit_Click1(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=" + HFD_Uid.Value));
    }
    public void Donate_AddNew()
    {
        string strSql = "insert into  Donate\n";
        strSql += "( Donor_Id, Donate_Date, Donate_Payment, Donate_Purpose, Invoice_Type, Donate_Amt, Donate_Fee, Donate_Accou,\n";
        strSql += " Donate_Forign, Dept_Id, Invoice_Title, Invoice_Pre, Invoice_No, Request_Date, Accoun_Bank, Accoun_Date,\n";
        strSql += " Donate_Type, Donation_NumberNo, Donation_SubPoenaNo, Accounting_Title, InvoceSend_Date, Act_Id,\n";
        strSql += " Check_No, Check_ExpireDate, Card_Bank, Card_Type, Account_No, Valid_Date, Card_Owner, Owner_IDNo, Relation,\n";
        strSql += " Authorize, Post_Name, Post_IDNo, Post_SavingsNo, Post_AccountNo,\n";
        strSql += " Comment, Invoice_PrintComment, Issue_Type, Issue_Type_Keep, Export, Create_Date, Create_DateTime, Create_User, Create_IP, Invoice_No_Old) values\n";
        strSql += "( @Donor_Id,@Donate_Date,@Donate_Payment,@Donate_Purpose,@Invoice_Type,@Donate_Amt,@Donate_Fee,@Donate_Accou,\n";
        strSql += "@Donate_Forign,@Dept_Id,@Invoice_Title,@Invoice_Pre,@Invoice_No,@Request_Date,@Accoun_Bank,@Accoun_Date,\n";
        strSql += "@Donate_Type,@Donation_NumberNo,@Donation_SubPoenaNo,@Accounting_Title,@InvoceSend_Date,@Act_Id,\n";
        strSql += "@Check_No,@Check_ExpireDate,@Card_Bank,@Card_Type,@Account_No,@Valid_Date,@Card_Owner,@Owner_IDNo,@Relation,\n";
        strSql += "@Authorize,@Post_Name,@Post_IDNo,@Post_SavingsNo,@Post_AccountNo,\n";
        strSql += "@Comment,@Invoice_PrintComment,@Issue_Type,@Issue_Type_Keep,@Export,@Create_Date,@Create_DateTime,@Create_User,@Create_IP, @Invoice_No_Old) ";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", HFD_Donor_Id.Value);
        dict.Add("Donate_Date", tbxDonate_Date.Text.Trim());
        dict.Add("Donate_Payment", ddlDonate_Payment.SelectedItem.Text);
        dict.Add("Donate_Purpose", ddlDonate_Purpose.SelectedItem.Text);
        dict.Add("Invoice_Type", tbxInvoice_Type.Text.Trim());
        dict.Add("Donate_Amt", tbxDonate_Amt.Text.Trim());
        dict.Add("Donate_Fee", tbxDonate_Fee.Text.Trim());
        dict.Add("Donate_Accou", tbxDonate_Accou.Text.Trim());
        dict.Add("Donate_Forign", tbxDonate_Forign.Text.Trim());
        dict.Add("Dept_Id", SessionInfo.DeptID);
        dict.Add("Invoice_Title", tbxInvoice_Title.Text.Trim());
        //收據編號
        dict.Add("Invoice_Pre", "A");
        //******設定收據編號******//
        String Year, Month, Day;
        Year = DateTime.Parse(tbxDonate_Date.Text).ToString("yyyy");
        Month = DateTime.Parse(tbxDonate_Date.Text).ToString("MM");
        Day = DateTime.Parse(tbxDonate_Date.Text).ToString("dd");
        /*string strSql2 = @"select isnull(MAX(Invoice_No),'') as Invoice_No from Donate
                    where Invoice_No like '%" + DateTime.Now.ToString("yyyyMMdd") + "%'";*/
        //20140804 收據編號更改規則，以捐款日期為依據來編號 by 詩儀
        string strSql2 = @"select isnull(MAX(Invoice_No),'') as Invoice_No from Donate
                        where Invoice_No like '" + Year + Month + Day + "%'";
        //****執行查詢流水號語法****//
        DataTable dt2 = NpoDB.QueryGetTable(strSql2);
        string Invoice_No = "";
        if (dt2.Rows.Count > 0)
        {
            string Invoice_No_value = dt2.Rows[0]["Invoice_No"].ToString();

            if (Invoice_No_value == "")
            {
                //Invoice_No = DateTime.Now.ToString("yyyyMMdd") + "0001";
                Invoice_No = DateTime.Parse(tbxDonate_Date.Text).ToString("yyyyMMdd") + "0001";
            }
            else
            {
                Invoice_No = (Convert.ToInt64(Invoice_No_value) + 1).ToString();
            }
        }
        //************************//
        dict.Add("Invoice_No", Invoice_No);
        HFD_Invoice_No.Value = Invoice_No;
        //----------------------------------------------------------------------
        //2014/5/13新增帶入資料
        //支票
        dict.Add("Check_No", tbxCheck_No.Text.Trim());
        dict.Add("Check_ExpireDate", tbxCheck_ExpireDate.Text.Trim());
        //信用卡授權書&ACH轉帳授權書
        dict.Add("Card_Bank", tbxCard_Bank.Text.Trim());
        dict.Add("Card_Type", tbxCard_Type.Text.Trim());
        dict.Add("Account_No", tbxAccount_No1.Text.Trim() + tbxAccount_No2.Text.Trim() + tbxAccount_No3.Text.Trim() + tbxAccount_No4.Text.Trim());
        dict.Add("Valid_Date", tbxValid_Date.Text.Trim());
        dict.Add("Card_Owner", tbxCard_Owner.Text.Trim());
        dict.Add("Owner_IDNo", tbxOwner_IDNo.Text.Trim());
        dict.Add("Relation", tbxRelation.Text.Trim());
        dict.Add("Authorize", tbxAuthorize.Text.Trim());
        //郵局轉帳授權書
        dict.Add("Post_Name", tbxPost_Name.Text.Trim());
        dict.Add("Post_IDNo", tbxPost_IDNo.Text.Trim());
        dict.Add("Post_SavingsNo", tbxPost_SavingsNo.Text.Trim());
        dict.Add("Post_AccountNo", tbxPost_AccountNo.Text.Trim());
        //----------------------------------------------------------------------
        dict.Add("Request_Date", tbxRequest_Date.Text.Trim());
        dict.Add("Accoun_Bank", ddlAccoun_Bank.SelectedItem.Text);
        dict.Add("Accoun_Date", tbxAccoun_Date.Text.Trim());
        dict.Add("Donate_Type", ddlDonate_Type.SelectedItem.Text);
        dict.Add("Donation_NumberNo", tbxDonation_NumberNo.Text.Trim());
        dict.Add("Donation_SubPoenaNo", tbxDonation_SubPoenaNo.Text.Trim());
        dict.Add("Accounting_Title", ddlAccounting_Title.SelectedItem.Text);
        dict.Add("InvoceSend_Date", tbxInvoceSend_Date.Text.Trim());
        if (ddlAct_Id.SelectedIndex != 0)
        {
            dict.Add("Act_Id", ddlAct_Id.SelectedValue);
        }
        else
        {
            dict.Add("Act_Id", "");
        }
        dict.Add("Comment", tbxComment.Text.Trim());
        dict.Add("Invoice_PrintComment", tbxInvoice_PrintComment.Text.Trim());
        dict.Add("Issue_Type", "");
        dict.Add("Issue_Type_Keep", "");
        dict.Add("Export", "N");
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        dict.Add("Invoice_No_Old", tbxInvoice_No.Text.Trim());
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    //20140812 更新最後捐款日期 by 詩儀
    public void Donor_Edit()
    {
        //先抓出第一次/最近一次的捐款日期、捐款次數和捐款總額
        string Begin_DonateDate, Last_DonateDate, Donate_Total_S;
        int Donate_No;
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;
        //****變數設定****//
        uid = HFD_Donor_Id.Value;
        //****設定查詢****//
        strSql = " select *  from Donor  where Donor_Id='" + uid + "'";
        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];

        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        string strSql2 = "";
        Begin_DonateDate = dr["Begin_DonateDate"].ToString();
        Last_DonateDate = DateTime.Parse(dr["Last_DonateDate"].ToString()).ToString("yyyy/MM/dd");
        Donate_No = Int16.Parse(dr["Donate_No"].ToString());
        Donate_Total_S = (Convert.ToInt64(dr["Donate_Total"])).ToString();
        int Donate_Total = Int32.Parse(Donate_Total_S);
        //更新以上欄位

        strSql2 = " update Donor set ";

        //strSql2 += " Donate_No = @Donate_No";
        //strSql2 += ", Donate_Total = @Donate_Total";
        //if (DateTime.Parse(DateTime.Parse(Last_DonateDate).ToString("yyyy/MM/dd")) < DateTime.Parse(tbxDonate_Date.Text.Trim()))
        //{
            strSql2 += "Last_DonateDate = @Last_DonateDate";
            //dict2.Add("Last_DonateDate", DateTime.Now.ToString("yyyy-MM-dd"));
            dict2.Add("Last_DonateDate", tbxDonate_Date.Text.Trim());
        //}
        strSql2 += " where Donor_Id = @Donor_Id";

        //dict2.Add("Donate_No", Donate_No + 1);
        //dict2.Add("Donate_Total", Donate_Total + int.Parse(tbxDonate_Accou.Text.Trim()));
        dict2.Add("Donor_Id", HFD_Donor_Id.Value);
        NpoDB.ExecuteSQLS(strSql2, dict2);
    }
}