using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class DonateMgr_Donate_Add : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Donor_Id");
            LoadDropDownListData();
            Form_DataBind();
        }
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " ReadOnly();", true);
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
        // 2014/4/8 更正沒有預計空白
        //ddlDonate_Purpose.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Purpose.SelectedIndex = 0;

        //收據開立
        Util.FillDropDownList(ddlInvoice_Type, Util.GetDataTable("CaseCode", "GroupName", "收據開立", "", ""), "CaseName", "CaseName", false);
        ddlInvoice_Type.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Type.SelectedIndex = 2;

        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.SelectedIndex = 0;

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

        //信用卡別  
        Util.FillDropDownList(ddlCard_Type, Util.GetDataTable("CaseCode", "GroupName", "信用卡別", "", ""), "CaseName", "CaseName", false);
        ddlCard_Type.Items.Insert(0, new ListItem("", ""));
        ddlCard_Type.SelectedIndex = 0;

        //有效月年
        Util.FillDropDownList(ddlMonth_Valid_Date, 1, 12, true, 0);
        Util.FillDropDownList(ddlYear_Valid_Date, Int32.Parse(DateTime.Now.Year.ToString().Substring(2, 2)), Int32.Parse(DateTime.Now.Year.ToString().Substring(2, 2)) + 15, true, 0);

        //外幣幣別
        Util.FillDropDownList(ddlDonate_Forign, Util.GetDataTable("CaseCode", "GroupName", "外幣幣別", "", ""), "CaseName", "CaseName", false);
        ddlDonate_Forign.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Forign.SelectedIndex = 0;
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
        strSql = @" Select * ,(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.Invoice_ZipCode+B.mValue+Invoice_Address Else A.mValue+DONOR.Invoice_ZipCode+Invoice_Address End End) as [地址]
                    From Donor 
                        Left Join CODECITY As A On Donor.Invoice_City=A.mCode Left Join CODECITY As B On Donor.Invoice_Area=B.mCode 
                    where Donor_id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("DonateInfo.aspx");

        DataRow dr = dt.Rows[0];

        //捐款人
        tbxDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
        //捐款人編號
        tbxDonor_Id.Text = dr["Donor_Id"].ToString().Trim();
        //身份別
        tbxDonor_Type.Text = dr["Donor_Type"].ToString().Trim();
        //收據地址
        tbxAddress.Text = dr["地址"].ToString();
        //收據開立
        tbxInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //最近捐款日：
        if (dr["Last_DonateDate"].ToString().Trim() != "")
        {
            tbxLast_DonateDate.Text = DateTime.Parse(dr["Last_DonateDate"].ToString().Trim()).ToShortDateString().ToString();
        }
        if (Session["Donate_Date"] != null)
        {
            tbxDonate_Date.Text = Session["Donate_Date"].ToString();
        }
        else
        {
            tbxDonate_Date.Text = DateTime.Now.ToShortDateString().ToString(); 
        }
        if (Session["Donate_Payment"] != null)
        {
            ddlDonate_Payment.Items.FindByText("匯款").Selected = false;
            ddlDonate_Payment.Items.FindByText(Session["Donate_Payment"].ToString().Trim()).Selected = true;
            switch (Session["Donate_Payment"].ToString())
            {
                case "支票":
                    PanelCheck.Visible = true;
                    break;
                case "信用卡授權書(一般)":
                    PanelCreditCard.Visible = true;
                    break;
                case "信用卡授權書(聯信)":
                    PanelCreditCard.Visible = true;
                    break;
                case "郵局帳戶轉帳授權書":
                    PanelAccount.Visible = true;
                    break;
                case "ACH轉帳授權書":
                    PanelACH.Visible = true;
                    break;
            }
        }
        if (Session["Donate_Purpose"] != null)
        {
            ddlDonate_Purpose.Items.FindByText("經常費").Selected = false;
            ddlDonate_Purpose.Items.FindByText(Session["Donate_Purpose"].ToString().Trim()).Selected = true;
        }
        
        //收據抬頭
        tbxInvoice_Title.Text = dr["Invoice_Title"].ToString().Trim();

        // 2014/4/8 修改可帶入收據維護紀錄的收據開立
        //2014/4/25 Modify by GoodTV Tanya:預設帶入「捐款人基本資料」的「收據開立」
        //ddlInvoice_Type.SelectedItem.Text = dr["Invoice_Type"].ToString().Trim();
        ddlInvoice_Type.SelectedValue = dr["Invoice_Type"].ToString().Trim();

    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        bool check_close = Util.Get_Close("1", SessionInfo.DeptID, tbxDonate_Date.Text, SessionInfo.UserID);
        if (check_close == false)
        {
            bool flag = false;
            try
            {
                Donate_AddNew();
                if (this.DonateMotive1.Checked || this.DonateMotive2.Checked || this.DonateMotive3.Checked || this.DonateMotive4.Checked || this.DonateMotive5.Checked || this.WatchMode1.Checked || this.WatchMode2.Checked || this.WatchMode3.Checked || this.WatchMode4.Checked || this.WatchMode5.Checked || this.WatchMode6.Checked || this.WatchMode7.Checked || this.WatchMode8.Checked || this.WatchMode9.Checked || txtToGoodTV.Text !="")
                {
                    Questionaire_Add();
                }
                flag = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (flag == true)
            {
                //Donor 資料表中的捐款資料增加
                Donor_Edit();
                SetSysMsg("捐款資料新增成功！");
                //******設定收據編號******//
                string strSql3 = @"select Donate_Id from Donate
                    where Invoice_No ='" + HFD_Invoice_No.Value + "'";
                //****執行查詢流水號語法****//
                DataTable dt3 = NpoDB.QueryGetTable(strSql3);
                DataRow dr3 = dt3.Rows[0];
                string Donate_Id = dr3["Donate_Id"].ToString().Trim();

                if (Session["cType"] == "DonateInfo")
                {
                    Response.Redirect(Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=" + Donate_Id));                    
                }
                else if (Session["cType"] == "DonorInfo_Edit")
                {
                    //Response.Redirect(Util.RedirectByTime("../DonorMgr/DonorInfo_Edit.aspx", "Donor_Id=" + HFD_Uid.Value));
                    Response.Redirect(Util.RedirectByTime("DonateDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
                }
                else if (Session["cType"] == "DonateDataList")
                {
                    Response.Redirect(Util.RedirectByTime("DonateDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
                }
                else
                {
                    //20140428 Modify by GoodTV Tanya:新增完，預設導到「捐款紀錄」
                    //Response.Redirect(Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=" + Donate_Id));
                    Response.Redirect(Util.RedirectByTime("DonateDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
                }
            }
        }
        else
        {
            ShowSysMsg("您輸入的捐款日期『" + tbxDonate_Date.Text + "』 已關帳無法新增 ！");
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        if (Session["cType"] == "DonateInfo")
        {
            Response.Redirect(Util.RedirectByTime("DonateInfo.aspx"));
        }
        else if (Session["cType"] == "DonorInfo_Edit")
        {
            Response.Redirect(Util.RedirectByTime("../DonorMgr/DonorInfo_Edit.aspx", "Donor_Id=" + HFD_Uid.Value));
        }
        else if (Session["cType"] == "DonateDataList")
        {
            Response.Redirect(Util.RedirectByTime("DonateDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
        }
        else
        {
            Response.Redirect(Util.RedirectByTime("DonateInfo.aspx"));
        }
    }
    protected void ddlDonate_Payment_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlDonate_Payment.SelectedItem.Text == "支票")
        {
            PanelCheck.Visible = true;
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = false;
            PanelACH.Visible = false;
            Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Check_ExpireDate();", true);
            tbxCheck_No.Text = "";
            tbxCheck_ExpireDate.Text = "";
        }
        else if (ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(一般)" || ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(聯信)")
        {
            PanelCheck.Visible = false;
            PanelCreditCard.Visible = true;
            PanelAccount.Visible = false;
            PanelACH.Visible = false;
            tbxCard_Bank.Text = "";
            if (ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(一般)")
            {
                ddlCard_Type.SelectedIndex = 0;
            }
            else
            {
                ddlCard_Type.SelectedIndex = 3;
            }
            tbxAccount_No1.Text = "";
            tbxAccount_No2.Text = "";
            tbxAccount_No3.Text = "";
            tbxAccount_No4.Text = "";
            ddlMonth_Valid_Date.SelectedIndex = 0;
            ddlYear_Valid_Date.SelectedIndex = 0;
            tbxCard_Owner.Text = "";
            tbxOwner_IDNo.Text = "";
            tbxRelation.Text = "";
            tbxAuthorize.Text = "";
        }
        else if (ddlDonate_Payment.SelectedItem.Text == "郵局帳戶轉帳授權書")
        {
            PanelCheck.Visible = false;
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = true;
            PanelACH.Visible = false;
            tbxPost_Name.Text = "";
            tbxPost_IDNo.Text = "";
            tbxPost_SavingsNo.Text = "";
            tbxPost_AccountNo.Text = "";
        }
        else if (ddlDonate_Payment.SelectedItem.Text == "ACH轉帳授權書")
        {
            PanelCheck.Visible = false;
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = false;
            PanelACH.Visible = true;
            tbxP_BANK.Text = "";
            tbxP_RCLNO.Text = "";
            tbxP_PID.Text = "";
        }
        else
        {
            PanelCheck.Visible = false;
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = false;
            PanelACH.Visible = false;
        }
    }
    public void Donate_AddNew()
    {
        string strSql = "insert into  Donate\n";
        strSql += "( Donor_Id, Donate_Date, Donate_Payment, Donate_Purpose, Invoice_Type, Donate_Amt, Donate_Fee, Donate_Accou,\n";
        strSql += " Donate_Forign, Donate_ForignAmt, Dept_Id, Invoice_Title, Issue_Type, Issue_Type_Keep, Invoice_Pre, Invoice_No, Request_Date, Accoun_Bank, Accoun_Date,\n";
        strSql += " Donate_Type, Donation_NumberNo, Donation_SubPoenaNo, Accounting_Title, InvoceSend_Date, Act_Id,\n";
        strSql += " Comment, Invoice_PrintComment, Export, Create_Date, Create_DateTime, Create_User, Create_IP,\n";
        strSql += " Check_No, Check_ExpireDate, Card_Bank, Card_Type, Account_No, Valid_Date, Card_Owner, Owner_IDNo, Relation, \n";
        strSql += " Authorize, Post_Name, Post_IDNo, Post_SavingsNo, Post_AccountNo, P_BANK, P_RCLNO, P_PID) values\n";
        strSql += "( @Donor_Id,@Donate_Date,@Donate_Payment,@Donate_Purpose,@Invoice_Type,@Donate_Amt,@Donate_Fee,@Donate_Accou,\n";
        strSql += "@Donate_Forign,@Donate_ForignAmt,@Dept_Id,@Invoice_Title,@Issue_Type,@Issue_Type_Keep,@Invoice_Pre,@Invoice_No,@Request_Date,@Accoun_Bank,@Accoun_Date,\n";
        strSql += "@Donate_Type,@Donation_NumberNo,@Donation_SubPoenaNo,@Accounting_Title,@InvoceSend_Date,@Act_Id,\n";
        strSql += "@Comment,@Invoice_PrintComment,@Export,@Create_Date,@Create_DateTime,@Create_User,@Create_IP,\n";
        strSql += "@Check_No,@Check_ExpireDate,@Card_Bank,@Card_Type,@Account_No,@Valid_Date,@Card_Owner,@Owner_IDNo,@Relation, \n";
        strSql += "@Authorize,@Post_Name,@Post_IDNo,@Post_SavingsNo,@Post_AccountNo,@P_BANK,@P_RCLNO,@P_PID) ";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", HFD_Uid.Value);
        dict.Add("Donate_Date", tbxDonate_Date.Text.Trim());
        dict.Add("Donate_Payment", ddlDonate_Payment.SelectedItem.Text);
        dict.Add("Donate_Purpose", ddlDonate_Purpose.SelectedItem.Text);
        Session["Donate_Date"] = tbxDonate_Date.Text.Trim();
        Session["Donate_Payment"] = ddlDonate_Payment.SelectedItem.Text;
        Session["Donate_Purpose"] = ddlDonate_Purpose.SelectedItem.Text;
        dict.Add("Invoice_Type", ddlInvoice_Type.SelectedItem.Text);
        dict.Add("Donate_Amt", tbxDonate_Amt.Text.Trim());
        dict.Add("Donate_Fee", tbxDonate_Fee.Text.Trim());
        dict.Add("Donate_Accou", tbxDonate_Accou.Text.Trim());
        dict.Add("Donate_Forign", ddlDonate_Forign.SelectedItem.Text);
        dict.Add("Donate_ForignAmt", tbxDonate_ForignAmt.Text.Trim());
        if (ddlDonate_Payment.SelectedItem.Text == "支票")
        {
            dict.Add("Check_No", tbxCheck_No.Text.Trim());
            dict.Add("Check_ExpireDate", tbxCheck_ExpireDate.Text.Trim());

            dict.Add("Card_Bank", "");
            dict.Add("Card_Type", "");
            dict.Add("Account_No", "");
            dict.Add("Valid_Date", "");
            dict.Add("Card_Owner", "");
            dict.Add("Owner_IDNo", "");
            dict.Add("Relation", "");
            dict.Add("Authorize", "");

            dict.Add("Post_Name", "");
            dict.Add("Post_IDNo", "");
            dict.Add("Post_SavingsNo", "");
            dict.Add("Post_AccountNo", "");

            dict.Add("P_BANK", "");
            dict.Add("P_RCLNO", "");
            dict.Add("P_PID", "");
        }
        else if (ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(一般)" || ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(聯信)")
        {
            dict.Add("Check_No", "");
            dict.Add("Check_ExpireDate", "");

            dict.Add("Card_Bank", tbxCard_Bank.Text.Trim());
            dict.Add("Card_Type", ddlCard_Type.SelectedItem.Text);
            dict.Add("Account_No", tbxAccount_No1.Text.Trim() + tbxAccount_No2.Text.Trim() + tbxAccount_No3.Text.Trim() + tbxAccount_No4.Text.Trim());
            dict.Add("Valid_Date", ddlMonth_Valid_Date.SelectedItem.Text.PadLeft(2, '0') + ddlYear_Valid_Date.SelectedItem.Text.PadLeft(2, '0'));
            dict.Add("Card_Owner", tbxCard_Owner.Text.Trim());
            dict.Add("Owner_IDNo", tbxOwner_IDNo.Text.Trim());
            dict.Add("Relation", tbxRelation.Text.Trim());
            dict.Add("Authorize", tbxAuthorize.Text.Trim());

            dict.Add("Post_Name", "");
            dict.Add("Post_IDNo", "");
            dict.Add("Post_SavingsNo", "");
            dict.Add("Post_AccountNo", "");

            dict.Add("P_BANK", "");
            dict.Add("P_RCLNO", "");
            dict.Add("P_PID", "");
        }
        else if (ddlDonate_Payment.SelectedItem.Text == "郵局帳戶轉帳授權書")
        {
            dict.Add("Check_No", "");
            dict.Add("Check_ExpireDate", "");

            dict.Add("Card_Bank", "");
            dict.Add("Card_Type", "");
            dict.Add("Account_No", "");
            dict.Add("Valid_Date", "");
            dict.Add("Card_Owner", "");
            dict.Add("Owner_IDNo", "");
            dict.Add("Relation", "");
            dict.Add("Authorize", "");

            dict.Add("Post_Name", tbxPost_Name.Text.Trim());
            dict.Add("Post_IDNo", tbxPost_IDNo.Text.Trim());
            dict.Add("Post_SavingsNo", tbxPost_SavingsNo.Text.Trim());
            dict.Add("Post_AccountNo", tbxPost_AccountNo.Text.Trim());

            dict.Add("P_BANK", "");
            dict.Add("P_RCLNO", "");
            dict.Add("P_PID", "");
        }
        //2014/5/15新增
        else if (ddlDonate_Payment.SelectedItem.Text == "ACH轉帳授權書")
        {
            dict.Add("Check_No", "");
            dict.Add("Check_ExpireDate", "");

            dict.Add("Card_Bank", "");
            dict.Add("Card_Type", "");
            dict.Add("Account_No", "");
            dict.Add("Valid_Date", "");
            dict.Add("Card_Owner", "");
            dict.Add("Owner_IDNo", "");
            dict.Add("Relation", "");
            dict.Add("Authorize", "");

            dict.Add("Post_Name", "");
            dict.Add("Post_IDNo", "");
            dict.Add("Post_SavingsNo", "");
            dict.Add("Post_AccountNo", "");

            dict.Add("P_BANK", tbxP_BANK.Text.Trim());
            dict.Add("P_RCLNO", tbxP_RCLNO.Text.Trim());
            dict.Add("P_PID", tbxP_PID.Text.Trim());
        }
        else
        {
            dict.Add("Check_No", "");
            dict.Add("Check_ExpireDate", "");

            dict.Add("Card_Bank", "");
            dict.Add("Card_Type", "");
            dict.Add("Account_No", "");
            dict.Add("Valid_Date", "");
            dict.Add("Card_Owner", "");
            dict.Add("Owner_IDNo", "");
            dict.Add("Relation", "");
            dict.Add("Authorize", "");

            dict.Add("Post_Name", "");
            dict.Add("Post_IDNo", "");
            dict.Add("Post_SavingsNo", "");
            dict.Add("Post_AccountNo", "");

            dict.Add("P_BANK", "");
            dict.Add("P_RCLNO", "");
            dict.Add("P_PID", "");
        }
        dict.Add("Dept_Id", SessionInfo.DeptID);
        dict.Add("Invoice_Title", tbxInvoice_Title.Text.Trim());
        //收據編號
        if (cbxInvoice_Pre.Checked == false)
        {
            //20140416 Add by GoodTV Tanya:增加紀錄是否為手開收據欄位
            dict.Add("Issue_Type", "");
            dict.Add("Issue_Type_Keep", "");
            dict.Add("Invoice_Pre", "A");
            String Year,Month,Day;
            Year = DateTime.Parse(tbxDonate_Date.Text).ToString("yyyy");
            Month = DateTime.Parse(tbxDonate_Date.Text).ToString("MM");
            Day = DateTime.Parse(tbxDonate_Date.Text).ToString("dd");
            //******設定收據編號******//
            //20140428 Modify by GoodTV Tanya:修改Get Invoice_No條件為like 'YYYYMMDD%'
            /*string strSql2 = @"select isnull(MAX(Invoice_No),'') as Invoice_No from Donate
                        where Invoice_No like '" + DateTime.Now.ToString("yyyyMMdd") + "%'";*/
            //20140804 收據編號更改規則，以捐款日期為依據來編號 by 詩儀
            string strSql2 = @"select isnull(MAX(Invoice_No),'') as Invoice_No from Donate
                        where Invoice_No like '" + Year+Month+Day + "%'";
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
        }
        if (cbxInvoice_Pre.Checked == true)
        {
            //20140416 Modify by GoodTV Tanya:修改手開收據規則
            dict.Add("Issue_Type", "M");
            dict.Add("Issue_Type_Keep", "M");
            dict.Add("Invoice_Pre", "");
            dict.Add("Invoice_No", tbxInvoice_No.Text.Trim());
            HFD_Invoice_No.Value = tbxInvoice_No.Text.Trim();
        }
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
        dict.Add("Export", "N");
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void Donor_Edit()
    {
        //先抓出第一次/最近一次的捐款日期、捐款次數和捐款總額
        string Begin_DonateDate, Last_DonateDate, Donate_Total_S;
        int Donate_No;
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;
        //****變數設定****//
        uid = HFD_Uid.Value;
        //****設定查詢****//
        strSql = " select *  from Donor  where Donor_Id='" + uid + "'";
        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];

        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        string strSql2 = "";
        if (dr["Begin_DonateDate"].ToString() == "" || DateTime.Parse(dr["Begin_DonateDate"].ToString()).ToString("yyyy/MM/dd") == "1900/01/01")
        {
            //初次捐款
            strSql2 = " update Donor set ";
            strSql2 += "  Begin_DonateDate = @Begin_DonateDate";
            strSql2 += ", Last_DonateDate = @Last_DonateDate";
            strSql2 += ", Donate_No = @Donate_No";
            strSql2 += ", Donate_Total = @Donate_Total";
            strSql2 += " where Donor_Id = @Donor_Id";

            dict2.Add("Begin_DonateDate", tbxDonate_Date.Text.Trim());
            dict2.Add("Last_DonateDate", tbxDonate_Date.Text.Trim());
            dict2.Add("Donate_No", "1");
            dict2.Add("Donate_Total", tbxDonate_Accou.Text.Trim());
            dict2.Add("Donor_Id", HFD_Uid.Value);
        }
        else
        {
            Begin_DonateDate = dr["Begin_DonateDate"].ToString();
            Last_DonateDate = DateTime.Parse(dr["Last_DonateDate"].ToString()).ToString("yyyy/MM/dd");
            Donate_No = Int16.Parse(dr["Donate_No"].ToString());
            Donate_Total_S = (Convert.ToInt64(dr["Donate_Total"])).ToString();
            int Donate_Total = Int32.Parse(Donate_Total_S);
            //更新以上欄位

            strSql2 = " update Donor set ";
            
            strSql2 += " Donate_No = @Donate_No";
            strSql2 += ", Donate_Total = @Donate_Total";
            if (DateTime.Parse(DateTime.Parse(Last_DonateDate).ToString("yyyy/MM/dd")) < DateTime.Parse(tbxDonate_Date.Text.Trim()))
            {
                strSql2 += ", Last_DonateDate = @Last_DonateDate";
                //dict2.Add("Last_DonateDate", DateTime.Now.ToString("yyyy-MM-dd"));
                dict2.Add("Last_DonateDate", tbxDonate_Date.Text.Trim());
            }
            strSql2 += " where Donor_Id = @Donor_Id";

            dict2.Add("Donate_No", Donate_No + 1);
            dict2.Add("Donate_Total", Donate_Total + int.Parse(tbxDonate_Accou.Text.Trim()));
            dict2.Add("Donor_Id", HFD_Uid.Value);
        }
        NpoDB.ExecuteSQLS(strSql2, dict2);
    }
    //20151105新增 奉獻動機＆收看管道調查
    public void Questionaire_Add()
    {
        DataTable dt = NpoDB.QueryGetTable("SELECT MAX(Donate_Id) as Donate_Id FROM Donate where Donor_Id='" + HFD_Uid.Value + "'");
        DataRow dr = dt.Rows[0];
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = "insert into  Donate_OnlineQuestion\n";
        strSql += "(Donate_Id, Donor_Id, DonateMotive1, DonateMotive2, DonateMotive3, DonateMotive4, DonateMotive5, \n";
        strSql += " WatchMode1, WatchMode2, WatchMode3, WatchMode4, WatchMode5, WatchMode6, WatchMode7, WatchMode8, WatchMode9, Device, ToGOODTV,Create_Date,Create_DateTime,DonateWay) values\n";
        strSql += "(@Donate_Id,@Donor_Id,@DonateMotive1,@DonateMotive2,@DonateMotive3,@DonateMotive4,@DonateMotive5, \n";
        strSql += "@WatchMode1,@WatchMode2,@WatchMode3,@WatchMode4,@WatchMode5,@WatchMode6,@WatchMode7,@WatchMode8,@WatchMode9,@Device,@ToGOODTV,@Create_Date,@Create_DateTime,@DonateWay)\n";

        dict.Add("Donate_Id", dr["Donate_Id"].ToString().Trim());
        dict.Add("Donor_Id", HFD_Uid.Value);
        //奉獻動機
        if (this.DonateMotive1.Checked)
        {
            dict.Add("DonateMotive1", "支持媒體宣教大平台，可廣傳福音");
        }
        else
        {
            dict.Add("DonateMotive1", "");
        }
        if (this.DonateMotive2.Checked)
        {
            dict.Add("DonateMotive2", "個人靈命得造就");
        }
        else
        {
            dict.Add("DonateMotive2", "");
        }
        if (this.DonateMotive3.Checked)
        {
            dict.Add("DonateMotive3", "支持優質節目製作");
        }
        else
        {
            dict.Add("DonateMotive3", "");
        }
        if (this.DonateMotive4.Checked)
        {
            dict.Add("DonateMotive4", "支持GOOD TV家庭事工");
        }
        else
        {
            dict.Add("DonateMotive4", "");
        }
        if (this.DonateMotive5.Checked)
        {
            dict.Add("DonateMotive5", "感恩奉獻");
        }
        else
        {
            dict.Add("DonateMotive5", "");
        }
        //收看管道
        if (this.WatchMode1.Checked)
        {
            dict.Add("WatchMode1", "GOOD TV電視頻道");
        }
        else
        {
            dict.Add("WatchMode1", "");
        }
        if (this.WatchMode2.Checked)
        {
            dict.Add("WatchMode2", "官網");
        }
        else
        {
            dict.Add("WatchMode2", "");
        }
        if (this.WatchMode3.Checked)
        {
            dict.Add("WatchMode3", "Facebook");
        }
        else
        {
            dict.Add("WatchMode3", "");
        }
        if (this.WatchMode4.Checked)
        {
            dict.Add("WatchMode4", "Youtube");
        }
        else
        {
            dict.Add("WatchMode4", "");
        }
        if (this.WatchMode5.Checked)
        {
            dict.Add("WatchMode5", "好消息月刊");
        }
        else
        {
            dict.Add("WatchMode5", "");
        }
        if (this.WatchMode6.Checked)
        {
            dict.Add("WatchMode6", "GOOD TV簡介刊物");
        }
        else
        {
            dict.Add("WatchMode6", "");
        }
        if (this.WatchMode7.Checked)
        {
            dict.Add("WatchMode7", "教會牧者");
        }
        else
        {
            dict.Add("WatchMode7", "");
        }
        if (this.WatchMode8.Checked)
        {
            dict.Add("WatchMode8", "親友");
        }
        else
        {
            dict.Add("WatchMode8", "");
        }
        if (this.WatchMode9.Checked)
        {
            dict.Add("WatchMode9", "報章雜誌");
        }
        else
        {
            dict.Add("WatchMode9", "");
        }
        dict.Add("Device", "一般網頁");
        dict.Add("ToGOODTV", txtToGoodTV.Text.Trim().ToString());
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        dict.Add("DonateWay", "非線上單筆奉獻");

        NpoDB.ExecuteSQLS(strSql, dict);
    }
}
