using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;

public partial class DonateMgr_Donate_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Donate_Id");
            LoadDropDownListData();
            Form_DataBind();
            //紀錄原始收據編號
            HFD_Invoice_No.Value = tbxInvoice_No.Text;
            //2014/1/21需求隱藏刪除鈕
            btnDel.Visible = false;
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

        //收據開立
        Util.FillDropDownList(ddlInvoice_Type, Util.GetDataTable("CaseCode", "GroupName", "收據開立", "", ""), "CaseName", "CaseName", false);
        ddlInvoice_Type.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Type.SelectedIndex = 0;

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

        //線上付款方式
        Util.FillDropDownList(ddlIEPayType, Util.GetDataTable("Donate_IePayType", "Display", "1", "", ""), "CodeName", "CodeID", false);
        ddlIEPayType.Items.Add(new ListItem("非信用卡", "99")); //開放可將非信用卡修改正確的付款方式
        ddlIEPayType.Items.Insert(0, new ListItem("", ""));
        ddlIEPayType.SelectedIndex = 0;

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
        strSql = @" select *  , (Case When dr.Invoice_City='' Then Invoice_Address Else Case 
                        When B.mValue<>C.mValue Then B.mValue+dr.Invoice_ZipCode+C.mValue+Invoice_Address 
                        Else B.mValue+dr.Invoice_ZipCode+Invoice_Address End End) as [收據地址]
                    from Donate do 
                        Left Join Donor dr on do.Donor_Id = dr.Donor_ID 
                        Left Join Dept D on do.Dept_Id = D.DeptID 
                        Left Join Act A on do.Act_Id = A.Act_Id  
                        Left Join CODECITY As B On dr.Invoice_City=B.mCode Left Join CODECITY As C On dr.Invoice_Area=C.mCode 
                        left join DONATE_IEPAY as I on do.od_sob = I.orderid
                    where dr.DeleteDate is null and Donate_Id='" + uid + "'";
        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("DonateInfo.aspx");

        DataRow dr = dt.Rows[0];

        //收據編號
        tbxInvoice_No_Top.Text = dr["Invoice_No"].ToString().Trim();
        //捐款人
        tbxDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
        //捐款人編號
        //Add by GoodTV Tanya:20140416
        tbxDonor_Id.Text = dr["Donor_Id"].ToString().Trim();
        //類別
        //Mark by GoodTV Tanya 20140416
        //tbxCategory.Text = dr["Category"].ToString().Trim();
        //身份別
        tbxDonor_Type.Text = dr["Donor_Type"].ToString().Trim();
        //收據地址
        tbxAddress.Text = dr["收據地址"].ToString();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //提款日期
        if (dr["Donate_Date"].ToString() != "")
        {
            tbxDonate_Date.Text = DateTime.Parse(dr["Donate_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //捐款方式
        ddlDonate_Payment.Text = dr["Donate_Payment"].ToString().Trim();
        //捐款用途
        ddlDonate_Purpose.Text = dr["Donate_Purpose"].ToString().Trim();
        //收據開立
        ddlInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();
        //線上付款方式
        ddlIEPayType.Text = dr["paytype"].ToString();
        HFD_IEPayType.Value = dr["paytype"].ToString();
        //線上訂單編號
        tbxOrderid.Text = dr["od_sob"].ToString();
        //捐款金額
        tbxDonate_Amt.Text = (Convert.ToInt64(dr["Donate_Amt"])).ToString();
        //手續費
        tbxDonate_Fee.Text = (Convert.ToInt64(dr["Donate_Fee"])).ToString();
        //實收金額
        tbxDonate_Accou.Text = (Convert.ToInt64(dr["Donate_Accou"])).ToString();
        //外幣幣別
        ddlDonate_Forign.Text = dr["Donate_Forign"].ToString();
        //外幣金額
        if (dr["Donate_ForignAmt"].ToString() != "")
        {
            tbxDonate_ForignAmt.Text = Convert.ToDouble(dr["Donate_ForignAmt"].ToString()).ToString();
        }
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
            ddlCard_Type.Text = dr["Card_Type"].ToString().Trim();
            //信用卡號
            if (dr["Account_No"].ToString() != "")
            {
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
            }
            //有效月年
            //20140513 增加判斷有效月年未填的狀態
            if (dr["Valid_Date"].ToString() == "")
            {
                ddlMonth_Valid_Date.Text = "";
                ddlYear_Valid_Date.Text = "";
            }
            else if (dr["Valid_Date"].ToString().Substring(0, 1).Trim() == "0")
            {
                ddlMonth_Valid_Date.Text = dr["Valid_Date"].ToString().Substring(1, 1).Trim();
                ddlYear_Valid_Date.Text = dr["Valid_Date"].ToString().Substring(2, 2).Trim();
            }
            else
            {
                ddlMonth_Valid_Date.Text = dr["Valid_Date"].ToString().Substring(0, 2).Trim();
                ddlYear_Valid_Date.Text = dr["Valid_Date"].ToString().Substring(2, 2).Trim();
            }
            ddlYear_Valid_Date.Text = dr["Valid_Date"].ToString().Substring(2, 2).Trim();
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
        else if (dr["Donate_Payment"].ToString().Trim() == "ACH轉帳授權書")
        {
            //收受行代號
            tbxP_BANK.Text = dr["P_BANK"].ToString().Trim();
            //收受者帳號
            tbxP_RCLNO.Text = dr["P_RCLNO"].ToString().Trim();
            //收受者身分證/統編
            tbxP_PID.Text = dr["P_PID"].ToString().Trim();
            PanelACH.Visible = true;
        }
        else if (dr["Donate_Payment"].ToString().Trim() == "網路信用卡")
        {
            PIEPayType.Visible = true;
            PIEPayOrderid.Visible = true;
            if (String.IsNullOrEmpty(tbxOrderid.Text))
            {
                ddlIEPayType.Enabled = false;
            }
        }

        //機構
        tbxDept.Text = dr["DeptShortName"].ToString().Trim();
        //收據抬頭
        tbxInvoice_Title.Text = dr["Invoice_Title"].ToString().Trim();
        //收據號碼
        //20140416 Modify by GoodTV Tanya:增加判斷「手開收據」是否勾選及「收據編號」是否唯讀 
        if (dr["Issue_Type"].ToString().Trim() == "M")
        {
            cbxInvoice_Pre.Checked = true;
            tbxInvoice_No.ReadOnly = false;
            tbxInvoice_No.Style.Value = "background-color:#ffffff";
        }
        else
        {
            cbxInvoice_Pre.Checked = false;
            tbxInvoice_No.ReadOnly = true;
            tbxInvoice_No.Style.Value = "background-color:#FFE1AF";
        }
        tbxInvoice_No.Text = dr["Invoice_No"].ToString().Trim();
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


        HFD_Donor_Id.Value = dr["Donor_Id"].ToString().Trim();
        //這次實收金額
        HFD_Donate_Accou.Value = tbxDonate_Accou.Text.Trim();
    }
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        bool check_close = Util.Get_Close("1", SessionInfo.DeptID, tbxDonate_Date.Text, SessionInfo.UserID);
        if (check_close == false)
        {
            bool flag = false;
            try
            {
                Donate_Edit();
                Donate_IEPay_Edit();
                flag = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (flag == true)
            {
                SetSysMsg("捐款資料修改成功！");
                // 2014/4/10 修改成儲存後頁面皆導到「捐款人基本資料維護-捐款紀錄」
                Response.Redirect(Util.RedirectByTime("DonateDataList.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
                /*
                if (Session["cType"] == "DonateInfo")
                {
                    Response.Redirect(Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=" + HFD_Uid.Value));
                }
                else if (Session["cType"] == "DonateDataList")
                {
                    Response.Redirect(Util.RedirectByTime("DonateDataList.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
                }
                else
                {
                    Response.Redirect(Util.RedirectByTime("DonateInfo.aspx"));
                }
                */
            }
        }
        else
        {
            AjaxShowMessage("您輸入的捐款日期『" + tbxDonate_Date.Text + "』 已關帳無法修改 ！");
        }
    }

    protected void btnDel_Click(object sender, EventArgs e)
    {
        bool check_close = Util.Get_Close("1", SessionInfo.DeptID, tbxDonate_Date.Text, SessionInfo.UserID);
        if (check_close == false)
        {
            string strSql = "delete from Donate where Donate_Id=@Donate_Id";
            Dictionary<string, object> dict = new Dictionary<string, object>();
            dict.Add("Donate_Id", HFD_Uid.Value);
            NpoDB.ExecuteSQLS(strSql, dict);
            //******修改Donor的資料******//
            Donor_Edit();

            SetSysMsg("捐款資料刪除成功！");
            if (Session["cType"].ToString() == "DonateInfo")
            {
                Response.Redirect(Util.RedirectByTime("DonateInfo.aspx"));
            }
            else if (Session["cType"].ToString() == "DonateDataList")
            {
                Response.Redirect(Util.RedirectByTime("DonateDataList.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
            }
            else
            {
                Response.Redirect(Util.RedirectByTime("DonateInfo.aspx"));
            }
        }
        else
        {
            AjaxShowMessage("您輸入的捐款日期『" + tbxDonate_Date.Text + "』 已關帳無法刪除 ！");
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=" + HFD_Uid.Value));
    }

    public void Donate_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update Donate set ";

        strSql += "  Donate_Date = @Donate_Date";
        strSql += ", Donate_Payment = @Donate_Payment";
        strSql += ", Donate_Purpose = @Donate_Purpose";
        strSql += ", Donate_Amt = @Donate_Amt";
        strSql += ", Donate_Fee = @Donate_Fee";
        strSql += ", Donate_Accou = @Donate_Accou";
        strSql += ", Donate_Forign = @Donate_Forign";
        strSql += ", Donate_ForignAmt = @Donate_ForignAmt";

        strSql += ", Check_No = @Check_No";
        strSql += ", Check_ExpireDate = @Check_ExpireDate";
        strSql += ", Card_Bank = @Card_Bank";
        strSql += ", Card_Type = @Card_Type";
        strSql += ", Account_No = @Account_No";
        strSql += ", Valid_Date = @Valid_Date";
        strSql += ", Card_Owner = @Card_Owner";
        strSql += ", Owner_IDNo = @Owner_IDNo";
        strSql += ", Relation = @Relation";
        strSql += ", Authorize = @Authorize";
        strSql += ", Post_Name = @Post_Name";
        strSql += ", Post_IDNo = @Post_IDNo";
        strSql += ", Post_SavingsNo = @Post_SavingsNo";
        strSql += ", Post_AccountNo = @Post_AccountNo";
        strSql += ", P_BANK = @P_BANK";
        strSql += ", P_RCLNO = @P_RCLNO";
        strSql += ", P_PID = @P_PID";

        strSql += ", Invoice_Title = @Invoice_Title";
        strSql += ", Invoice_Type = @Invoice_Type";
        strSql += ", Invoice_Pre = @Invoice_Pre";
        //20140509 新增by Ian_Kao 手開收據
        strSql += ", Issue_Type = @Issue_Type";
        strSql += ", Issue_Type_Keep = @Issue_Type_Keep";
        strSql += ", Invoice_No = @Invoice_No";
        strSql += ", Request_Date = @Request_Date";
        strSql += ", Accoun_Bank = @Accoun_Bank";
        strSql += ", Accoun_Date = @Accoun_Date";
        strSql += ", Donate_Type = @Donate_Type";
        strSql += ", Donation_NumberNo = @Donation_NumberNo";
        strSql += ", Donation_SubPoenaNo = @Donation_SubPoenaNo";
        strSql += ", Accounting_Title = @Accounting_Title";
        strSql += ", InvoceSend_Date= @InvoceSend_Date";
        strSql += ", Act_Id = @Act_Id";
        strSql += ", Comment = @Comment";
        strSql += ", Invoice_PrintComment= @Invoice_PrintComment";
        strSql += ", LastUpdate_Date = @LastUpdate_Date";
        strSql += ", LastUpdate_DateTime = @LastUpdate_DateTime";
        strSql += ", LastUpdate_User = @LastUpdate_User";
        strSql += ", LastUpdate_IP= @LastUpdate_IP";
        strSql += " where Donate_Id = @Donate_Id";

        dict.Add("Donate_Date", tbxDonate_Date.Text.Trim());
        dict.Add("Donate_Payment", ddlDonate_Payment.SelectedItem.Text);
        dict.Add("Donate_Purpose", ddlDonate_Purpose.SelectedItem.Text);
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
        //2014/5/15新增ACH轉帳授權欄位
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

            //收受行代號
            dict.Add("P_BANK", tbxP_BANK.Text.Trim());
            //收受者帳號
            dict.Add("P_RCLNO", tbxP_RCLNO.Text.Trim());
            //收受者身分證/統編
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
        dict.Add("Invoice_Title", tbxInvoice_Title.Text.Trim());
        if (cbxInvoice_Pre.Checked == false)
        {
            dict.Add("Issue_Type", "");
            dict.Add("Issue_Type_Keep", "");
            dict.Add("Invoice_Pre", "A");
            dict.Add("Invoice_No", tbxInvoice_No.Text.Trim());
        }
        if (cbxInvoice_Pre.Checked == true)
        {
            dict.Add("Issue_Type", "M");
            dict.Add("Issue_Type_Keep", "");
            dict.Add("Invoice_Pre", "");
            dict.Add("Invoice_No", tbxInvoice_No.Text.Trim());

            //dict.Add("Invoice_No", HFD_Invoice_No.Value);
        }
        dict.Add("Request_Date", tbxRequest_Date.Text.Trim());
        dict.Add("Accoun_Bank", ddlAccoun_Bank.SelectedItem.Text);
        dict.Add("Accoun_Date", tbxAccoun_Date.Text.Trim());
        dict.Add("Donate_Type", ddlDonate_Type.SelectedItem.Text);
        dict.Add("Invoice_Type", ddlInvoice_Type.SelectedItem.Text);
        dict.Add("Donation_NumberNo", tbxDonation_NumberNo.Text.Trim());
        dict.Add("Donation_SubPoenaNo", tbxDonation_SubPoenaNo.Text.Trim());
        dict.Add("Accounting_Title", ddlAccounting_Title.SelectedItem.Text);
        dict.Add("InvoceSend_Date", tbxInvoceSend_Date.Text.Trim());
        dict.Add("Act_Id", ddlAct_Id.SelectedValue);
        dict.Add("Comment", tbxComment.Text.Trim());
        dict.Add("Invoice_PrintComment", tbxInvoice_PrintComment.Text.Trim());
        dict.Add("LastUpdate_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("LastUpdate_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        dict.Add("LastUpdate_User", SessionInfo.UserName);
        dict.Add("LastUpdate_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());

        dict.Add("Donate_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);

        if (int.Parse(tbxDonate_Accou.Text.Trim()) > int.Parse(HFD_Donate_Accou.Value))//修改的捐款大於原本的捐款
        {
            //******修改Donor的資料******//
            Dictionary<string, object> dict_Donor = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Donor = " update Donor set ";
            int plus = int.Parse(tbxDonate_Accou.Text.Trim()) - int.Parse(HFD_Donate_Accou.Value);
            strSql_Donor += "  Donate_Total = Donate_Total +'" + plus + "'";
            strSql_Donor += " where Donor_Id = @Donor_Id";
            dict_Donor.Add("Donor_Id", HFD_Donor_Id.Value);
            NpoDB.ExecuteSQLS(strSql_Donor, dict_Donor);
        }
        else if (int.Parse(tbxDonate_Accou.Text.Trim()) < int.Parse(HFD_Donate_Accou.Value))//修改的捐款小於原本的捐款
        {
            //******修改Donor的資料******//
            Dictionary<string, object> dict_Donor = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Donor = " update Donor set ";
            int minus = int.Parse(HFD_Donate_Accou.Value) - int.Parse(tbxDonate_Accou.Text.Trim());
            strSql_Donor += "  Donate_Total = Donate_Total -'" + minus + "'";
            strSql_Donor += " where Donor_Id = @Donor_Id";
            dict_Donor.Add("Donor_Id", HFD_Donor_Id.Value);
            NpoDB.ExecuteSQLS(strSql_Donor, dict_Donor);
        }
        
    }

    public void Donor_Edit()
    {
        //******修改Donor的資料******//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Donor set ";
        strSql += "  Donate_No = Donate_No -1";

        //找尋同個人是否有其他筆捐款紀錄
        string strSql2 = @"select isnull(MAX(CONVERT(NVARCHAR, Create_Date, 111)),'') as Create_Date from Donate ";
        strSql2 += " where Donor_Id='" + HFD_Donor_Id.Value + "'";
        DataTable dt = NpoDB.GetDataTableS(strSql2, null);
        DataRow dr = dt.Rows[0];
        //找不到其他筆
        if (dr["Create_Date"].ToString() == "")
        {
            strSql += "  ,Begin_DonateDate = ''";
            strSql += "  ,Last_DonateDate = ''";
        }
        else
        {
            strSql += "  ,Last_DonateDate = '" + dr["Create_Date"].ToString() + "'";
        }
        strSql += "  ,Donate_Total = Donate_Total -'" + int.Parse(HFD_Donate_Accou.Value) + "'";
        strSql += " where Donor_Id = @Donor_Id";
        dict.Add("Donor_Id", HFD_Donor_Id.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }

    private void Donate_IEPay_Edit()
    {
        if (!String.IsNullOrEmpty(HFD_IEPayType.Value) && HFD_IEPayType.Value != ddlIEPayType.Text)
        {
            //****變數宣告****//
            Dictionary<string, object> dict = new Dictionary<string, object>();

            //****設定SQL指令****//
            string strSql = " update DONATE_IEPAY set ";
            strSql += "  paytype = @paytype ";
            strSql += " where orderid = @orderid";
            dict.Add("orderid", tbxOrderid.Text);
            dict.Add("paytype", ddlIEPayType.Text);
            NpoDB.ExecuteSQLS(strSql, dict);
        }

    }

    protected void ddlDonate_Payment_SelectedIndexChanged1(object sender, EventArgs e)
    {
        if (ddlDonate_Payment.SelectedItem.Text == "支票")
        {
            PanelCheck.Visible = true;
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = false;
            PanelACH.Visible = false;
            Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Check_ExpireDate();", true);
            PIEPayType.Visible = false;
            PIEPayOrderid.Visible = false;
        }
        else if (ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(一般)" || ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(聯信)")
        {
            PanelCheck.Visible = false;
            PanelCreditCard.Visible = true;
            PanelAccount.Visible = false;
            PanelACH.Visible = false;
            if (ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(聯信)")
            {
                ddlCard_Type.SelectedIndex = 3;
            }
            else
            {
                ddlCard_Type.SelectedIndex = 0;
            }
            PIEPayType.Visible = false;
            PIEPayOrderid.Visible = false;
        }
        else if (ddlDonate_Payment.SelectedItem.Text == "郵局帳戶轉帳授權書")
        {
            PanelCheck.Visible = false;
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = true;
            PanelACH.Visible = false;
            PIEPayType.Visible = false;
            PIEPayOrderid.Visible = false;
        }
        else if (ddlDonate_Payment.SelectedItem.Text == "ACH轉帳授權書")
        {
            PanelCheck.Visible = false;
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = false;
            PanelACH.Visible = true;
            PIEPayType.Visible = false;
            PIEPayOrderid.Visible = false;
        }
        else if (ddlDonate_Payment.SelectedItem.Text == "網路信用卡")
        {
            PanelCheck.Visible = false;
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = false;
            PanelACH.Visible = false;
            PIEPayType.Visible = true;
            PIEPayOrderid.Visible = true;
            if (String.IsNullOrEmpty(tbxOrderid.Text))
            {
                tbxOrderid.Enabled = false;
                ddlIEPayType.Enabled = false;
            }
        }
        else
        {
            PanelCheck.Visible = false;
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = false;
            PanelACH.Visible = false;
            PIEPayType.Visible = false;
            PIEPayOrderid.Visible = false;
        }
    }
}