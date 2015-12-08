using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;


public partial class DonateMgr_Pledge_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Pledge_Id");
            LoadDropDownListData();
            Form_DataBind();
        }
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //20140415 Modify by GoodTV Tanya:授權方式改從資料庫抓代碼清單
        //授權方式
        Util.FillDropDownList(ddlDonate_Payment, Util.GetDataTable("CaseCode", "GroupName", "授權方式", "", ""), "CaseName", "CaseName", false);
        ddlDonate_Payment.Items.Insert(0, new ListItem("", ""));        
        //ddlDonate_Payment.Items.Insert(1, new ListItem("信用卡授權書", "信用卡授權書"));
        //ddlDonate_Payment.Items.Insert(2, new ListItem("帳戶轉帳授權書", "帳戶轉帳授權書"));
        //ddlDonate_Payment.Items.Insert(3, new ListItem("ACH轉帳授權書", "ACH轉帳授權書"));
        //ddlDonate_Payment.SelectedIndex = 0;

        //指定用途
        Util.FillDropDownList(ddlDonate_Purpose, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseName", false);
        ddlDonate_Purpose.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Purpose.SelectedIndex = 0;

        //轉帳週期 
        ddlDonate_Period.Items.Insert(0, new ListItem("單筆", "單筆"));
        ddlDonate_Period.Items.Insert(1, new ListItem("月繳", "月繳"));
        ddlDonate_Period.Items.Insert(2, new ListItem("季繳", "季繳"));
        ddlDonate_Period.Items.Insert(3, new ListItem("半年繳", "半年繳"));
        ddlDonate_Period.Items.Insert(4, new ListItem("年繳", "年繳"));
        ddlDonate_Period.SelectedIndex = 1;

        //信用卡別  
        Util.FillDropDownList(ddlCard_Type, Util.GetDataTable("CaseCode", "GroupName", "信用卡別", "", ""), "CaseName", "CaseName", false);
        ddlCard_Type.Items.Insert(0, new ListItem("", ""));
        ddlCard_Type.SelectedIndex = 0;

        //有效月年
        Util.FillDropDownList(ddlMonth_Valid_Date, 1, 12, true, 0);
        Util.FillDropDownList(ddlYear_Valid_Date, Int32.Parse(DateTime.Now.Year.ToString().Substring(2, 2)), Int32.Parse(DateTime.Now.Year.ToString().Substring(2, 2)) + 15, true, 0);

        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.SelectedIndex = 0;

        //收據開立
        Util.FillDropDownList(ddlInvoice_Type, Util.GetDataTable("CaseCode", "GroupName", "收據開立", "", ""), "CaseName", "CaseName", false);
        ddlInvoice_Type.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Type.SelectedIndex = 2;

        //入帳銀行改為文字方塊格式 2014/04/17
        /*Util.FillDropDownList(ddlAccoun_Bank, Util.GetDataTable("CaseCode", "GroupName", "入帳銀行", "", ""), "CaseName", "CaseName", false);
        ddlAccoun_Bank.Items.Insert(0, new ListItem("", ""));
        ddlAccoun_Bank.SelectedIndex = 0;*/

        //會計科目
        Util.FillDropDownList(ddlAccounting_Title, Util.GetDataTable("CaseCode", "GroupName", "款項會計科目", "", ""), "CaseName", "CaseName", false);
        ddlAccounting_Title.Items.Insert(0, new ListItem("", ""));
        ddlAccounting_Title.SelectedIndex = 0;

        //募款活動
        Util.FillDropDownList(ddlAct_Id, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        ddlAct_Id.Items.Insert(0, new ListItem("", ""));
        ddlAct_Id.SelectedIndex = 0;

        //授權狀態
        ddlStatus.Items.Insert(0, new ListItem("授權中", "授權中"));
        ddlStatus.Items.Insert(1, new ListItem("停止", "停止"));
        ddlStatus.SelectedIndex = 0;
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
        strSql = @" select * , (Case When D.Invoice_City='' Then D.Invoice_Address Else Case When B.mValue<>C.mValue Then B.mValue+D.Invoice_ZipCode+C.mValue+D.Invoice_Address Else B.mValue+D.Invoice_ZipCode+D.Invoice_Address End End) as [地址]
                    from Pledge P 
                        Left Join Donor D on P.Donor_Id = D.Donor_Id 
                        Left Join Dept De on P.Dept_Id = De.DeptID 
                        Left Join Act A on P.Act_Id = A.Act_Id 
                        Left Join CODECITY As B On D.Invoice_City=B.mCode 
                        Left Join CODECITY As C On D.Invoice_Area=C.mCode
                    where Pledge_id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("PledgeInfo.aspx");

        DataRow dr = dt.Rows[0];

        HFD_Donor_Id.Value = dr["Donor_Id"].ToString().Trim();
        //授權編號
        tbxPledge_Id.Text = dr["Pledge_Id"].ToString().Trim();
        //捐款人
        tbxDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
        //捐款人編號
        tbxDonor_Id.Text = dr["Donor_Id"].ToString().Trim();
        //身份別
        tbxDonor_Type.Text = dr["Donor_Type"].ToString().Trim();
        //收據地址
        tbxAddress.Text = dr["地址"].ToString().Trim();
        //收據開立
        tbxInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //最近捐款日：
        if (dr["Last_DonateDate"].ToString().Trim() != "")
        {
            tbxLast_DonateDate.Text = DateTime.Parse(dr["Last_DonateDate"].ToString().Trim()).ToShortDateString().ToString();
        }
        //收據抬頭
        tbxInvoice_Title.Text = dr["Invoice_Title"].ToString().Trim();
        //授權方式
        ddlDonate_Payment.Text = dr["Donate_Payment"].ToString().Trim();
        //指定用途
        ddlDonate_Purpose.Text = dr["Donate_Purpose"].ToString().Trim();
        //每次扣款
        tbxDonate_Amt.Text = (Convert.ToInt64(dr["Donate_Amt"])).ToString();
        //授權期間
        tbxDonate_FromDate.Text = DateTime.Parse(dr["Donate_FromDate"].ToString()).ToString("yyyy/MM/dd");
        tbxDonate_ToDate.Text = DateTime.Parse(dr["Donate_ToDate"].ToString()).ToString("yyyy/MM/dd");
        //轉帳週期
        ddlDonate_Period.Text = dr["Donate_Period"].ToString().Trim();
        //下次扣款日
        tbxNext_DonateDate.Text = (dr["Next_DonateDate"].ToString() == "") ? "" : DateTime.Parse(dr["Next_DonateDate"].ToString()).ToString("yyyy/MM/dd");
        if (dr["Donate_Payment"].ToString().Trim() == "信用卡授權書(一般)" || dr["Donate_Payment"].ToString().Trim() == "信用卡授權書(聯信)")
        {
            //銀行別
            tbxCard_Bank.Text = dr["Card_Bank"].ToString().Trim();
            //信用卡別
            ddlCard_Type.Text = dr["Card_Type"].ToString().Trim();
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
        //機構
        ddlDept.Text = dr["DeptShortName"].ToString().Trim();
        //收據開立
        ddlInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();
        //收據抬頭
        tbxInvoice_Title.Text = dr["Invoice_Title"].ToString().Trim();
        //入帳銀行
        tbxAccoun_Bank.Text = dr["Accoun_Bank"].ToString().Trim();
        //會計科目
        ddlAccounting_Title.SelectedValue = dr["Accounting_Title"].ToString().Trim();
        //募款活動
        ddlAct_Id.SelectedValue = dr["Act_Id"].ToString().Trim();

        //授權狀態
        ddlStatus.Text = dr["Status"].ToString().Trim();
        //停止日期
        if (dr["Break_Date"].ToString() != "")
        {
            tbxBreak_Date.Text = DateTime.Parse(dr["Break_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //停止原因
        tbxBreak_Reason.Text = dr["Break_Reason"].ToString().Trim();
        //備註
        tbxComment.Text = dr["Comment"].ToString().Trim();
        //資料建檔人員
        tbxCreate_User.Text = dr["Create_User"].ToString().Trim();
        //資料建檔日期
        if (dr["Create_Date"].ToString() != "")
        {
            tbxCreate_Date.Text = DateTime.Parse(dr["Create_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //最後異動人員
        tbxLastUpdate_User.Text = dr["LastUpdate_User"].ToString().Trim();
        //最後異動日期
        if (dr["LastUpdate_Date"].ToString() != "")
        {
            tbxLastUpdate_Date.Text = DateTime.Parse(dr["LastUpdate_Date"].ToString()).ToString("yyyy/MM/dd");
        }
    }
    protected void ddlDonate_Payment_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(一般)" || ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(聯信)") 
        {
            PanelAccount.Visible = false;
            PanelCreditCard.Visible = true;
            PanelACH.Visible = false;
            if (ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(聯信)")
            {
                ddlCard_Type.SelectedIndex = 3;
            }
            tbxAccoun_Bank.Text = "臺灣銀行";
        }
        if (ddlDonate_Payment.SelectedItem.Text == "郵局帳戶轉帳授權書")
        {
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = true;
            PanelACH.Visible = false;
            tbxAccoun_Bank.Text = "郵局";
        }
        if (ddlDonate_Payment.SelectedItem.Text == "ACH轉帳授權書")
        {
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = false;
            PanelACH.Visible = true;
            tbxAccoun_Bank.Text = "臺灣銀行";
        }
    }
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            Pledge_Edit();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("轉帳授權書修改成功！");
            if (Session["cType"].ToString() == "PledgeInfo")
            {
                Response.Redirect(Util.RedirectByTime("Pledge_Edit.aspx", "Pledge_Id=" + HFD_Uid.Value));
            }
            else if (Session["cType"].ToString() == "PledgeDataList")
            {
                Response.Redirect(Util.RedirectByTime("PledgeDataList.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
            }
            else
            {
                Response.Redirect(Util.RedirectByTime("PledgeInfo.aspx"));
            }
        }
    }
    protected void btnDel_Click(object sender, EventArgs e)
    {
        string strSql = "delete from Pledge where Pledge_Id=@Pledge_Id";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Pledge_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);

        //******修改Donor的資料******//
        Dictionary<string, object> dict_Donor = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql_Donor = " update Donor set ";
        strSql_Donor += "  Donate_No = Donate_No -1";


        //找尋同個人是否有其他筆授權紀錄
        string strSql2 = @"select isnull(MAX(CONVERT(NVARCHAR, Create_Date, 111)),'') as Create_Date from Pledge ";
        strSql2 += " where Donor_Id='" + HFD_Donor_Id.Value + "'";
        DataTable dt = NpoDB.GetDataTableS(strSql2, null);
        DataRow dr = dt.Rows[0];
        //找不到其他筆
        if (dr["Create_Date"].ToString() == "")
        {
            strSql_Donor += "  ,Last_PledgeDate = ''";
        }
        else
        {
            strSql_Donor += "  ,Last_PledgeDate = '" + dr["Create_Date"].ToString() + "'";
        }
        strSql_Donor += " where Donor_Id = @Donor_Id";
        dict_Donor.Add("Donor_Id", HFD_Donor_Id.Value);
        NpoDB.ExecuteSQLS(strSql_Donor, dict_Donor);

        SetSysMsg("資料刪除成功！");

        if (Session["cType"].ToString() == "PledgeInfo")
        {
            Response.Redirect(Util.RedirectByTime("PledgeInfo.aspx"));
        }
        else if (Session["cType"].ToString() == "PledgeDataList")
        {
            Response.Redirect(Util.RedirectByTime("PledgeDataList.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
        }
        else
        {
            Response.Redirect(Util.RedirectByTime("PledgeInfo.aspx"));
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        if (Session["cType"].ToString() == "PledgeInfo")
        {
            Response.Redirect(Util.RedirectByTime("PledgeInfo.aspx"));
        }
        else if (Session["cType"].ToString() == "PledgeDataList")
        {
            Response.Redirect(Util.RedirectByTime("PledgeDataList.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
        }
        else
        {
            Response.Redirect(Util.RedirectByTime("PledgeInfo.aspx"));
        }
    }
    public void Pledge_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update Pledge set ";

        strSql += "  Donate_Payment = @Donate_Payment";
        strSql += ", Donate_Purpose = @Donate_Purpose";
        strSql += ", Donate_Purpose_Type = @Donate_Purpose_Type";
        strSql += ", Donate_Type = @Donate_Type";
        strSql += ", Donate_Amt = @Donate_Amt";
        strSql += ", Donate_FromDate = @Donate_FromDate";
        strSql += ", Donate_ToDate = @Donate_ToDate";
        strSql += ", Donate_Period = @Donate_Period";
        strSql += ", Next_DonateDate = @Next_DonateDate";
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
        strSql += ", Dept_Id = @Dept_Id";
        strSql += ", Invoice_Type = @Invoice_Type";
        strSql += ", Invoice_Title = @Invoice_Title";
        strSql += ", Accoun_Bank = @Accoun_Bank";
        strSql += ", Accounting_Title = @Accounting_Title";
        strSql += ", Act_id = @Act_id";
        strSql += ", Status = @Status";
        strSql += ", Break_Date = @Break_Date";
        strSql += ", Break_Reason = @Break_Reason";
        strSql += ", Comment = @Comment";
        strSql += ", LastUpdate_Date = @LastUpdate_Date";
        strSql += ", LastUpdate_DateTime = @LastUpdate_DateTime";
        strSql += ", LastUpdate_User = @LastUpdate_User";
        strSql += ", LastUpdate_IP = @LastUpdate_IP";
        strSql += ", P_BANK = @P_BANK";
        strSql += ", P_RCLNO = @P_RCLNO";
        strSql += ", P_PID = @P_PID";
        strSql += " where Pledge_Id = @Pledge_Id";

        dict.Add("Donate_Payment", ddlDonate_Payment.SelectedItem.Text);
        dict.Add("Donate_Purpose", ddlDonate_Purpose.SelectedItem.Text);
        string[] Donate_Purpose_Type = { "", "A", "B", "C", "D", "E", "F" };
        for (int i = 0; i <= ddlDonate_Purpose.Items.Count; i++)
        {
            if (ddlDonate_Purpose.SelectedIndex == i)
            {
                dict.Add("Donate_Purpose_Type", Donate_Purpose_Type[i]);
            }
        }
        if (ddlDonate_Period.SelectedValue == "單筆")
        {
            dict.Add("Donate_Type", "單次捐款");
        }
        else
        {
            dict.Add("Donate_Type", "長期捐款");
        }
        dict.Add("Donate_Amt", tbxDonate_Amt.Text.Trim());
        dict.Add("Donate_FromDate", tbxDonate_FromDate.Text.Trim());
        dict.Add("Donate_ToDate", tbxDonate_ToDate.Text.Trim());
        dict.Add("Donate_Period", ddlDonate_Period.SelectedItem.Text);
        dict.Add("Next_DonateDate", tbxNext_DonateDate.Text.Trim());
        //20140509 修改 by Ian_Kao 
        if (ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(一般)" || ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(聯信)")
        {
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
        else if (ddlDonate_Payment.SelectedItem.Text == "ACH轉帳授權書")
        {
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
        dict.Add("Dept_Id", ddlDept.SelectedValue);
        dict.Add("Invoice_Type", ddlInvoice_Type.SelectedItem.Text);
        dict.Add("Invoice_Title", tbxInvoice_Title.Text.Trim());
        dict.Add("Accoun_Bank", tbxAccoun_Bank.Text.Trim());
        dict.Add("Accounting_Title", ddlAccounting_Title.SelectedItem.Text);
        if (ddlAct_Id.SelectedIndex != 0)
        {
            dict.Add("Act_Id", ddlAct_Id.SelectedValue);
        }
        else
        {
            dict.Add("Act_Id", "");
        }
        dict.Add("Status", ddlStatus.SelectedItem.Text);
        dict.Add("Comment", tbxComment.Text.Trim());
        if (tbxBreak_Date.Text.Trim() == "")
        {
            dict.Add("Break_Date", "null");
        }
        else
        {
            dict.Add("Break_Date", tbxBreak_Date.Text.Trim());
        }
        dict.Add("Break_Reason", tbxBreak_Reason.Text.Trim());
        dict.Add("LastUpdate_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("LastUpdate_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        dict.Add("LastUpdate_User", SessionInfo.UserName);
        dict.Add("LastUpdate_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());

        dict.Add("Pledge_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}