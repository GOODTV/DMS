using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class DonateMgr_Pledge_Add : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Donor_Id");
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
        ddlDonate_Payment.SelectedIndex = 0;

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
        ddlDonate_Period.SelectedIndex = 0;

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
        strSql = @" select * , (Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.Invoice_ZipCode+B.mValue+Invoice_Address Else A.mValue+DONOR.Invoice_ZipCode+Invoice_Address End End) as [地址]
                    from Donor 
                        Left Join CODECITY As A On Donor.Invoice_City=A.mCode Left Join CODECITY As B On Donor.Invoice_Area=B.mCode 
                    where Donor_id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("PledgeInfo.aspx");

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
        Session["IDNo"] = dr["IDNo"].ToString().Trim();
        //最近捐款日：
        if (dr["Last_DonateDate"].ToString().Trim() != "")
        {
            tbxLast_DonateDate.Text = DateTime.Parse(dr["Last_DonateDate"].ToString().Trim()).ToShortDateString().ToString();
        }
        //收據抬頭
        tbxInvoice_Title.Text = dr["Invoice_Title"].ToString().Trim();

        //授權期間起
        tbxDonate_FromDate.Text = DateTime.Now.ToString("yyyy/MM") + "/01";   
        //下次扣款日期
        tbxNext_DonateDate.Text = DateTime.Now.ToString("yyyy/MM") + "/01";
        //20140813下方的收據開立與捐款人收據開立一樣
        ddlInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();
    }
    protected void ddlDonate_Payment_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(一般)" || ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(聯信)")
        {
            PanelAccount.Visible = false;
            PanelCreditCard.Visible = true;
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
            tbxAccoun_Bank.Text = "臺灣銀行";
        }
        if (ddlDonate_Payment.SelectedItem.Text == "郵局帳戶轉帳授權書")
        {
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = true;
            PanelACH.Visible = false;
            tbxPost_Name.Text = "";
            tbxPost_IDNo.Text = "";
            tbxPost_SavingsNo.Text = "";
            tbxPost_AccountNo.Text = "";
            tbxAccoun_Bank.Text = "郵局";
        }
        if (ddlDonate_Payment.SelectedItem.Text == "ACH轉帳授權書")
        {
            PanelCreditCard.Visible = false;
            PanelAccount.Visible = false;
            PanelACH.Visible = true;
            tbxP_BANK.Text = "";
            tbxP_RCLNO.Text = "";
            tbxP_PID.Text = Session["IDNo"].ToString().Trim();
            tbxAccoun_Bank.Text = "臺灣銀行";
        }
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            Pledge_AddNew();
            if (this.DonateMotive1.Checked || this.DonateMotive2.Checked || this.DonateMotive3.Checked || this.DonateMotive4.Checked || this.DonateMotive5.Checked || this.WatchMode1.Checked || this.WatchMode2.Checked || this.WatchMode3.Checked || this.WatchMode4.Checked || this.WatchMode5.Checked || this.WatchMode6.Checked || this.WatchMode7.Checked || this.WatchMode8.Checked || this.WatchMode9.Checked || txtToGoodTV.Text != "")
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
            //Donor 資料表中的授權日期更新
            Donor_Edit();

            //******設定收據編號******//
            string strSql2 = @"select Pledge_Id from Pledge
                where Donor_Id ='" + HFD_Uid.Value + "' and Create_DateTime='" + HFD_Create_DateTime.Value + "'";
            //****執行查詢流水號語法****//
            DataTable dt2 = NpoDB.QueryGetTable(strSql2);
            DataRow dr2 = dt2.Rows[0];
            string Pledge_Id = dr2["Pledge_Id"].ToString().Trim();
            SetSysMsg("轉帳授權書新增成功！");

            if (Session["cType"] == "PledgeInfo")
            {
                Response.Redirect(Util.RedirectByTime("Pledge_Edit.aspx", "Pledge_Id=" + Pledge_Id));
            }
            else if (Session["cType"] == "PledgeDataList")
            {
                Response.Redirect(Util.RedirectByTime("PledgeDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
            }
            else
            {
                Response.Redirect(Util.RedirectByTime("PledgeInfo.aspx"));
            }
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        if (Session["cType"] == "PledgeInfo")
        {
            Response.Redirect(Util.RedirectByTime("PledgeInfo.aspx"));
        }
        else if (Session["cType"] == "PledgeDataList")
        {
            Response.Redirect(Util.RedirectByTime("PledgeDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
        }
        else
        {
            Response.Redirect(Util.RedirectByTime("PledgeInfo.aspx"));
        }
    }
    public void Pledge_AddNew()
    {
        string strSql = "insert into  Pledge\n";
        strSql += "( Pledge_Id, Donor_Id, Donate_Payment, Donate_Purpose, Donate_Purpose_Type, Donate_Type, Donate_Amt, Donate_FromDate,\n";
        strSql += " Donate_ToDate, Donate_Period, Next_DonateDate, Card_Bank, Card_Type, Account_No, Valid_Date, Card_Owner,\n";
        strSql += " Owner_IDNo, Relation, Authorize, Post_Name, Post_IDNo, Post_SavingsNo, Post_AccountNo, Dept_Id, Invoice_Type,\n";
        strSql += " Invoice_Title, Accoun_Bank, Accounting_Title, Act_id, Status, Comment, Create_Date, Create_DateTime,\n";
        strSql += " Create_User, Create_IP, P_BANK, P_RCLNO, P_PID) values\n";
        strSql += "( (SELECT MAX(Pledge_Id) FROM Pledge)+1,@Donor_Id,@Donate_Payment,@Donate_Purpose,@Donate_Purpose_Type,@Donate_Type,@Donate_Amt,@Donate_FromDate,\n";
        strSql += "@Donate_ToDate,@Donate_Period,@Next_DonateDate,@Card_Bank,@Card_Type,@Account_No,@Valid_Date,@Card_Owner,\n";
        strSql += "@Owner_IDNo,@Relation,@Authorize,@Post_Name,@Post_IDNo,@Post_SavingsNo,@Post_AccountNo,@Dept_Id,@Invoice_Type,\n";
        strSql += "@Invoice_Title,@Accoun_Bank,@Accounting_Title,@Act_id,@Status,@Comment, @Create_Date,@Create_DateTime,\n";
        strSql += "@Create_User,@Create_IP,@P_BANK,@P_RCLNO,@P_PID) ";
        strSql += "\n";
        strSql += "SELECT MAX(Pledge_Id) FROM Pledge";


        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", HFD_Uid.Value);
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
        if (ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(一般)" || ddlDonate_Payment.SelectedItem.Text == "信用卡授權書(聯信)")
        {
            //if (tbxAccount_No1.Text.Length == 4 && tbxAccount_No2.Text.Length == 4 && tbxAccount_No3.Text.Length == 4 && tbxAccount_No4.Text.Length == 4)
            //{
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
            /*}
            else
            {
                ShowSysMsg("信用卡卡號不符合格式(各4碼一共16碼)");
                return;
            }*/
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
        dict.Add("Status", "授權中");
        dict.Add("Comment", tbxComment.Text.Trim());
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        string Create_DateTime = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        NpoDB.ExecuteSQLS(strSql, dict);

        HFD_Create_DateTime.Value = Create_DateTime;
    }
    public void Donor_Edit()
    {
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
        strSql2 = " update Donor set ";
        strSql2 += "  Last_PledgeDate = @Last_PledgeDate";
        strSql2 += " where Donor_Id = @Donor_Id";

        dict2.Add("Last_PledgeDate", DateTime.Now.ToString("yyyy-MM-dd"));
        dict2.Add("Donor_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql2, dict2);
    }
    //20140813 增加授權期間重複查檢 by 詩儀
    protected void tbxDonate_ToDate_Changed(object sender, EventArgs e)
    {
        string strSql = "";

        DataTable dt = null;
        Dictionary<string, object> dict = new Dictionary<string, object>();
        strSql = @"SELECT CONVERT(VARCHAR(12),MIN(Donate_FromDate),111) Donate_FromDate 
                     ,CONVERT(VARCHAR(12),MAX(Donate_ToDate),111) Donate_ToDate 
                     FROM dbo.PLEDGE 
                     WHERE Status='授權中' AND Donor_Id='"+ HFD_Uid.Value +"'";
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        String Donate_FromDate = dr["Donate_FromDate"].ToString();
        String Donate_ToDate = dr["Donate_ToDate"].ToString();
        if (Donate_FromDate != "" && Donate_ToDate != "")
        {
            if (DateTime.Parse(Donate_FromDate) <= DateTime.Parse(tbxDonate_FromDate.Text.Trim()) && DateTime.Parse(Donate_ToDate) >= DateTime.Parse(tbxDonate_FromDate.Text.Trim()))
            {
                ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('授權期間重覆！');</script>");
                return;
            }
        }
    }
    //20151105新增 奉獻動機＆收看管道調查
    public void Questionaire_Add()
    {
        DataTable dt = NpoDB.QueryGetTable("SELECT MAX(Pledge_Id) as Pledge_Id FROM Pledge where Donor_Id='" + HFD_Uid.Value + "'");
        DataRow dr = dt.Rows[0];
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = "insert into  Donate_OnlineQuestion\n";
        strSql += "(Pledge_Id, Donor_Id, DonateMotive1, DonateMotive2, DonateMotive3, DonateMotive4, DonateMotive5, \n";
        strSql += " WatchMode1, WatchMode2, WatchMode3, WatchMode4, WatchMode5, WatchMode6, WatchMode7, WatchMode8, WatchMode9, Device, ToGOODTV,Create_Date,Create_DateTime,DonateWay) values\n";
        strSql += "(@Pledge_Id,@Donor_Id,@DonateMotive1,@DonateMotive2,@DonateMotive3,@DonateMotive4,@DonateMotive5, \n";
        strSql += "@WatchMode1,@WatchMode2,@WatchMode3,@WatchMode4,@WatchMode5,@WatchMode6,@WatchMode7,@WatchMode8,@WatchMode9,@Device,@ToGOODTV,@Create_Date,@Create_DateTime,@DonateWay)\n";

        dict.Add("Pledge_Id", dr["Pledge_Id"].ToString().Trim());
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
        dict.Add("DonateWay", "紙本定期定額");

        NpoDB.ExecuteSQLS(strSql, dict);
    }
}