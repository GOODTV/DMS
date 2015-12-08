using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;
using System.Web.UI.HtmlControls;

public partial class DonateMgr_Donate_Detail : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Donate_Id");
            //Form_DataBind();
        }
        Form_DataBind();
        //20140515 修改 by Ian_Kao 交換function執行順序
        Check_Close();
        Export();
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
        strSql = @" select *  , (Case When ISNULL(dr.Invoice_City,'')='' Then Invoice_Address Else Case 
                        When B.mValue<>C.mValue Then B.mValue+dr.Invoice_ZipCode+C.mValue+Invoice_Address 
                        Else B.mValue+dr.Invoice_ZipCode+Invoice_Address End End) as [收據地址]
                    from Donate do 
                        Left Join Donor dr on do.Donor_Id = dr.Donor_ID 
                        Left Join Dept D on do.Dept_Id = D.DeptID 
                        Left Join Act A on do.Act_Id = A.Act_Id  
                        Left Join CODECITY As B On dr.Invoice_City=B.mCode Left Join CODECITY As C On dr.Invoice_Area=C.mCode 
                        left join DONATE_IEPAY as I on do.od_sob = I.orderid
                        left join Donate_IePayType as di on I.paytype = di.CodeID
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
        //收據編號
        // 2014/4/10 修正收據編號
        //Mark by GoodTV Tanya:因下方已有收據編號
        //tbxInvoice_No_Top.Text = dr["Invoice_No"].ToString().Trim();
        //tbxInvoice_No_Top.Text = dr["Invoice_Pre"].ToString().Trim() + dr["Invoice_No"].ToString().Trim();
        //捐款人
        tbxDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
        //捐款人編號
        //Add by GoodTV Tanya 20140416        
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
            Session["Donate_Date"] = DateTime.Parse(dr["Donate_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //捐款方式
        tbxDonate_Payment.Text = dr["Donate_Payment"].ToString().Trim();
        //線上付款方式
        tbxIEPayType.Text = dr["CodeName"].ToString();
        //線上訂單編號
        tbxOrderid.Text = dr["od_sob"].ToString();
        //捐款用途
        tbxDonate_Purpose.Text = dr["Donate_Purpose"].ToString().Trim();
        //收據開立
        tbxInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();

        if (dr["Donate_Payment"].ToString().Trim() == "支票")
        {
            tbxCheck_No.Text = dr["Check_No"].ToString().Trim();
            if (dr["Check_ExpireDate"].ToString() != "")
            {
                tbxCheck_ExpireDate.Text = DateTime.Parse(dr["Check_ExpireDate"].ToString()).ToString("yyyy/MM/dd");
            }
            else
            {
                tbxCheck_ExpireDate.Text = dr["Check_ExpireDate"].ToString().Trim();
            }
            PanelCheck.Visible = true;
        }
        else if (dr["Donate_Payment"].ToString().Trim() == "信用卡授權書(一般)" || dr["Donate_Payment"].ToString().Trim() == "信用卡授權書(聯信)")
        {
            //銀行別
            tbxCard_Bank.Text = dr["Card_Bank"].ToString().Trim();
            //信用卡別
            tbxCard_Type.Text = dr["Card_Type"].ToString().Trim();
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
        else if (dr["Donate_Payment"].ToString().Trim() == "網路信用卡")
        {
            PIEPayType.Visible = true;
            PIEPayOrderid.Visible = true;
        }
        //實收金額
        HFD_Donate_Accou.Value = (Convert.ToInt64(dr["Donate_Accou"])).ToString();
        //機構
        tbxDept.Text = dr["DeptShortName"].ToString().Trim();
        //收據抬頭
        tbxInvoice_Title.Text = dr["Invoice_Title"].ToString().Trim();
        //收據號碼
        //Add by GoodTV Tanya:增加判斷是否勾選「手開收據」
        if (dr["Issue_Type"].ToString().Trim() == "M")
            cbxInvoice_Pre.Checked = true;
        tbxInvoice_No.Text = dr["Invoice_Pre"].ToString().Trim() + dr["Invoice_No"].ToString().Trim();
        HFD_Invoice_No.Value = dr["Invoice_No"].ToString().Trim();
        //請款日期
        if (dr["Request_Date"].ToString()!="" && DateTime.Parse(dr["Request_Date"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxRequest_Date.Text = DateTime.Parse(dr["Request_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //入帳銀行
        tbxAccoun_Bank.Text = dr["Accoun_Bank"].ToString().Trim();
        //沖帳日期
        if (dr["Accoun_Date"].ToString() != "" && DateTime.Parse(dr["Accoun_Date"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxAccoun_Date.Text = DateTime.Parse(dr["Accoun_Date"].ToString()).ToString("yyyy/MM/dd"); 
        }
        //捐款類別
        tbxDonate_Type.Text = dr["Donate_Type"].ToString().Trim();
        //傳票號碼
        tbxDonation_NumberNo.Text = dr["Donation_NumberNo"].ToString().Trim();
        //劃撥 / 匯款單號
        tbxDonation_SubPoenaNo.Text = dr["Donation_SubPoenaNo"].ToString().Trim();
        //會計科目
        tbxAccounting_Title.Text = dr["Accounting_Title"].ToString().Trim();
        //收據寄送
        if (dr["InvoceSend_Date"].ToString() != "" && DateTime.Parse(dr["InvoceSend_Date"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxInvoceSend_Date.Text  = DateTime.Parse(dr["InvoceSend_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //專案活動
        tbxAct_Id.Text = dr["Act_ShortName"].ToString().Trim();
        //捐款備註
        tbxComment.Text = dr["Comment"].ToString().Trim();
        //收據備註
        tbxInvoice_PrintComment.Text = dr["Invoice_PrintComment"].ToString().Trim();
        //首印日期
        if (dr["Invoice_Print_Date"].ToString() != "")
        {
            tbxInvoice_Print_Date.Text = DateTime.Parse(dr["Invoice_Print_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //首印經手人
        tbxInvoice_FirstPrint_User.Text = dr["Invoice_FirstPrint_User"].ToString().Trim();
        //最後補印日期
        if (dr["Invoice_RePrint_Date"].ToString() != "")
        {
            tbxInvoice_RePrint_Date.Text = DateTime.Parse(dr["Invoice_RePrint_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //最後補印經手人
        tbxInvoice_LastPrint_User.Text = dr["Invoice_LastPrint_User"].ToString().Trim();
        //年度證明
        //首印日期
        if (dr["Invoice_Yearly_Print_Date"].ToString() != "")
        {
            tbxInvoice_Yearly_Print_Date.Text = DateTime.Parse(dr["Invoice_Yearly_Print_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //首印經手人
        tbxInvoice_Yearly_FirstPrint_User.Text = dr["Invoice_Yearly_FirstPrint_User"].ToString().Trim();
        //最後補印日期
        if (dr["Invoice_Yearly_RePrint_Date"].ToString() != "")
        {
            tbxInvoice_Yearly_RePrint_Date.Text = DateTime.Parse(dr["Invoice_Yearly_RePrint_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //最後補印經手人
        tbxInvoice_Yearly_LastPrint_User.Text = dr["Invoice_Yearly_LastPrint_User"].ToString().Trim();
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
        // 2014/4/9 修改"收據列印"、"收據補印"按鈕的判斷
        if (tbxInvoice_Type.Text == "年度證明")
        {
            btnPrint.Visible = false;
            btnRePrint.Visible = false;
        }
        else
        {
            btnPrint.Visible = true;
            btnRePrint.Visible = true;
        }

    }
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("Donate_Edit.aspx", "Donate_Id=" + HFD_Uid.Value));
    }
    protected void btnExport_Click(object sender, EventArgs e)
    {
        bool check_close = Util.Get_Close("1", SessionInfo.DeptID, tbxDonate_Date.Text, SessionInfo.UserID);
        if (check_close == false)
        {
            bool flag = false;
            try
            {
                //******修改Donate的資料******//
                Donate_Edit();
                flag = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (flag == true)
            {
                //******修改Donor的資料******//
                Donor_Edit();
                SetSysMsg("收據已作廢！");
                Response.Redirect(Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=" + HFD_Uid.Value));
            }
        }
    }
    protected void btn_ReExport_Click(object sender, EventArgs e)
    {
        //本次捐款如果是重新取號而報廢的話,把新取號的單子報廢
        string strSql_Donate = "select Donate_Id from Donate where Invoice_No_Old = @Invoice_No_Old";
        Dictionary<string, object> dict_Donate = new Dictionary<string, object>();
        dict_Donate.Add("Invoice_No_Old", HFD_Invoice_No.Value);
        DataTable dt = NpoDB.GetDataTableS(strSql_Donate, dict_Donate);
        if (dt.Rows.Count != 0)
        {
            string Donate_Id = NpoDB.GetScalarS(strSql_Donate, dict_Donate);
            for (int i = 0; i < 100; i++)
            {
                strSql_Donate = "select * from Donate where Donate_Id = @Donate_Id";
                dict_Donate.Add("Donate_Id", Donate_Id);
                dt = NpoDB.GetDataTableS(strSql_Donate, dict_Donate);
                DataRow dr = dt.Rows[0];
                dict_Donate.Clear();//
                //20140509 修改 by Ian_Kao 手開收據
                //if (dr["Export"].ToString() == "N")
                if (dr["Issue_Type"].ToString() == "M")
                {
                    //找到尚未報廢且最新取號後跳出迴圈
                    break;
                }
                else
                {
                    //未找到尚未報廢且最新取號的
                    strSql_Donate = "select Donate_Id from Donate where Invoice_No_Old = @Invoice_No_Old";
                    dict_Donate.Add("Invoice_No_Old", dr["Invoice_No"].ToString());
                    Donate_Id = NpoDB.GetScalarS(strSql_Donate, dict_Donate);
                }
            }
            //20140509 修改 by Ian_Kao 手開收據
            ////把最新取號的Export修改成'Y'(錯誤邏輯)
            ////把最新取號的Issue_Type修改成'M'
            //****變數宣告****//
            Dictionary<string, object> dict = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql = " update Donate set ";
            strSql += " Issue_Type= @Issue_Type";
            //strSql += " Export= @Export";
            strSql += " where Donate_Id = @Donate_Id";
            //dict.Add("Export", "Y");
            dict.Add("Issue_Type", "M");
            dict.Add("Donate_Id", Donate_Id);
            NpoDB.ExecuteSQLS(strSql, dict);
        }

        //單純報廢單次捐款
        else
        {
            //******修改Donor的資料******//
            Dictionary<string, object> dict_Donor = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Donor = " update Donor set ";
            strSql_Donor += "  Donate_No = Donate_No + 1";

            //找尋同個人是否有其他筆捐款紀錄
            string strSql2 = @"select isnull(MAX(CONVERT(NVARCHAR, Create_Date, 111)),'') as Create_Date from Donate ";
            //20140514 修改 by Ian_Kao 判斷式應該是尋找有無其他"未作廢"的捐款紀錄,因此Issue_Type != 'D' 但是此情況不包含 Issue_Type is null,所以要特別搜搜尋此情況
            //strSql2 += " where Donor_Id='" + HFD_Donor_Id.Value + "' and ( Issue_Type != '' or  Issue_Type != 'M') ";//Export != 'N'";
            strSql2 += " where Donor_Id='" + HFD_Donor_Id.Value + "' and (Issue_Type is null or Issue_Type != 'D') ";//Export != 'N'";
            DataTable dt2 = NpoDB.GetDataTableS(strSql2, null);
            DataRow dr2 = dt2.Rows[0];
            //找不到其他筆
            if (dr2["Create_Date"].ToString() == "")
            {
                //strSql_Donor += "  ,Begin_DonateDate = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";
                strSql_Donor += "  ,Begin_DonateDate = '" + Session["Donate_Date"] + "'";
                //strSql_Donor += "  ,Last_DonateDate = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";
                strSql_Donor += "  ,Last_DonateDate = '" + Session["Donate_Date"] + "'";
            }
            else//有其他筆
            {
                //strSql_Donor += "  ,Last_DonateDate = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";
                strSql_Donor += "  ,Last_DonateDate = '" + Session["Donate_Date"] + "'";
            }
            strSql_Donor += "  ,Donate_Total = Donate_Total +'" + int.Parse(HFD_Donate_Accou.Value) + "'";
            strSql_Donor += " where Donor_Id = @Donor_Id";
            dict_Donor.Add("Donor_Id", HFD_Donor_Id.Value);
            NpoDB.ExecuteSQLS(strSql_Donor, dict_Donor);
        }

        //把本次Export修改成'N'(錯誤)
        //把本次Issue_Type修改成''
        //****變數宣告****//
        Dictionary<string, object> dict3 = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql3 = " update Donate set ";
        //strSql3 += " Export= @Export";
        strSql3 += " Issue_Type= @Issue_Type";
        strSql3 += " where Donate_Id = @Donate_Id";
        //dict3.Add("Export", "N");
        dict3.Add("Issue_Type", "");
        dict3.Add("Donate_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql3, dict3);

        SetSysMsg("收據已還原！");
        Response.Redirect(Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=" + HFD_Uid.Value));
    }
    protected void btn_Re_No_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("Donate_New_InvoiceNo.aspx", "Donate_Id=" + HFD_Uid.Value));
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
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
    protected void btn_Add_self_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("Donate_Add.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
    }
    protected void btn_Add_other_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("DonorQry.aspx"));
    }
    public void Export()
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;

        //****變數設定****//
        uid = HFD_Uid.Value;

        //****設定查詢****//
        strSql = " select Donate_Amt, Donate_Fee, Donate_Accou, Donate_Forign, Donate_ForignAmt, Issue_Type, Invoice_Print  from Donate where Donate_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);

        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("DonateInfo.aspx");

        DataRow dr = dt.Rows[0];
        if (dr["Issue_Type"].ToString() == "" || dr["Issue_Type"].ToString() == "M")
        //if (dr["Export"].ToString() == "N")
        {
            //捐款金額
            tbxDonate_Amt.Text = (Convert.ToInt64(dr["Donate_Amt"])).ToString();
            //手續費
            tbxDonate_Fee.Text = (Convert.ToInt64(dr["Donate_Fee"])).ToString();
            //實收金額
            tbxDonate_Accou.Text = (Convert.ToInt64(dr["Donate_Accou"])).ToString();

            //外幣幣別
            tbxDonate_Forign.Text = dr["Donate_Forign"].ToString();
            //外幣金額
            if (dr["Donate_ForignAmt"].ToString() != "")
            {
                tbxDonate_ForignAmt.Text = Convert.ToDouble(dr["Donate_ForignAmt"].ToString()).ToString();
            }
            //20140509 修改 by Ian_Kao 判斷是否是補印收據
            //Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Export_N();", true);
            //20140515 修改 by Ian_Kao 增加判斷Invoice_Print是0 = 未列印
            if (dr["Invoice_Print"].ToString() == "" || dr["Invoice_Print"].ToString() == "0")
            {
                Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Export_N(1);", true);
            }
            else if (dr["Invoice_Print"].ToString() == "1")
            {
                Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Export_N(2);", true);
            }
        }
        else if (dr["Issue_Type"].ToString() == "D")
        //if (dr["Export"].ToString() == "Y")
        {
            //捐款金額
            tbxDonate_Amt.Text = "0";
            //手續費
            tbxDonate_Fee.Text = "0";
            //實收金額
            tbxDonate_Accou.Text = "0";
            Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Export_Y();", true);
        }
    }
    public void Check_Close()
    {
        //20140515 修改 by Ian_Kao 修改判斷關帳的日期為系統當日而不是捐款日期、判斷是"捐款關帳日"並非"捐物資關帳日"
        //bool check_close = Util.Get_Close("2", SessionInfo.DeptID, tbxDonate_Date.Text, SessionInfo.UserID);
        //bool check_close = Util.Get_Close("1", SessionInfo.DeptID, DateTime.Now.ToString("yyyy/MM/dd"), SessionInfo.UserID);
        //20140516 修改判斷關帳的日期為「捐款日期」
        bool check_close = Util.Get_Close("1", SessionInfo.DeptID, tbxDonate_Date.Text, SessionInfo.UserID);
        if (check_close == true)
        {
            //20140515 新增 by Ian_Kao 判斷是列印收據還是補印收據
            //****變數宣告****//
            string strSql, uid;
            DataTable dt;

            //****變數設定****//
            uid = HFD_Uid.Value;

            //****設定查詢****//
            //20140515 修改 by Ian_Kao 增加Invoice_Type欄位判斷"收據開立"
            //strSql = " select Donate_Amt, Donate_Fee, Donate_Accou, Donate_Forign, Issue_Type, Invoice_Print  from Donate where Donate_Id='" + uid + "'";
            strSql = " select Invoice_Type, Donate_Amt, Donate_Fee, Donate_Accou, Donate_Forign, Issue_Type, Invoice_Print  from Donate where Donate_Id='" + uid + "'";

            //****執行語法****//
            dt = NpoDB.QueryGetTable(strSql);

            //資料異常
            if (dt.Rows.Count <= 0)
                //todo : add Default.aspx page
                Response.Redirect("DonateInfo.aspx");

            DataRow dr = dt.Rows[0];
            if (dr["Invoice_Print"].ToString() == "" || dr["Invoice_Print"].ToString() == "0")
            {
                Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Close(1);", true);
            }
            else if (dr["Invoice_Print"].ToString() == "1")
            {
                Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Close(2);", true);
            }
            //20140515 新增 by Ian_Kao 若超過關帳日期且「收據開立」為「年度證明」，「收據列印」的按鈕要disable。
            if (dr["Invoice_Type"].ToString() == "年度證明")
            {
                btnPrint.Visible = false;
                btnRePrint.Visible = false;
            }
        }
    }
    public void Donate_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Donate set ";
        //20140509 修改 by Ian_Kao 修改作廢收據的邏輯
        //strSql += " Export= @Export";
        strSql += " Issue_Type = @Issue_Type";
        strSql += " where Donate_Id = @Donate_Id";
        //dict.Add("Export", "Y");
        dict.Add("Issue_Type", "D");
        dict.Add("Donate_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void Donor_Edit()
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Donor set ";
        strSql += "  Donate_No = Donate_No -1";

        //找尋同個人是否有其他筆捐款紀錄
        string strSql2 = @"select isnull(MAX(CONVERT(NVARCHAR, Donate_Date, 111)),'') as Donate_Date from Donate ";
        strSql2 += " where Donor_Id='" + HFD_Donor_Id.Value + "' and Issue_Type != 'D'";  //and Export != 'Y'";
        DataTable dt2 = NpoDB.GetDataTableS(strSql2, null);
        DataRow dr2 = dt2.Rows[0];
        //找不到其他筆
        if (dr2["Donate_Date"].ToString() == "")
        {
            strSql += "  ,Begin_DonateDate = ''";
            strSql += "  ,Last_DonateDate = ''";
        }
        else//有其他筆
        {
            strSql += "  ,Last_DonateDate = '" + dr2["Donate_Date"].ToString() + "'";
        }
        strSql += "  ,Donate_Total = Donate_Total -'" + int.Parse(HFD_Donate_Accou.Value) + "'";
        strSql += " where Donor_Id = @Donor_Id";
        dict.Add("Donor_Id", HFD_Donor_Id.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}