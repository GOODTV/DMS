using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateCreditCard : System.Web.UI.Page
{
    DataTable dtOnce = new DataTable();
    DataTable dtPeriod = new DataTable();
    string strRemoteAddr = "";
    string strDonateAmt;
    string strDonatePeriod;

    protected void Page_Load(object sender, EventArgs e)
    {
        strRemoteAddr = Request.ServerVariables["REMOTE_ADDR"];
        if (strRemoteAddr == "::1")
        {
            strRemoteAddr = "127.0.0.1";
        }

        if (Session["ItemOnce"] != null)
        {
            dtOnce = (DataTable)Session["ItemOnce"];
        }
        if (Session["ItemPeriod"] != null)
        {
            dtPeriod = (DataTable)Session["ItemPeriod"];
        }

        if (!IsPostBack)
        {

            Session["InsertPeriod"] = "N";

            LoadFormData();
            if (Session["DonorName"] != null)
            {
                lblTitle.Text = Session["DonorName"].ToString() + lblTitle.Text;
            }
            LoadDropDownList();

        }

    }
    //-------------------------------------------------------------------------------------------------------------
    private void LoadDropDownList()
    {
        /*ddlCardType.Items.Clear();
        ddlCardType.Items.Add(new ListItem("", ""));
        ddlCardType.Items.Add(new ListItem("VISA", "VISA"));
        ddlCardType.Items.Add(new ListItem("MASTER", "MASTER"));
        ddlCardType.Items.Add(new ListItem("JCB", "JCB"));*/

        ddlValidMonth.Items.Add(new ListItem("", ""));
        for (int i = 1; i < 13; i++)
        {
            ddlValidMonth.Items.Add(new ListItem(i.ToString("00"), i.ToString("00")));
        }

        ddlValidYear.Items.Add(new ListItem("", ""));
        for (int i = 0; i < 16; i++)
        {
            ddlValidYear.Items.Add(new ListItem((DateTime.Now.Year + i - 2000).ToString(), (DateTime.Now.Year + i - 2000).ToString()));
        }
    }
    //---------------------------------------------------------------------------
    private void LoadFormData()
    {
        List<ControlData> list = new List<ControlData>();
        list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_CardType", "rdoCardType1", "VISA", "VISA", false, ""));
        list.Add(new ControlData("Radio", "RDO_CardType", "rdoCardType2", "MASTER", "MASTER", false, ""));
        list.Add(new ControlData("Radio", "RDO_CardType", "rdoCardType3", "JCB", "JCB", false, ""));
        lblCardType.Text = HtmlUtil.RenderControl(list);
        lblCardType.Font.Size = 10;
    }
    //---------------------------------------------------------------------------
    private void InsertPledgeTemp()
    {
        for (int i = 0; i < dtPeriod.Rows.Count; i++)
        {
            //string strPayment = dtPeriod.Rows[i]["奉獻金額"].ToString();
            Dictionary<string, object> dict = new Dictionary<string, object>();
            List<ColumnData> list = new List<ColumnData>();
            SetColumnData(list, dtPeriod.Rows[i]);
            string strSql = Util.CreateInsertCommand("PLEDGE_Temp", list, dict);
            NpoDB.ExecuteSQLS(strSql, dict);
            Session["Msg"] = "Success!";
            Session["InsertPeriod"] = "Y";
        }
    }
    //-------------------------------------------------------------------------------------------------------------
    private void SetColumnData(List<ColumnData> list, DataRow dr)
    {
        string strSql = @"select * from Donor where Donor_Id=@Donor_Id and DeleteDate is null";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", Session["DonorID"].ToString());

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        string strInvoiceType = "";
        string strInvoiceTitle = "";
        if (dt.Rows.Count > 0)
        {
            DataRow dr1 = dt.Rows[0];
            strInvoiceType = dr1["Invoice_Type"].ToString();
            strInvoiceTitle = dr1["Invoice_Title"].ToString();
        }
        //Donor_Id	int	Checked 捐款人編號
        list.Add(new ColumnData("Donor_Id", Session["DonorID"].ToString(), true, false, true));
        //Donate_Payment	nvarchar(20)	Checked
        list.Add(new ColumnData("Donate_Payment", "信用卡授權書(一般)", true, false, false));
        //Donate_Purpose	nvarchar(20)	Checked 捐款用途
        //20140924線上奉獻項目皆改為「經常費」
        //list.Add(new ColumnData("Donate_Purpose", dr["奉獻項目"].ToString(), true, false, false));
        list.Add(new ColumnData("Donate_Purpose", "經常費", true, false, false));
        //Donate_Purpose_Type	nvarchar(10)	Checked 捐款用途ID
        list.Add(new ColumnData("Donate_Purpose_Type", "D", true, false, false));
        //Donate_Type	nvarchar(8)	Checked  D:Donate  
        list.Add(new ColumnData("Donate_Type", "長期捐款", true, false, false));
        //Invoice_Type 20140724新增[收據開立]
        list.Add(new ColumnData("Invoice_Type", strInvoiceType, true, false, false));
        //Invoice_Title 20140724新增[收據抬頭]
        list.Add(new ColumnData("Invoice_Title", strInvoiceTitle, true, false, false));
        //Donate_Amt	money	Checked
        list.Add(new ColumnData("Donate_Amt", dr["奉獻金額"].ToString(), true, false, false));
        //Donate_FirstAmt	money	Checked
        list.Add(new ColumnData("Donate_FirstAmt", "0", true, false, false));

        string strFromDate = "";
        string strToDate = "";
        string strBeginYear = "";
        string strBeginMonth = "";
        string strEndYear = "";
        string strEndMonth= "";
        string period = "";
        switch (dr["繳費方式"].ToString())
        {
            case "Monthly":
                strFromDate = dr["開始年"].ToString() + "/" + dr["開始月"].ToString() + "/01";
                strToDate = dr["結束年"].ToString() + "/" + dr["結束月"].ToString() + "/01";
                strBeginYear = dr["開始年"].ToString();
                strBeginMonth = dr["開始月"].ToString();
                strEndYear = dr["結束年"].ToString();
                strEndMonth = dr["結束月"].ToString();
                period = "月繳";
                break;
            case "Quarterly":
                //int iFrom= (Convert.ToInt16(dr["開始月"].ToString()) - 1) * 3 + 1;
                //int iTo = Convert.ToInt16(dr["開始月"].ToString()) * 3;
                //strFromDate = dr["開始年"].ToString() + "/" + iFrom.ToString("00") + "/01";
                //strToDate = dr["結束年"].ToString() + "/" + iTo.ToString("00") + "/01";
                strFromDate = dr["開始年"].ToString() + "/" + dr["開始月"].ToString() + "/01";
                strToDate = dr["結束年"].ToString() + "/" + dr["結束月"].ToString() + "/01";
                strBeginYear = dr["開始年"].ToString();
                strBeginMonth = dr["開始月"].ToString();
                strEndYear = dr["結束年"].ToString();
                strEndMonth = dr["結束月"].ToString();
                period = "季繳";
                break;
            case "Annual":
                strFromDate = dr["開始年"].ToString() + "/01/01";
                strToDate = dr["結束年"].ToString() + "/12/01";
                strBeginYear = dr["開始年"].ToString();
                strBeginMonth = "1";
                strEndYear = dr["結束年"].ToString();
                strEndMonth = "12";
                period = "年繳";
                break;
        }
        
        //Donate_FromDate	datetime	Checked
        list.Add(new ColumnData("Donate_FromDate", strFromDate, true, false, false));
        //Donate_ToDate	datetime	Checked
        list.Add(new ColumnData("Donate_ToDate", strToDate, true, false, false));
        //Donate_Period	nvarchar(10)	Checked
        //list.Add(new ColumnData("Donate_Period", dr["繳費方式"].ToString(), true, false, false));
        list.Add(new ColumnData("Donate_Period", period, true, false, false));
        //Next_DonateDate	datetime	Checked
        list.Add(new ColumnData("Next_DonateDate", strFromDate, true, false, false));
        //Card_Bank	nvarchar(20)	Checked
        list.Add(new ColumnData("Card_Bank", txtCardBank.Text.Trim(), true, false, false));
        //Card_Type	nvarchar(20)	Checked
        list.Add(new ColumnData("Card_Type", Util.GetControlValue("RDO_CardType"), true, false, false));
        //Account_No	nvarchar(20)	Checked
        list.Add(new ColumnData("Account_No", tbxAccount_No1.Text.Trim() + tbxAccount_No2.Text.Trim() + tbxAccount_No3.Text.Trim() + tbxAccount_No4.Text.Trim(), true, false, false));
        //Accoun_Bank
        list.Add(new ColumnData("Accoun_Bank", "臺灣銀行", true, false, false));
        //Valid_Date	nvarchar(4)	Checked
        list.Add(new ColumnData("Valid_Date", ddlValidMonth.SelectedValue + ddlValidYear.SelectedValue, true, false, false));
        //Card_Owner	nvarchar(20)	Checked
        list.Add(new ColumnData("Card_Owner", txtCardOwner.Text.Trim(), true, false, false));
        //Owner_IDNo	nvarchar(10)	Checked
        //20140613 移除「與持卡人關係」、「持卡人身份證字號」
        //list.Add(new ColumnData("Owner_IDNo", txtOwnerIdNo.Text.Trim(), true, false, false));
        //Relation	nvarchar(10)	Checked
        //list.Add(new ColumnData("Relation", txtRelation.Text.Trim(), true, false, false));
        //Authorize	nvarchar(10)	Checked        
        list.Add(new ColumnData("Authorize", txtAuthorize.Text.Trim(), true, false, false));
        //Create_Date	datetime	Checked
        list.Add(new ColumnData("Create_Date", Util.GetDBDateTime(), true, false, false));
        //Create_DateTime	datetime	Checked
        list.Add(new ColumnData("Create_DateTime", Util.GetDBDateTime(), true, false, false));
        //Create_User	nvarchar(20)	Checked
        list.Add(new ColumnData("Create_User", "線上金流", true, false, false));
        //Create_IP	nvarchar(16)	Checked
        list.Add(new ColumnData("Create_IP", strRemoteAddr, true, false, false));
        //Pledge_BeginYear	int	Checked
        list.Add(new ColumnData("Pledge_BeginYear", strBeginYear, true, false, false));
        //Pledge_BeginMonth	int	Checked
        list.Add(new ColumnData("Pledge_BeginMonth", strBeginMonth, true, false, false));
        //Pledge_EndYear	int	Checked
        list.Add(new ColumnData("Pledge_EndYear", strEndYear, true, false, false));
        //Pledge_EndMonth	int	Checked
        list.Add(new ColumnData("Pledge_EndMonth", strEndMonth, true, false, false));
        
        //Dept_Id  
        list.Add(new ColumnData("Dept_Id", "C001", true, false, false));
        //Status
        list.Add(new ColumnData("Status", "授權中", true, false, false));

        strDonateAmt = dr["奉獻金額"].ToString();
        strDonatePeriod = dr["繳費方式"].ToString();

    }
    //---------------------------------------------------------------------------
    protected void btnConfirm_Click(object sender, EventArgs e)
    {
        if (dtOnce.Rows.Count == 0 && dtPeriod.Rows.Count == 0)
        {
            Session["Msg"] = "無奉獻項目，請選擇奉獻項目！";
            Response.Redirect(Util.RedirectByTime("DonateOnlineAll.aspx"));            
        }
        InsertPledgeTemp();
        Session["ItemPeriod"] = null;
        Session["CardOwner"] = txtCardOwner.Text;
        Session["Account_No"] = tbxAccount_No1.Text + tbxAccount_No2.Text + tbxAccount_No3.Text + tbxAccount_No4.Text;
        
        //發送mail*****************************************

        //string SmtpServer = System.Configuration.ConfigurationManager.AppSettings["MailServer"];
        string MailSubject = "";
        string MailBody = "";
        string MailTo = "";
        string MailFrom = "";
        string strBody = "";

        strBody = "Dear " + Session["DonorName"].ToString() + " <BR/><BR/>";
        strBody += "We have received the information of your recurring donation. Thank you.<BR/><BR/>";

        //20140616 移除只剩「持卡人姓名」、「信用卡卡號」
        strBody += "<table border='1' cellpadding='1' cellspacing='0'>";
        //strBody += "<tr><th align='right'>發卡銀行</th><td>" + txtCardBank.Text.Trim() + "</td>";
        //strBody += "<th align='right'>信用卡種類</th><td>" + ddlCardType.SelectedValue + "</td>";
        if (Session["DonatePeriod"] != null)//只寄送定期奉獻的資料
        {
            strBody += Session["DonatePeriod"].ToString();
        }
        strBody += "</table><BR>";
        strBody += "<span>Card holder: " + txtCardOwner.Text.Trim() + "<span><BR>";
        strBody += "<span>Last 4 Digit of Card: ************" + tbxAccount_No4.Text.Trim() + "</span><BR><BR>";
        //strBody += "<th align='right'>授權碼</th><td>" + txtAuthorize.Text.Trim() + "</td>";
        //strBody += "<tr><th align='right'>有效期限</th><td>" + ddlValidMonth.SelectedValue + "月" + ddlValidYear.SelectedValue + "年</td>";
        //strBody += "<th align='right'>與持卡人關係</th><td>" + txtRelation.Text.Trim() + "</td>";
        //strBody += "<th align='right'>持卡人身份證字號</th><td>" + ((txtOwnerIdNo.Text.Trim() == "") ? "" : txtOwnerIdNo.Text.Trim().Substring(0, 6)+"****") + "</td></table><BR>";

        strBody += "We appreciate your offering and support. May God bless you! <BR>";
        strBody += "Any question you may call us at <BR>Tel：02-8024-3911<BR>Fax：02-8024-3933<BR>";
        strBody += "Address: 6F., No.911 Zhongzheng Rd., Zhonghe Dist., New Taipei City 23544, Taiwan, R.O.C.<BR>";
        SendEMailObject MailObject = new SendEMailObject();
        //MailObject.SmtpServer = SmtpServer;
        MailSubject = " GOODTV - Donation Confirmation (Recurring)";
        MailBody = strBody;
        MailTo = (String)Session["Email"];
        MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];

        string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, strBody);

        if (MailObject.ErrorCode != 0)
        {
            this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        }


        //發送內部通知mail *****************************************
        if (Session["DonorID"] == null)
        {
            Session["DonorName"] = null;
            Util.ShowMsg("Timeout! Please sign in again.");
            Response.Redirect(Util.RedirectByTime("CheckOut.aspx"));
        }

        string strDonorID = Session["DonorID"].ToString();
        string strPhone;

        string strSql = @"select * from Donor where Donor_Id=@Donor_Id and DeleteDate is null";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", strDonorID);

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            DataRow dr = dt.Rows[0];

            if (!string.IsNullOrEmpty(dr["Tel_Office"].ToString()))
            {
                strPhone = "(" + dr["Tel_Office_Loc"].ToString() + ")" + dr["Tel_Office"].ToString() + " 分機: " + dr["Tel_Office_Ext"].ToString();
            }
            else
            {
                strPhone = dr["Tel_Home"].ToString();
            }

            string strAddress = dr["Address"].ToString();
            string strInvoiceType = dr["Invoice_Type"].ToString();
            string strIsAnonymous = dr["IsAnonymous"].ToString();
            string strCity = Util.GetCityName(dr["City"].ToString());
            string strArea = Util.GetAreaName(dr["Area"].ToString());
            string strIsAbroad = dr["IsAbroad"].ToString();
            string strOverseasCountry = dr["OverseasCountry"].ToString();
            string strOverseasAddress = dr["OverseasAddress"].ToString();

            strBody = Session["DonorName"].ToString() + " 奉獻資料如下：<BR/><BR/>";
            strBody += "您輸入授權用信用卡資料如下：<BR/><BR/>";
            strBody += "捐款方式：定期定額奉獻 " + strDonatePeriod + "<BR/>";
            strBody += "捐款金額：" + strDonateAmt + " 元<BR/>";
            strBody += "E-mail：" + dr["Email"].ToString() + "<BR/>";
            strBody += "電話：" + strPhone + "<BR/>";
            strBody += "收據地址：" +
                (strIsAbroad == "N" ? strCity + strArea + strAddress : strOverseasAddress + " " + strOverseasCountry) + "<BR/>";
            strBody += "收據寄發：" + strInvoiceType + "<BR/>";
            strBody += "徵信錄：" + (strIsAnonymous == "Y" ? "不刊登" : "刊登") + " <BR/>";

            MailObject = new SendEMailObject();
            //MailObject.SmtpServer = SmtpServer;
            MailSubject = " GOODTV線上奉獻_新奉獻通知(定期定額奉獻)";
            MailBody = strBody;
            MailTo = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];
            MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];

            result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, strBody);


            //********************************************

            if (dtOnce.Rows.Count > 0)
            {
                Response.Redirect(Util.RedirectByTime("DonateInfoConfirm.aspx"));
            }
            else
            {
                Response.Redirect(Util.RedirectByTime("DonateDone.aspx"));
            }
        }
    }
    //---------------------------------------------------------------------------
    protected void btnPrev_Click(object sender, EventArgs e)
    {
        //Response.Redirect("ShoppingCart.aspx?Mode=FromDefault");
        Response.Redirect(Util.RedirectByTime("DonateInfoConfirm.aspx"));
    }
}