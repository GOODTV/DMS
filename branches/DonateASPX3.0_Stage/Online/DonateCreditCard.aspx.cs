using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateCreditCard : System.Web.UI.Page
{

    string strEmail = "";

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {

            //POST接收值
            //捐款方式
            HFD_chkItem.Value = Request.Params["HFD_chkItem"];
            //奉獻金額
            HFD_Amount.Value = Request.Params["HFD_Amount"];
            //繳費方式
            HFD_PayType.Value = Request.Params["HFD_PayType"];
            //捐款人編號
            HFD_DonorID.Value = Request.Params["HFD_DonorID"];
            //捐款人名稱
            HFD_DonorName.Value = Request.Params["HFD_DonorName"];
            lblTitle.Text = Request.Params["HFD_DonorName"];

            LoadFormData();
            LoadDropDownList();
            
        }

    }

    //-------------------------------------------------------------------------------------------------------------
    private void LoadDropDownList()
    {

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
        //lblCardType.Font.Size = 10;
    }

    //---------------------------------------------------------------------------
    private void InsertPledgeTemp()
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();
        SetColumnData(list);
        string strSql = Util.CreateInsertCommand("PLEDGE_Temp", list, dict);
        NpoDB.ExecuteSQLS(strSql, dict);
    }

    //-------------------------------------------------------------------------------------------------------------
    private void SetColumnData(List<ColumnData> list)
    {
        string strSql = @"select * from Donor where Donor_Id=@Donor_Id and DeleteDate is null ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", HFD_DonorID.Value);

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        string strInvoiceType = "";
        string strInvoiceTitle = "";
        if (dt.Rows.Count > 0)
        {
            DataRow dr1 = dt.Rows[0];
            strInvoiceType = dr1["Invoice_Type"].ToString();
            strInvoiceTitle = dr1["Invoice_Title"].ToString();
            strEmail = dr1["Email"].ToString();
        }
        //Donor_Id	int	Checked 捐款人編號
        list.Add(new ColumnData("Donor_Id", HFD_DonorID.Value, true, false, true));
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
        list.Add(new ColumnData("Donate_Amt", HFD_Amount.Value, true, false, false));
        //Donate_FirstAmt	money	Checked
        list.Add(new ColumnData("Donate_FirstAmt", "0", true, false, false));

        string strFromDate = "";
        string strToDate = "2099/12/01";
        string strBeginYear = Util.GetDBDateTime().Year.ToString();
        string strBeginMonth = "";
        string strEndYear = "2099";
        string strEndMonth = "12";
        switch (HFD_PayType.Value)
        {
            case "月繳":
                strBeginMonth = Util.GetDBDateTime().Month.ToString();
                strFromDate = strBeginYear + "/" + strBeginMonth + "/01";
                break;
            case "季繳":
                strBeginMonth = Util.GetDBDateTime().Month.ToString();
                strFromDate = strBeginYear + "/" + strBeginMonth + "/01";
                break;
            case "年繳":
                strBeginMonth = "1";
                strFromDate = strBeginYear + "/01/01";
                break;
        }
        
        //Donate_FromDate	datetime	Checked
        list.Add(new ColumnData("Donate_FromDate", strFromDate, true, false, false));
        //Donate_ToDate	datetime	Checked
        list.Add(new ColumnData("Donate_ToDate", strToDate, true, false, false));
        //Donate_Period	nvarchar(10)	Checked
        list.Add(new ColumnData("Donate_Period", HFD_PayType.Value, true, false, false));
        //Next_DonateDate	datetime	Checked
        list.Add(new ColumnData("Next_DonateDate", strFromDate, true, false, false));
        //Card_Bank	nvarchar(20)	Checked
        list.Add(new ColumnData("Card_Bank", txtCardBank.Text.Trim(), true, false, false));
        //Card_Type	nvarchar(20)	Checked
        list.Add(new ColumnData("Card_Type", Util.GetControlValue("RDO_CardType"), true, false, false));
        //Account_No	nvarchar(20)	Checked
        list.Add(new ColumnData("Account_No", tbxAccount_No1.Text + tbxAccount_No2.Text + tbxAccount_No3.Text + tbxAccount_No4.Text, true, false, false));
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
        string strRemoteAddr = Request.ServerVariables["REMOTE_ADDR"];
        if (strRemoteAddr == "::1")
        {
            strRemoteAddr = "127.0.0.1";
        }
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

    }

    //---------------------------------------------------------------------------
    protected void btnConfirm_Click(object sender, EventArgs e)
    {

        //沒輸入防呆
        string strAccountNo = tbxAccount_No1.Text + tbxAccount_No2.Text + tbxAccount_No3.Text + tbxAccount_No4.Text;
        if (String.IsNullOrEmpty(strAccountNo)) return;

        //儲存線上定期授權暫存資料
        InsertPledgeTemp();
        //發送mail*****************************************

        //string SmtpServer = System.Configuration.ConfigurationManager.AppSettings["MailServer"];
        string MailSubject = "";
        string MailBody = "";
        string MailTo = "";
        string MailFrom = "";
        string strBody = "";

        strBody = "親愛的 " + HFD_DonorName.Value + " 平安！<br /><br />";
        strBody += "感謝您的奉獻！我們已收到您的定期授權資料，將會按您所指定的授權周期進行扣款。<br /><br />";
        //strBody += "您的奉獻授權資料如下：<br /><br />";

        //20140616 移除只剩「持卡人姓名」、「信用卡卡號」
        //strBody += "<table border='1' cellpadding='1' cellspacing='0'>";
        //strBody += "<tr><th align='right'>發卡銀行</th><td>" + txtCardBank.Text.Trim() + "</td>";
        //strBody += "<th align='right'>信用卡種類</th><td>" + ddlCardType.SelectedValue + "</td>";
        //strBody += "<span style='font-weight: bold;' align='center'>※　定　期　定　額　奉　獻</span>";
        //strBody += "<table border='1' width='80%'><tr><th><span>奉獻項目</span></th>";
        //strBody += "<th><span>奉獻周期</span></th><th><span>金額</span></th></tr>";
        //strBody += "<tr><td style='text-align: left;'><span>為GOODTV奉獻</span></td>";
        //strBody += "<td style='text-align: center;'><span>" + HFD_PayType.Value + "</span></td>";
        //strBody += "<td style='text-align: right;'><span><font color='red'>NT$ " + String.Format("{0:0,0}", Convert.ToInt32(HFD_Amount.Value)) + "</font></span></td>";
        //strBody += "</tr></table>";
        //strBody += "</table><br />";
        strBody += "您的奉獻資料如下：<br /><br />";
        strBody += "奉獻項目：為GOODTV奉獻<br />";
        strBody += "捐款方式：定期定額奉獻 " + HFD_PayType.Value + "<br />";
        strBody += "捐款金額：" + HFD_Amount.Value + " 元<br />";
        strBody += "<br />您輸入授權用信用卡資料如下：<br /><br />";
        strBody += "<span>持卡人姓名　　" + txtCardOwner.Text.Trim() + "<span><br />";
        strBody += "<span>信用卡卡號　　************" + tbxAccount_No4.Text.Trim() + "</span><br /><br />";
        //strBody += "<th align='right'>授權碼</th><td>" + txtAuthorize.Text.Trim() + "</td>";
        //strBody += "<tr><th align='right'>有效期限</th><td>" + ddlValidMonth.SelectedValue + "月" + ddlValidYear.SelectedValue + "年</td>";
        //strBody += "<th align='right'>與持卡人關係</th><td>" + txtRelation.Text.Trim() + "</td>";
        //strBody += "<th align='right'>持卡人身份證字號</th><td>" + ((txtOwnerIdNo.Text.Trim() == "") ? "" : txtOwnerIdNo.Text.Trim().Substring(0, 6)+"****") + "</td></table><br />";

        strBody += "再一次感謝您對 GOODTV 的奉獻支持，願神大大賜福給您！ <br />";
        strBody += "有任何關於定期授權問題請和我們聯絡： <br />總機：02-8024-3911<br />傳真：02-8024-3933<br />";
        strBody += "地址：235 新北市中和區中正路911號6樓 捐款服務組<br />";
        SendEMailObject MailObject = new SendEMailObject();
        //MailObject.SmtpServer = SmtpServer;
        MailSubject = " GOODTV線上奉獻（定期定額）確認信";
        MailBody = strBody;
        MailTo = strEmail;
        MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];

        string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, strBody);

        if (MailObject.ErrorCode != 0)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        }


        //發送內部通知mail *****************************************
        string strPhone;
        string strSql = @"select * from Donor where Donor_Id=@Donor_Id and DeleteDate is null ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", HFD_DonorID.Value);

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

            strBody = Request.Params["HFD_DonorName"] + " 奉獻資料如下：<br /><br />";
            //strBody += "您輸入授權用信用卡資料如下：<br /><br />";
            strBody += "捐款方式：定期定額奉獻 " + HFD_PayType.Value + "<br />";
            strBody += "捐款金額：" + HFD_Amount.Value + " 元<br />";
            strBody += "E-mail：" + dr["Email"].ToString() + "<br />";
            strBody += "電話：" + strPhone + "<br />";
            strBody += "收據地址：" +
                (strIsAbroad == "N" ? strCity + strArea + strAddress : strOverseasAddress + " " + strOverseasCountry) + "<br />";
            strBody += "收據寄發：" + strInvoiceType + "<br />";
            strBody += "徵信錄：" + (strIsAnonymous == "Y" ? "不刊登" : "刊登") + " <br />";

            MailObject = new SendEMailObject();
            //MailObject.SmtpServer = SmtpServer;
            MailSubject = " GOODTV線上奉獻_新奉獻通知(定期定額奉獻)";
            MailBody = strBody;
            MailTo = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];
            MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];

            result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, strBody);

            //********************************************

            ClientScript.RegisterStartupScript(this.GetType(), "postform", "<script> document.forms[0].action='DonateDone.aspx'; document.forms[0].submit(); </script>");
            //Response.Redirect("DonateDone.aspx");

        }
    }

}
