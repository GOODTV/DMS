using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateInfoConfirm: System.Web.UI.Page
{

    // 2015/2/25 增加Log
    static log4net.ILog logger = log4net.LogManager.GetLogger("RollingLogFileAppender");
    
    //for insert DONATE_WEB
    string strDonorID;
    string strDonorName;
    string strSex;
    string strIDNo;
    string strBirthday;
    string strEducation;
    string strOccupation;
    string strMarriage;
    string strCellPhone;
    string strTelOfficeRegion;
    string strTelOffice;
    string strTelOfficeExt;
    string strTelHome;
    string strZipCode;
    string strCityCode;
    string strAreaCode;
    string strAddress;
    string strEmail;
    string strInvoiceZipCode;
    string strInvoiceCityCode;
    string strInvoiceAreaCode;
    string strInvoiceAddress;
	string strPhone;
    //string strDonate_Type;
    string strInvoiceType;
    // 2014/7/3 增加金流網站變數
    string strIpayHttp = System.Configuration.ConfigurationManager.AppSettings["ipay_http"];
    string StrEmailToDonations = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];
    string MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
    string strIsAnonymous;
    string strCity;
    string strArea;
    string strIsAbroad;
    string strOverseasCountry;
    string strOverseasAddress;
    string strInvoice_Title;

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {

            // 2014/10/29 (重要，請勿刪除)修改至金流網站的變數值，更改從Web.config內的變數得來 Transaction Site value begin
            storeid.Value = System.Configuration.ConfigurationManager.AppSettings["storeid"];
            password.Value = System.Configuration.ConfigurationManager.AppSettings["password"];
            //不傳送給金流(金流網站的編碼為big5，姓名會亂碼，所以不傳送)
            //customer.Value = System.Configuration.ConfigurationManager.AppSettings["customer"];

            //POST接收值
            //捐款方式
            HFD_chkItem.Value = Request.Params["HFD_chkItem"];
            //奉獻金額
            HFD_Amount.Value = Request.Params["HFD_Amount"];
            //繳費方式
            HFD_PayType.Value = Request.Params["HFD_PayType"];
            //已登入帳號的捐款人編號
            HFD_DonorID.Value = Request.Params["HFD_DonorID"];

            // Transaction Site value end

            LoadFormData();

            if (HFD_chkItem.Value.Contains("單筆奉獻"))
            {
                this.msgOnce.Visible = true;
                ShowCartOnce();
                lblStep3.Text = " 確認奉獻明細 ";
                lblStep4.Visible = false;
                lblStep5.Visible = false;
            }
            else if (HFD_chkItem.Value.Contains("定期定額奉獻"))
            {
                this.msgPeriod.Visible = true;
                ShowCartPeriod();
                this.btnConfirm.Text = "確認, 開始填寫授權資料";
            }

        }
    }

    //---------------------------------------------------------------------------
    private void LoadFormData()
    {

        strDonorID = HFD_DonorID.Value;
        string strSql = @"select * from Donor where Donor_Id=@Donor_Id and DeleteDate is null ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", strDonorID);

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            DataRow dr = dt.Rows[0];
            lblDonorName.Text = "奉獻天使姓名：" + dr["Donor_Name"].ToString();
			
			if(!string.IsNullOrEmpty(dr["Tel_Office"].ToString()))
			{
				strPhone= "(" + dr["Tel_Office_Loc"].ToString() + ")" + dr["Tel_Office"].ToString() + " 分機: " + dr["Tel_Office_Ext"].ToString();
			}else{
                strPhone = dr["Cellular_Phone"].ToString();
			}
				
            lblPhoneNum.Text = "連絡電話：" + strPhone;	
			
			lblEmail.Text = "E-mail信箱：" + dr["Email"].ToString();
			
            //for 單筆奉獻金流使用
            //customer:　購買客戶的姓名 (可不填)
            //不傳送給金流(金流網站的編碼為big5，姓名會亂碼，所以不傳送)
            //customer.Value = dr["Donor_Name"].ToString();
            //cellphone:　購買客戶的行動電話 (可不填)
            cellphone.Value = dr["Cellular_Phone"].ToString();
            //param :　提供客戶自行運用(可不填)
            param.Value = dr["Donor_Id"].ToString();

            strDonorName = dr["Donor_Name"].ToString();
            HFD_DonorName.Value = dr["Donor_Name"].ToString();
            strSex = dr["Sex"].ToString();
            strIDNo = dr["IDNo"].ToString();
            strBirthday = dr["Birthday"].ToString();
            strEducation = dr["Education"].ToString();
            strOccupation = dr["Occupation"].ToString();
            strMarriage = dr["Marriage"].ToString();
            strCellPhone = dr["Cellular_Phone"].ToString();
            strTelOfficeRegion = dr["Tel_Office_Loc"].ToString();
            strTelOffice = dr["Tel_Office"].ToString();
            strTelOfficeExt = dr["Tel_Office_Ext"].ToString();
            strEmail = dr["Email"].ToString();
            strTelHome = dr["Tel_Home"].ToString();
            strZipCode = dr["ZipCode"].ToString();
            strCityCode = dr["City"].ToString();
            strAreaCode = dr["Area"].ToString();
            strAddress = dr["Address"].ToString();
            strInvoice_Title = dr["Invoice_Title"].ToString();

            if (dr["Invoice_City"].ToString() != "")
            {
                strInvoiceZipCode = dr["Invoice_ZipCode"].ToString();
                strInvoiceCityCode = dr["Invoice_City"].ToString();
                strInvoiceAreaCode = dr["Invoice_Area"].ToString();
                strInvoiceAddress = dr["Invoice_Address"].ToString();
            }
            else
            {
                strInvoiceZipCode = dr["ZipCode"].ToString();
                strInvoiceCityCode = dr["City"].ToString();
                strInvoiceAreaCode = dr["Area"].ToString();
                strInvoiceAddress = dr["Address"].ToString();
            }
            strCity = Util.GetCityName(dr["City"].ToString());  //聯絡地址-縣市
            strArea = Util.GetAreaName(dr["Area"].ToString());  //鄉鎮 Area

            // 2014/6/5 修正收據開立判斷
            strInvoiceType = dr["Invoice_Type"].ToString();

            // 2014/7/4 增加徵信錄欄位,台灣地址
            strIsAnonymous = dr["IsAnonymous"].ToString();
            strIsAbroad = dr["IsAbroad"].ToString();
            strOverseasCountry = dr["OverseasCountry"].ToString();
            strOverseasAddress = dr["OverseasAddress"].ToString();

        }
    }
    //---------------------------------------------------------------------------
    protected void btnConfirm_Click(object sender, EventArgs e)
    {

        // 2014/11/6 增加金流網站回覆頁面(DonateReturnUrl.aspx)的網頁重新整理記號清除
        //Session.Remove("IsRefresh");

        if (HFD_chkItem.Value == "定期定額奉獻")
        {
            //Session["DonateContent"] = lblDonateContent.Text;
            ClientScript.RegisterStartupScript(this.GetType(), "s", "<script> document.forms[0].action='DonateCreditCard.aspx'; document.forms[0].submit(); </script>");
            //Response.Redirect("DonateCreditCard.aspx");
        }
        else if (HFD_chkItem.Value == "單筆奉獻")
        {
            //Response.Redirect(Util.RedirectByTime("DonateSingle.aspx"));

            //for 單筆奉獻金流使用

            LoadFormData();

            //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
            //orderid:　訂單編號(不得重複,勿超過15碼)
            orderid.Value = DateTime.Now.Ticks.ToString().Substring(0, 15);
 
            //purpose: 項目                    
            purpose.Value = "為GOODTV奉獻";
            //account:　金額(正整數)
            account.Value = HFD_Amount.Value;

            //新增交易資料(DONATE_WEB)
            InsertDonateWeb();

            // 2014/7/4 增加內部Email通知
            InternalSendMail();
            logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", DonateInfoConfirm.aspx.cs 程式內部: btnConfirm_Click 單筆奉獻 介接思遠金流, URL:" + HttpUtility.UrlDecode(Request.Form.ToString()));

            //介接思遠金流
            string script = "<script> document.forms[0].action='" + strIpayHttp + "'; document.forms[0].submit(); </script>";
            ClientScript.RegisterStartupScript(this.GetType(), "postform", script);
        }
        else
        {
            Response.Redirect("DonateOnlineAll.aspx");
        }

    }

    private void InternalSendMail()
    {
        //發送內部通知mail*****************************************
        string strBody = strDonorName + " 奉獻資料如下：<BR/><BR/>";
        strBody += "捐款方式：單筆奉獻 <BR/>";
        strBody += "捐款金額：" + account.Value + " 元 <BR/>";
        strBody += "E-mail：" + strEmail + "  <BR/>";
        strBody += "電話：" + strPhone + "<BR/>";
        strBody += "收據地址：" + (strIsAbroad == "N" ? strCity + strArea + strAddress : strOverseasAddress + " " + strOverseasCountry) + "<BR/>";
        strBody += "收據寄發：" + strInvoiceType + " <BR/>";
        strBody += "徵信錄：" + (strIsAnonymous == "Y" ? "不刊登" : "刊登") + " <BR/>";

        SendEMailObject MailObject = new SendEMailObject();
        string MailSubject = " GOODTV線上奉獻_新奉獻通知(" + strDonorName + "_單筆奉獻_等待授權)";
        string MailBody = strBody;

        string result = MailObject.SendEmail(StrEmailToDonations, MailFrom, MailSubject, strBody);

        if (MailObject.ErrorCode != 0)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        }
    
    }

    //-------------------------------------------------------------------------------------------------------------
    private void InsertDonateWeb()
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();
        SetColumnDataDonateWeb(list);
        string strSql = "";
        strSql = Util.CreateInsertCommand("DONATE_WEB", list, dict);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    //-------------------------------------------------------------------------------------------------------------
    private void SetColumnDataDonateWeb(List<ColumnData> list)
    {
        //新增交易資料(DONATE_WEB)
        // 2014/6/4 修正收據開立判斷
        list.Add(new ColumnData("Donate_Type", "單次捐款", true, false, false));
        //list.Add(new ColumnData("Donate_Type", strDonate_Type, true, false, false));
        list.Add(new ColumnData("od_sob", orderid.Value, true, false, false));
        list.Add(new ColumnData("Dept_Id", "C001", true, false, false));
        list.Add(new ColumnData("Donor_Id", strDonorID, true, false, false));
        list.Add(new ColumnData("Donate_CreateDate", Util.GetDBDateTime().ToShortDateString(), true, false, false));
        list.Add(new ColumnData("Donate_CreateDateTime", Util.GetDBDateTime(), true, false, false));
        string strRemoteAddr = Request.ServerVariables["REMOTE_ADDR"];
        if (strRemoteAddr == "::1")
        {
            strRemoteAddr = "127.0.0.1";
        }
        list.Add(new ColumnData("Donate_CreateIP", strRemoteAddr, true, false, false));
        list.Add(new ColumnData("Donate_Amount", account.Value, true, false, false));
        //list.Add(new ColumnData("Donate_DonorName", customer.Value, true, false, false));
        list.Add(new ColumnData("Donate_DonorName", strDonorName, true, false, false));
        list.Add(new ColumnData("Donate_Sex", strSex, true, false, false));
        list.Add(new ColumnData("Donate_IDNO", strIDNo, true, false, false));
        list.Add(new ColumnData("Donate_Birthday", (String.IsNullOrEmpty(strBirthday) ? null : Util.FixDateTime(strBirthday)), true, false, false));
        list.Add(new ColumnData("Donate_Education", strEducation, true, false, false));
        list.Add(new ColumnData("Donate_Occupation", strOccupation, true, false, false));
        list.Add(new ColumnData("Donate_Marriage", strMarriage, true, false, false));
        list.Add(new ColumnData("Donate_CellPhone", strCellPhone, true, false, false));
        list.Add(new ColumnData("Donate_TelOffice_Region", strTelOfficeRegion, true, false, false));
        list.Add(new ColumnData("Donate_TelOffice", strTelOffice, true, false, false));
        list.Add(new ColumnData("Donate_TelOffice_Ext", strTelOfficeExt, true, false, false));
        list.Add(new ColumnData("Donate_TelHome", strTelHome, true, false, false));
        list.Add(new ColumnData("Donate_ZipCode", strZipCode, true, false, false));
        list.Add(new ColumnData("Donate_CityCode", strCityCode, true, false, false));
        list.Add(new ColumnData("Donate_AreaCode", strAreaCode, true, false, false));
        list.Add(new ColumnData("Donate_Address", strAddress, true, false, false));
        list.Add(new ColumnData("Donate_Email", strEmail, true, false, false));
        //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
        //list.Add(new ColumnData("Donate_Purpose", purpose.Value, true, false, false));
        list.Add(new ColumnData("Donate_Purpose", "經常費", true, false, false));
        // 2014/6/4 修正收據開立判斷
        //list.Add(new ColumnData("Donate_Invoice_Type", "單次收據", true, false, false));
        list.Add(new ColumnData("Donate_Invoice_Type", strInvoiceType, true, false, false));
        list.Add(new ColumnData("Donate_Invoice_Title", strInvoice_Title, true, false, false));
        list.Add(new ColumnData("Donate_Invoice_IDNo", strIDNo, true, false, false));
        list.Add(new ColumnData("Donate_Invoice_ZipCode", strInvoiceZipCode, true, false, false));
        list.Add(new ColumnData("Donate_Invoice_CityCode", strInvoiceCityCode, true, false, false));
        list.Add(new ColumnData("Donate_Invoice_AreaCode", strInvoiceAreaCode, true, false, false));
        list.Add(new ColumnData("Donate_Invoice_Address", strInvoiceAddress, true, false, false));
        // 2014/6/4 先改為Donate_Update='T' 暫停此功能
        //list.Add(new ColumnData("Donate_Update", "N", true, false, false));
        list.Add(new ColumnData("Donate_Update", "T", true, false, false));
        list.Add(new ColumnData("Donate_Export", "N", true, false, false));

    }

    //---------------------------------------------------------------------------
    protected void btnGoHome_Click(object sender, EventArgs e)
    {
        Response.Redirect("DonateOnlineAll.aspx");
    }

    //-------------------------------------------------------------------------
    public void ShowCartOnce()
    {
        lblDonateContent.Text = "<span style='font-weight: bold;'>※ 單 筆 奉 獻</span>";
        lblDonateContent.Text += "<table width='100%' class='table_h'><tr style='background-color: rgb(205, 225, 226);'>";
        lblDonateContent.Text += "<th><span>用途</span></th><th><span>金額</span></th></tr><tr><td style='text-align: center;'>";
        lblDonateContent.Text += "<span>為GOODTV奉獻</span></td><td style='text-align: right;'><span><font color='red'>NT$ " +
            String.Format("{0:0,0}", Convert.ToInt32(HFD_Amount.Value)) + "</font></span></td></tr></table>";
    }

    //-------------------------------------------------------------------------
    public void ShowCartPeriod()
    {

        lblDonateContent.Text = "<span style='font-weight: bold;text-align: center;'>※ 定 期 定 額 奉 獻</span>";
        lblDonateContent.Text += "<table width='100%' class='table_h'><tr><th><span>奉獻項目</span></th>";
        lblDonateContent.Text += "<th><span>奉獻週期</span></th><th><span>金額</span></th></tr><tr><td style='text-align: center;'><span>為GOODTV奉獻</span></td>";
        lblDonateContent.Text += "<td style='text-align: center;'><span>" + Request.Params["HFD_PayType"] + "</span></td>";
        lblDonateContent.Text += "<td style='text-align: right;'><span><font color='red'>NT$ " +
            String.Format("{0:0,0}", Convert.ToInt32(HFD_Amount.Value)) + "</font></span></td></tr></table>";
    }

}
