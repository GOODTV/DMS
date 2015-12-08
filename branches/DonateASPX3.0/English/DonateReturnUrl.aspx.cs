using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateReturnUrl : System.Web.UI.Page
{
    DataTable dtOnce = new DataTable();
    DataTable dtPeriod = new DataTable();
    string strOrderId;
    string strStatus;
    string strAmount;
    string strPayType;
    string strDonorId;
    string Mailhead = "";
    string SmtpServer = System.Configuration.ConfigurationSettings.AppSettings["MailServer"];
    string MailSubject = "";
    string MailBody = "";
    string MailTo = "";
    string MailFrom = System.Configuration.ConfigurationSettings.AppSettings["MailFrom"];
    string strBody = "";
    string gdonor_name = "";
    string strSql = "";
    string gEmail = "";
    string gTitle = "";
    string gType = "";
    string gIsAnonymous = "";
    string gAbroad = "";
    string gCity = "";
    string gArea = "";
    string gZipcode = "";
    string gStreet = "";
    string gSection = "";
    string gLane = "";
    string gAlley = "";
    string gHouseNo = "";
    string gHouseNoSub = "";
    string gFloor = "";
    string gFloorSub = "";
    string gRoom = "";
    string gAttn = "";
    string strInvoice_No = "";
    string strInvoice_Pre = "";
    //string strDonate_Type;
    //string strInvoiceType;
    string gInvoice_Address = "";
    // 2014/7/3 增加金流網站變數
    //string strIpayHttp = System.Configuration.ConfigurationSettings.AppSettings["ipay_http"];
    string StrEmailToDonations = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];
    string strPayTypeName = "";
    string strPayformat = "";
    string strAuthdate = "";

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {

            //LoadFormData();
            //if (Session["DonorName"] != null)
            //{
            //    lblTitle.Text = Session["DonorName"].ToString() + lblTitle.Text;
            //}
            //LoadDropDownList();

            //if (dtOnce.Rows.Count > 0)
            //{
            //    ShowCartOnce(dtOnce);
            //}

            //if (dtPeriod.Rows.Count > 0)
            //{
            //    ShowCartPeriod(dtPeriod);
            //}

            strOrderId = Request.Form["orderid"].ToString();
            strStatus = Request.Form["status"].ToString();
            strAmount = Request.Form["account"].ToString();
            strPayType = Request.Form["paytype"].ToString();
            strDonorId = Request.Form["param"].ToString();
            strPayformat = Request.Form["payformat"].ToString();
            strAuthdate = Request.Form["authdate"].ToString();
            account.Value = Request.Form["account"].ToString();
            param.Value = Request.Form["param"].ToString();

            if (Session["ItemOnce"] != null)
            {
                dtOnce = (DataTable)Session["ItemOnce"];
                //strDonate_Type = "單次捐款";
                //strInvoiceType = (Util.GetDBValue("Donor", "Invoice_Type", "Donor_Id", strDonorId) == "不寄") ? "不寄" : "單次收據";
            }
            if (Session["ItemPeriod"] != null)
            {
                dtPeriod = (DataTable)Session["ItemPeriod"];
                //strDonate_Type = "長期捐款";
                //strInvoiceType = (Util.GetDBValue("Donor", "Invoice_Type", "Donor_Id", strDonorId) == "不寄") ? "不寄" : "年度證明";
            }

            // 2014/9/24 付款方式
            switch (strPayType)
            {
                case "1":
                    strPayTypeName = "信用卡";
                    break;
                case "8":
                    strPayTypeName = "WEB ATM";
                    break;
                case "16":
                    strPayTypeName = "郵局電子帳單付款";
                    break;
                case "17":
                    strPayTypeName = "郵局ATM轉帳";
                    break;
                case "2":
                    strPayTypeName = "iePay儲值帳戶付款";
                    break;
                case "4":
                    strPayTypeName = "PayPal";
                    break;
                case "5":
                    strPayTypeName = "其他超商 電子帳單";
                    break;
                case "9":
                    strPayTypeName = "7-11 電子帳單";
                    break;
                case "12":
                    strPayTypeName = "玉山銀行eCoin";
                    break;
                case "25":
                    strPayTypeName = "24hr超商取貨付款";
                    break;
                case "30":
                    strPayTypeName = "7-11 ibon";
                    break;
                case "35":
                    strPayTypeName = "條碼";
                    break;
                case "39":
                    strPayTypeName = "全家FamilyPort";
                    break;
                default:
                    strPayTypeName = "未知";
                    break;
            }

            lblTitle.Text += "<br/><br/>交易序號：" + strOrderId;
            InsertIEPay();
            if (strStatus == "0")
            {
                this.btnBackPay.Visible = false;
                //20140710 更改授權狀態顯示內容
                switch (strPayType)
                {
                    case "1":
                        lblTitle.Text += "<br/><br/>授權狀態：成功"; //信用卡
                        break;
                    case "8":
                        lblTitle.Text += "<br/><br/>授權狀態：成功"; //WEB ATM
                        break;
                    case "16":
                        lblTitle.Text += "<br/><br/>捐款流程已完成：請至郵局使用電子帳單繳款即可"; //郵局電子帳單付款
                        break;
                    case "17":
                        lblTitle.Text += "<br/><br/>捐款流程已完成：請使用郵局ATM轉帳"; //郵局ATM轉帳
                        break;
                    //正式金流付款方式
                    case "2":
                        lblTitle.Text += "<br/><br/>捐款流程已完成：請使用您的iePay儲值帳戶繳款即可"; //iePay儲值帳戶付款
                        break;
                    case "4":
                        lblTitle.Text += "<br/><br/>捐款流程已完成：請使用PayPal繳款即可"; //PayPal
                        break;
                    case "5":
                        lblTitle.Text += "<br/><br/>捐款流程已完成：請至其他超商使用電子帳單繳款即可"; //其他超商 電子帳單
                        break;
                    case "9":
                        lblTitle.Text += "<br/><br/>捐款流程已完成：請至7-11超商店內使用電子帳單繳款即可"; //7-11 電子帳單
                        break;
                    case "12":
                        lblTitle.Text += "<br/><br/>捐款流程已完成：請使用玉山銀行eCoin功能繳款即可"; //玉山銀行eCoin
                        break;
                    case "25":
                        lblTitle.Text += "<br/><br/>捐款流程已完成：請至24hr超商取貨繳款即可"; //24hr超商取貨付款
                        break;
                    case "30":
                        lblTitle.Text += "<br/><br/>捐款流程已完成：請記下序號，至7-11超商店內使用ibon機器列印出帳單後繳款即可"; //7-11 ibon
                        break;
                    case "35":
                        lblTitle.Text += "<br/><br/>捐款流程已完成：請持條碼繳費單至超商或銀行繳款即可"; //條碼
                        break;
                    case "39":
                        lblTitle.Text += "<br/><br/>捐款流程已完成：請記下序號，至全家超商店內使用FamiPort機器列印出帳單後至櫃台繳款即可"; //全家FamilyPort
                        break;
                    default:
                        lblTitle.Text += "<br/><br/>捐款流程已完成";
                        break;
                }

                Session["sta"] = "OK";

                // 2014/7/9 只有線上信用卡轉帳授權成功才算是確定繳款
                if (strPayType == "1")
                {
                    InsertDonate();
                    // 2014/7/3 增加重新統計捐款人捐款資訊
                    UpdateDonor();
                    // 2014/9/24 增加授權成功內部通知
                    InternalSendMail(true);
                }
                dtOnce.Rows[0].Delete();
                //Session["InsertPeriod"] = null;
                // 2014/6/6 增加完成線上奉獻就清除Session
                Session.Clear();

            }
            else
            {
                this.btnBackPay.Visible = true;

                lblTitle.Text += "<br/><br/>授權狀態：失敗 (" + strStatus +")";
                Session["sta"] = "";
                // 2014/8/11 增加奉獻天使授權失敗的內部通知信
                InternalSendMail(false);

            }
            lblTitle.Text += "<br/><br/>您的奉獻金額：" + strAmount + " 元";

            // 2014/6/6 改變判斷邏輯方式
            switch (strPayType)
            {
                case "1":
                    lblTitle.Text += "<br/><br/>付款方式：信用卡";
                    break;
                case "8":
                    lblTitle.Text += "<br/><br/>付款方式：WEB ATM";
                    break;
                case "16":
                    lblTitle.Text += "<br/><br/>付款方式：郵局電子帳單付款";
                    break;
                case "17":
                    lblTitle.Text += "<br/><br/>付款方式：郵局ATM轉帳";
                    break;
                //20140702 新增 正式金流付款方式
                case "2":
                    lblTitle.Text += "<br/><br/>付款方式：iePay儲值帳戶付款";
                    break;
                case "4":
                    lblTitle.Text += "<br/><br/>付款方式：PayPal";
                    break;
                case "5":
                    lblTitle.Text += "<br/><br/>付款方式：其他超商 電子帳單";
                    break;
                case "9":
                    lblTitle.Text += "<br/><br/>付款方式：7-11 電子帳單";
                    break;
                case "12":
                    lblTitle.Text += "<br/><br/>付款方式：玉山銀行eCoin";
                    break;
                case "25":
                    lblTitle.Text += "<br/><br/>付款方式：24hr超商取貨付款";
                    break;
                case "30":
                    lblTitle.Text += "<br/><br/>付款方式：7-11 ibon";
                    break;
                case "35":
                    lblTitle.Text += "<br/><br/>付款方式：條碼";
                    break;
                case "39":
                    lblTitle.Text += "<br/><br/>付款方式：全家FamilyPort";
                    break;
                default:
                    lblTitle.Text += "<br/><br/>付款方式代碼：" + strPayType;
                    break;
            }
            /*
            if (strPayType == "1")
            {
                lblTitle.Text += "<br/><br/>付款方式：信用卡";
            }
			if (strPayType == "8")
            {
                lblTitle.Text += "<br/><br/>付款方式：WEB ATM";
            }
			if (strPayType == "16")
            {
                lblTitle.Text += "<br/><br/>付款方式：郵局電子帳單付款";
            }
			if (strPayType == "17")
            {
                lblTitle.Text += "<br/><br/>付款方式：郵局ATM轉帳";
            }
            */
            //else
            //{
            //   lblTitle.Text += "<br/><br/>付款方式代碼：" + strPayType;
            //}
            //lblTitle.Text += "<br/>param：" + Request.Form["param"].ToString();

            // 2014/6/10 Session 改成從資料庫取得
            lblTitle.Text = gdonor_name + lblTitle.Text;

        }
    }

    private void UpdateDonor()
    {

        string strSQL = @"DECLARE @Begin_DonateDate datetime " +
       "DECLARE @Last_DonateDate datetime " +
       "DECLARE @Donate_No numeric " +
       "DECLARE @Donate_Total numeric " +
       "Select Top 1 @Begin_DonateDate=Donate_Date From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 Order By Donate_Date " +
       "Select Top 1 @Last_DonateDate=Donate_Date From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 Order By Donate_Date Desc " +
       "Select @Donate_No=Count(*) From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 " +
       "Select @Donate_Total=IsNull(Sum(Donate_Amt),0) From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 " +
       "Update DONOR Set Begin_DonateDate=@Begin_DonateDate,Last_DonateDate=@Last_DonateDate,Donate_No=@Donate_No,Donate_Total=@Donate_Total Where Donor_Id=@Donor_Id ; ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", strDonorId);
        NpoDB.ExecuteSQLS(strSQL, dict);

    }

    private void InternalSendMail(bool YesNo)
    {

        //if (Session["sta"] == "")
        //{
            DataTable dt = null;
            DataRow dr = null;
            Dictionary<string, object> dict1 = new Dictionary<string, object>();
            strSql = @"select *
                from DONOR 
                where Donor_Id=@DonorID
                ";
            dict1.Add("DonorID", strDonorId);
            dt = NpoDB.GetDataTableS(strSql, dict1);
            if (dt.Rows.Count != 0)
            {
                dr = dt.Rows[0];
                gdonor_name = dr["donor_name"].ToString();
                gEmail = dr["Email"].ToString();
                gTitle = dr["Invoice_Title"].ToString();
                gType = dr["Invoice_Type"].ToString();
                gIsAnonymous = dr["IsAnonymous"].ToString() == "Y" ? "不刊登" : "刊登";
                gCity = Util.GetCityName(dr["City"].ToString());  //聯絡地址-縣市
                gArea = Util.GetAreaName(dr["Area"].ToString());  //鄉鎮 Area
                //gZipcode = dr["ZipCode"].ToString();
                if (dr["Invoice_Address"].ToString() == "")
                {
                    //通訊地址
                    gAbroad = dr["IsAbroad"].ToString() == "Y" ? "海外地區" : "台灣";
                    gInvoice_Address = dr["Address"].ToString();
                    gAttn = dr["Attn"].ToString(); //Attn
                }
                else
                {
                    //收據地址
                    gAbroad = dr["IsAbroad_Invoice"].ToString() == "Y" ? "海外地區" : "台灣";
                    gInvoice_Address = dr["Invoice_Address"].ToString();
                    gAttn = dr["Invoice_Attn"].ToString(); //Attn
                }

            }

            string strAddr = (gAbroad == "台灣") ? gAbroad + gCity + gArea + gInvoice_Address + gAttn : gAbroad + gInvoice_Address;
            string strPhone;
            if (!string.IsNullOrEmpty(dr["Tel_Office"].ToString()))
            {
                strPhone = "(" + dr["Tel_Office_Loc"].ToString() + ")" + dr["Tel_Office"].ToString() + " 分機: " + dr["Tel_Office_Ext"].ToString();
            }
            else
            {
                strPhone = dr["Cellular_Phone"].ToString();
            }

            //發送內部通知mail*****************************************
            string strBody = gdonor_name + " 奉獻資料如下：<BR/><BR/>";
            strBody += "捐款方式：單筆奉獻 <BR/>";
            strBody += "付款方式：" + strPayTypeName + "<BR/>";
            strBody += "訂單編號：" + strOrderId + " <BR/>";
            strBody += "付款狀態：" + (YesNo == true ? "授權成功" : "授權失敗 代碼：" + strStatus) + "<BR/>";
            strBody += String.Format("捐款金額：{0:#,0} 元 <BR/>", strAmount);
            strBody += "E-mail：" + gEmail + "  <BR/>";
            strBody += "電話：" + strPhone + "<BR/>";
            strBody += "收據地址：" + strAddr + "<BR/>";
            strBody += "收據寄發：" + gType + " <BR/>";
            strBody += "徵信錄：" + gIsAnonymous + " <BR/>";

            SendEMailObject MailObject = new SendEMailObject();
            string MailSubject = " GOODTV線上奉獻_新奉獻通知(" + gdonor_name + "_單筆奉獻_" + (YesNo == true ? "授權成功" : "授權失敗") + ")";
            string MailBody = strBody;

            string result = MailObject.SendEmail(StrEmailToDonations, MailFrom, MailSubject, strBody);

            if (MailObject.ErrorCode != 0)
            {
                this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
            }
            //********************************************
        //}
    }

    //-------------------------------------------------------------------------------------------------------------
    private void InsertIEPay()
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();
        SetColumnDataIEPay(list);
        string strSql = "";
        strSql = Util.CreateInsertCommand("DONATE_IEPAY", list, dict);
        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        String strSql_check = "select orderid from DONATE_IEPAY where orderid = '" + strOrderId + "'";
        DataTable dt = NpoDB.GetDataTableS(strSql_check, dict2);
        if (dt.Rows.Count == 0)
        {
            NpoDB.ExecuteSQLS(strSql, dict);
        }
        else
        {
            this.Page.RegisterStartupScript("s", "<script>alert('請勿重新整理網頁！');</script>");
        }
    }
    //-------------------------------------------------------------------------------------------------------------
    private void SetColumnDataIEPay(List<ColumnData> list)
    {
        //orderid
        list.Add(new ColumnData("orderid", strOrderId, true, false, false));
        //status
        list.Add(new ColumnData("status", strStatus, true, false, false));
        //account
        list.Add(new ColumnData("account", strAmount, true, false, false));
        //paytype
        list.Add(new ColumnData("paytype", strPayType, true, false, false));
        //payformat
        list.Add(new ColumnData("payformat", strPayformat /*Request.Form["payformat"].ToString()*/, true, false, false));
        //authdate
        list.Add(new ColumnData("authdate", strAuthdate /*Request.Form["authdate"].ToString()*/, true, false, false));
        //param
        list.Add(new ColumnData("param", strDonorId /*Request.Form["param"].ToString()*/, true, false, false));
        ////authcode
        //list.Add(new ColumnData("authcode", Util.DateTime2String(txtBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnNull), true, true, false));
        ////return_date
        //list.Add(new ColumnData("return_date", txtIDNo.Text, true, true, false));
        ////return_ip
        //list.Add(new ColumnData("return_ip", Util.GetControlValue("RDO_LiveRegion") == "台灣" ? "N" : "Y", true, true, false));
        ////return_url1
        //list.Add(new ColumnData("return_url1", Util.GetToday(DateType.yyyyMMdd), true, false, false));
    }
    //-------------------------------------------------------------------------------------------------------------
    private void InsertDonate()
    {      
        //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
        for (int i = 0; i < dtOnce.Rows.Count; i++)
        {
            Dictionary<string, object> dict = new Dictionary<string, object>();
            List<ColumnData> list = new List<ColumnData>();
            
            string purpose = dtOnce.Rows[i]["奉獻項目"].ToString();
            string account = dtOnce.Rows[i]["奉獻金額"].ToString();
            
            SetColumnDataDonate(list,purpose,account);
            string strSql = "";
            strSql = Util.CreateInsertCommand("DONATE", list, dict);
            NpoDB.ExecuteSQLS(strSql, dict);
        }

        //發送mail*****************************************
        if (Session["sta"] == "OK")
        {
            DataTable dt = null;
            DataRow dr = null;
            Dictionary<string, object> dict1 = new Dictionary<string, object>();
            strSql = @"select *
                from DONOR 
                where Donor_Id=@DonorID
                ";
            dict1.Add("DonorID", strDonorId);
            dt = NpoDB.GetDataTableS(strSql, dict1);
            if (dt.Rows.Count != 0)
            {
                dr = dt.Rows[0];
                gdonor_name = dr["donor_name"].ToString();
                gEmail = dr["Email"].ToString();
                gTitle = dr["Invoice_Title"].ToString();
                // 2014/6/4 修正收據開立判斷
                //gType = dr["Invoice_Type"].ToString() == "Y" ? "是" : "否";
                gType = dr["Invoice_Type"].ToString() == "不寄" ? "否" : "是";
                gIsAnonymous = dr["IsAnonymous"].ToString() == "Y" ? "不刊登" : "刊登";
                gAbroad = dr["IsAbroad"].ToString() == "Y" ? "海外地區" : "台灣";//居住地區
                gCity = Util.GetCityName(dr["City"].ToString());  //聯絡地址-縣市
                gArea = Util.GetAreaName(dr["Area"].ToString());  //鄉鎮 Area
                gZipcode = dr["ZipCode"].ToString();
                gStreet = dr["Street"].ToString();//街道
                gSection = dr["Section"].ToString(); //段
                gLane = dr["Lane"].ToString();  //巷
                gAlley = dr["Alley"].ToString();//弄
                gHouseNo = dr["HouseNo"].ToString(); //號
                gHouseNoSub = dr["HouseNoSub"].ToString(); //號之
                gFloor = dr["Floor"].ToString();  //樓
                gFloorSub = dr["FloorSub"].ToString();  //樓之
                gRoom = dr["Room"].ToString(); //室
                gAttn = dr["Attn"].ToString(); //Attn
                gInvoice_Address = dr["Invoice_Address"].ToString(); ;

            }

            string strAddr = "";
            string strAddr2 = "";

            if (gStreet != "")
            {
                strAddr += gStreet;
            }
            if (gSection != "")
            {
                strAddr += gSection + "段";
            }
            if (gLane != "")
            {
                strAddr += gLane + "巷";
            }
            if (gAlley != "")
            {
                strAddr += gAlley + "弄";
            }
            if (gHouseNo != "")
            {
                strAddr += gHouseNo;
            }
            if (gHouseNoSub != "")
            {
                strAddr += "之" + gHouseNoSub;
            }
            if (gHouseNo != "")
            {
                strAddr += "號";
            }
            if (gFloor != "")
            {
                strAddr += gFloor + "樓"; ;
            }
            if (gFloorSub != "")
            {
                strAddr += "之" + gFloorSub;
            }
            if (gRoom != "")
            {
                strAddr += " -" + gRoom + "室";
            }
            if (gAttn != "")
            {
                strAddr += "(" + gAttn + ")";
            }
            string Sname = "";
            strAddr2 = (gAbroad == "台灣") ? gAbroad + gCity + gArea + strAddr : gAbroad + gInvoice_Address;
            Sname = Util.GetDBValue("donor", "Donor_Name", "donor_ID", strDonorId);
            strInvoice_No = Util.GetDBValue("DONATE", "Invoice_No", "od_sob", strOrderId);
            strInvoice_Pre = Util.GetDBValue("DONATE", "Invoice_Pre", "od_sob", strOrderId);

            strBody = "Dear " + gdonor_name + " <BR><BR>";
            strBody += "Thank you for your monetary gift：<BR>";
            strBody += "Your account(email): " + gEmail + " <BR>";
            strBody += "Receipt No.：" + strInvoice_No + "  <BR>";
            strBody += "Name of Receipt：" + gTitle + "<BR>";
            strBody += "Amount：NT$" + strAmount + "元<BR>";
            strBody += "Mailing address：" + strAddr2 + " <BR>";
            if (gType=="是")
                strBody += "Send receipt：Yes <BR>";
            else
                strBody += "Send receipt：No <BR>";
            //strBody += "徵信錄：" + gIsAnonymous + " <BR><BR>";
            //strBody += "信用卡授權時間：" + Util.GetDBDateTime() + "<P>";
            strBody += "Any question you may call us at: 02-8024-3911.<BR>";
            strBody += "May God bless you!<BR>";
            strBody += "Yours sincerely,<BR>";
            strBody += "GOODTV<BR>";
            SendEMailObject MailObject = new SendEMailObject();
            MailObject.SmtpServer = SmtpServer;
            MailSubject = " GOODTV - Donation Confirmation";
            MailBody = strBody;
            MailTo = gEmail;

            string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, strBody);

            if (MailObject.ErrorCode != 0)
            {
                this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
            }
            //********************************************

        }

    }
    //-------------------------------------------------------------------------------------------------------------
    private void SetColumnDataDonate(List<ColumnData> list, string purpose, string account)
    {
      
        //'新增捐款紀錄
        //Donate_Fee=0
        //Donate_Accou=Request.Form("amount")
        //If Request.Form("amount")<>"" Then
        //  Donate_Fee=CDbl(Request.Form("amount"))*0.02
        //  Donate_Accou=CDbl(Request.Form("amount"))-CDbl(Donate_Fee)
        //End If
        //SQL3="DONATE"
        //Set RS3=Server.CreateObject("ADODB.RecordSet")
        //RS3.Open SQL3,Conn,1,3
        //RS3.Addnew
        //RS3("od_sob")=Request.Form("orderid")
        list.Add(new ColumnData("od_sob", strOrderId, true, false, false));
        //RS3("Donor_Id")=RS1("Donor_Id")
        list.Add(new ColumnData("Donor_Id", strDonorId, true, false, false));
        //RS3("Donate_Date")=Date()
        //2014/6/4 改正捐款日期不含時間
        list.Add(new ColumnData("Donate_Date", Util.GetDBDateTime().ToShortDateString(), true, false, false));
        //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
        //RS3("Donate_Amt")=Request.Form("amount")
        list.Add(new ColumnData("Donate_Amt", account, true, false, false));
        //RS3("Donate_AmtD")=Request.Form("amount")
        list.Add(new ColumnData("Donate_AmtD", account, true, false, false));
        //RS3("Donate_Fee")=Donate_Fee
        list.Add(new ColumnData("Donate_Fee", "0", true, false, false));
        //RS3("Donate_FeeD")=Donate_Fee
        list.Add(new ColumnData("Donate_FeeD", "0", true, false, false));
        //RS3("Donate_Accou")=Donate_Accou
        list.Add(new ColumnData("Donate_Accou", account, true, false, false));
        //RS3("Donate_AccouD")=Donate_Accou
        list.Add(new ColumnData("Donate_AccouD", account, true, false, false));
        //RS3("Donate_AmtM")="0"
        list.Add(new ColumnData("Donate_AmtM", "0", true, false, false));
        //RS3("Donate_FeeM")="0"
        list.Add(new ColumnData("Donate_FeeM", "0", true, false, false));
        //RS3("Donate_AccouM")="0"
        list.Add(new ColumnData("Donate_AccouM", "0", true, false, false));
        //RS3("Donate_AmtA")="0"
        list.Add(new ColumnData("Donate_AmtA", "0", true, false, false));
        //RS3("Donate_FeeA")="0"
        list.Add(new ColumnData("Donate_FeeA", "0", true, false, false));
        //RS3("Donate_AccouA")="0"
        list.Add(new ColumnData("Donate_AccouA", "0", true, false, false));
        //RS3("Donate_AmtS")="0"
        list.Add(new ColumnData("Donate_AmtS", "0", true, false, false));
        //RS3("Donate_RateS")="0"
        list.Add(new ColumnData("Donate_RateS", "0", true, false, false));
        //RS3("Donate_FeeS")="0"
        list.Add(new ColumnData("Donate_FeeS", "0", true, false, false));
        //RS3("Donate_AccouS")="0"
        list.Add(new ColumnData("Donate_AccouS", "0", true, false, false));
        //RS3("Donate_Payment")="網路信用卡"
        list.Add(new ColumnData("Donate_Payment", "網路信用卡", true, false, false));
        //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
        //RS3("Donate_Purpose")=RS1("Donate_Purpose")
        //list.Add(new ColumnData("Donate_Purpose", purpose, true, false, false));
        list.Add(new ColumnData("Donate_Purpose", "經常費", true, false, false));
        //RS3("Donate_Purpose_Type")="D"
        list.Add(new ColumnData("Donate_Purpose_Type", "D", true, false, false));
        //RS3("Donate_Type")="單次捐款"
        // 2014/6/5 修正收據開立判斷
        list.Add(new ColumnData("Donate_Type", "單次捐款", true, false, false));
        //list.Add(new ColumnData("Donate_Type", strDonate_Type, true, false, false));
        //RS3("Donate_Forign")=""
        //RS3("Donate_Desc")=""
        //RS3("IsBeductible")="N"
        list.Add(new ColumnData("IsBeductible", "N", true, false, false));
        //RS3("Donate_Amt2")="0"
        list.Add(new ColumnData("Donate_Amt2", "0", true, false, false));
        //RS3("Card_Bank")=""
        //RS3("Card_Type")=""
        //RS3("Account_No")=""
        //RS3("Valid_Date")=""
        //RS3("Card_Owner")=""
        //RS3("Owner_IDNo")=""
        //RS3("Relation")=""
        //RS3("Authorize")=""
        //RS3("Check_No")=""
        //RS3("Check_ExpireDate")=null
        //RS3("Post_Name")=""
        //RS3("Post_IDNo")=""
        //RS3("Post_SavingsNo")=""
        //RS3("Post_AccountNo")=""
        //RS3("Dept_Id")=RS1("Dept_Id")
        list.Add(new ColumnData("Dept_Id", "C001", true, false, false));
        //RS3("Invoice_Title")=RS1("Donate_Invoice_Title")     RS1 is from donate_web
        string strInvoiceTitle = Util.GetDBValue("Donor", "Invoice_Title", "Donor_Id", strDonorId);
        list.Add(new ColumnData("Invoice_Title", strInvoiceTitle, true, false, false));
        //RS3("Invoice_Pre")=Invoice_Pre
        string strInvoicePre = Util.GetDBValue("Dept", "Invoice_Pre", "Dept_Id", "C001");
        list.Add(new ColumnData("Invoice_Pre", strInvoicePre, true, false, false));
        //RS3("Invoice_No")=Invoice_No
        list.Add(new ColumnData("Invoice_No", GetInvoiceNo(), true, false, false));
        //RS3("Invoice_Print")="0"
        list.Add(new ColumnData("Invoice_Print", "0", true, false, false));
        //RS3("Invoice_Print_Add")="0"
        list.Add(new ColumnData("Invoice_Print_Add", "0", true, false, false));
        //RS3("Invoice_Print_Yearly_Add")="0"
        list.Add(new ColumnData("Invoice_Print_Yearly_Add", "0", true, false, false));
        //RS3("Request_Date")=null
        //RS3("Accoun_Bank")=""
        list.Add(new ColumnData("Accoun_Bank", "臺灣銀行", true, false, false));
        //RS3("Accoun_Date")=null
        //RS3("Invoice_type")=RS1("Donate_Invoice_Type")
        // 2014/6/5 修正收據開立判斷
        //list.Add(new ColumnData("Invoice_type", "單次收據", true, false, false));
        list.Add(new ColumnData("Invoice_type", Util.GetDBValue("Donor", "Invoice_Type", "Donor_Id", strDonorId), true, false, false));
        //RS3("Accounting_Title")=""
        //If RS1("Donate_Act_Id")<>"" Then
        //  RS3("Act_id")=RS1("Donate_Act_Id")
        //Else
        //  RS3("Act_id")=null
        //End If
        //RS3("Comment")=""
        //RS3("Invoice_PrintComment")=""
        //RS3("Issue_Type")=""
        //RS3("Issue_Type_Keep")=""
        //RS3("Export")="N"
        list.Add(new ColumnData("Issue_Type", "", true, false, false));
        list.Add(new ColumnData("Issue_Type_Keep", "", true, false, false));
        list.Add(new ColumnData("Export", "N", true, false, false));
        //RS3("Create_Date")=Date()
        list.Add(new ColumnData("Create_Date", Util.GetDBDateTime(), true, false, false));
        //RS3("Create_DateTime")=Now()
        list.Add(new ColumnData("Create_DateTime", Util.GetDBDateTime(), true, false, false));
        //RS3("Create_User")="線上金流"
        list.Add(new ColumnData("Create_User", "線上金流", true, false, false));
        //RS3("Create_IP")=Request.ServerVariables("REMOTE_ADDR")
        list.Add(new ColumnData("Create_IP", Request.ServerVariables["REMOTE_ADDR"], true, false, false));
    }
    

    private string GetInvoiceNo()
    {
        string strRet = "";
        string strInvoicePre = Util.GetDBValue("Dept", "Invoice_Pre", "Dept_Id", "C001");
        // 2014/6/4 修正SQL語法
        //string strSql = "select top 1 invoice_no from donate where invoice_no like @InvoiceNo";
        string strSql = "select top 1 invoice_no from donate where invoice_no like @InvoiceNo order by invoice_no desc";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        // 2014/6/4 修正收據編號定義
        //dict.Add("InvoiceNo", strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd") + "%");
        dict.Add("InvoiceNo", Util.GetDBDateTime().ToString("yyyyMMdd") + "%");
        
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            //取得當日最大流水號
            // 2014/6/4 修正收據編號定義
            //string strSN = dt.Rows[0]["invoice_no"].ToString().Replace(strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd"), "");
            string strSN = dt.Rows[0]["invoice_no"].ToString().Replace(Util.GetDBDateTime().ToString("yyyyMMdd"), "");
            //取得當日最新流水號
            strSN = (Convert.ToInt16(strSN) + 1).ToString("0000");
            //20140515修正InvoiceNo，不加Invoice_Pre
            //strRet = strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd") + strSN;
            strRet = Util.GetDBDateTime().ToString("yyyyMMdd") + strSN;
        }
        else
        {
            //strRet = strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd") + "0001";
            strRet = Util.GetDBDateTime().ToString("yyyyMMdd") + "0001";
        }

        return strRet;
    }
    //-------------------------------------------------------------------------------------------------------------
    private void LoadDropDownList()
    {
    }
    //---------------------------------------------------------------------------
    private void LoadFormData()
    {
    }
    //---------------------------------------------------------------------------
    protected void btnBackPay_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("DonateInfoConfirm.aspx"));

        //Response.Redirect(Util.RedirectByTime("DonateSingle.aspx"));

        //for 單筆奉獻金流使用
        //account:　金額(正整數)
        //account.Value = dtOnce.Rows[0]["奉獻金額"].ToString();
        //orderid:　訂單編號(不得重複,勿超過15碼)
        //orderid.Value = DateTime.Now.Ticks.ToString().Substring(0, 15);
        //param :　提供客戶自行運用(可不填)
        //param.Value = strDonorId;

        //string script = "<script> document.forms[0].action='" + strIpayHttp + "'; document.forms[0].submit(); </script>";
        //Page.ClientScript.RegisterStartupScript(this.GetType(), "postform", script);
    }
    protected void btnConfirm_Click(object sender, EventArgs e)
    {
        //if (dtPeriod.Rows.Count > 0)
        //{
        //    Response.Redirect(Util.RedirectByTime("DonateCreditCard.aspx"));
        //}
        //else
        //{
            //if (dtOnce.Rows.Count > 0)
            //{
            //    //Response.Redirect(Util.RedirectByTime("DonateSingle.aspx"));

            //    //for 單筆奉獻金流使用
            //    //account:　金額(正整數)
            //    account.Value = dtOnce.Rows[0]["奉獻金額"].ToString();
            //    //orderid:　訂單編號(不得重複,勿超過15碼)
            //    orderid.Value = DateTime.Now.Ticks.ToString().Substring(0, 15);
            //    //param :　提供客戶自行運用(可不填)
            //    param.Value = strDonorId;

            //    string script = "<script> document.forms[0].action='http://gateway.demo.linkuswell.com.tw/cardfinance.php'; document.forms[0].submit(); </script>";
            //    Page.ClientScript.RegisterStartupScript(this.GetType(), "postform", script);
            //}
            //else
            //{
                //Response.Redirect(Util.RedirectByTime("DonateOnlineAll.aspx"));
                Response.Redirect(Util.RedirectByTime("donate_index.html"));
            //}
        //}
    }
    //-------------------------------------------------------------------------
    public void ShowCartOnce(DataTable dtSrc)
    {
//        if (dtSrc.Rows.Count > 0)
//        {
//            DataTable dtSum = dtSrc.Clone();
//            //計算合計
//            int iSum = 0;
//            DataRow drSum = dtSum.NewRow();
//            //DataColumn[] keys = new DataColumn[1];
//            //keys[0] = dtSum.Columns[0];
//            //dtSum.PrimaryKey = keys;
//            for (int i = 0; i < dtSrc.Rows.Count; i++)
//            {
//                iSum += Convert.ToInt32(dtSrc.Rows[i]["奉獻金額"].ToString());
//                drSum = dtSum.NewRow();
//                drSum["Uid"] = dtSrc.Rows[i]["Uid"].ToString();
//                drSum["奉獻項目"] = dtSrc.Rows[i]["奉獻項目"].ToString();
//                drSum["奉獻金額"] = dtSrc.Rows[i]["奉獻金額"].ToString();
//                dtSum.Rows.Add(drSum);
//            }
//            drSum = dtSum.NewRow();
//            drSum["奉獻項目"] = "單筆奉獻合計";
//            drSum["奉獻金額"] = iSum.ToString();
//            dtSum.Rows.Add(drSum);

//            //Grid initial
//            //沒有需要特別處理的欄位時
//            NPOGridView npoGridView = new NPOGridView();
//            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
//            npoGridView.dataTable = dtSum;
//            npoGridView.Keys.Add("Uid");
//            npoGridView.DisableColumn.Add("Uid");
//            npoGridView.ShowPage = false;
//            npoGridView.TableID = "ItemOnce";

//            //-------------------------------------------------------------------------
//            NPOGridViewColumn col = new NPOGridViewColumn("奉獻項目");
//            col.ColumnType = NPOColumnType.NormalText;
//            col.ColumnName.Add("奉獻項目");
//            npoGridView.Columns.Add(col);
//            //-------------------------------------------------------------------------
//            col = new NPOGridViewColumn("奉獻金額");
//            col.ColumnType = NPOColumnType.NormalText;
//            col.ColumnName.Add("奉獻金額");
//            npoGridView.Columns.Add(col);
//            //-------------------------------------------------------------------------

//            lblDonateContent.Text += @"                
//                    <span align='center' style='color: blue; font-weight: bold;'>
//                        ※　單　筆　奉　獻　項　目　清　單
//                    </span>
//                " + npoGridView.Render();
//        }
    }
    //-------------------------------------------------------------------------
    public void ShowCartPeriod(DataTable dtSrc)
    {
//        if (dtSrc.Rows.Count > 0)
//        {
//            DataTable dtNew = new DataTable();

//            dtNew.Columns.Add("Uid");
//            dtNew.Columns.Add("奉獻項目");
//            dtNew.Columns.Add("奉獻內容");
//            dtNew.Columns.Add("奉獻金額");

//            for (int i = 0; i < dtSrc.Rows.Count; i++)
//            {
//                DataRow drNew = dtNew.NewRow();
//                drNew["Uid"] = dtSrc.Rows[i]["Uid"].ToString();
//                drNew["奉獻項目"] = dtSrc.Rows[i]["奉獻項目"].ToString();
//                string strContent = "";
//                //月繳
//                if (dtSrc.Rows[i]["繳費方式"].ToString() == "月繳")
//                {
//                    strContent = dtSrc.Rows[i]["開始年"].ToString() + "年";
//                    strContent += dtSrc.Rows[i]["開始月"].ToString() + "月至";
//                    strContent += dtSrc.Rows[i]["結束年"].ToString() + "年";
//                    strContent += dtSrc.Rows[i]["結束月"].ToString() + "月止；";
//                }
//                //季繳
//                if (dtSrc.Rows[i]["繳費方式"].ToString() == "季繳")
//                {
//                    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年第";
//                    //strContent += dtSrc.Rows[i]["開始月"].ToString() + "季至";
//                    //strContent += dtSrc.Rows[i]["結束年"].ToString() + "年第";
//                    //strContent += dtSrc.Rows[i]["結束月"].ToString() + "季止；";
//                }
//                //年繳
//                if (dtSrc.Rows[i]["繳費方式"].ToString() == "年繳")
//                {
//                    strContent = dtSrc.Rows[i]["開始年"].ToString() + "年至";
//                    strContent += dtSrc.Rows[i]["結束年"].ToString() + "年止；";
//                }
//                strContent += dtSrc.Rows[i]["繳費方式"].ToString();
//                drNew["奉獻內容"] = strContent;
//                drNew["奉獻金額"] = dtSrc.Rows[i]["奉獻金額"].ToString();
//                dtNew.Rows.Add(drNew);
//            }

//            //Grid initial
//            //沒有需要特別處理的欄位時
//            NPOGridView npoGridView = new NPOGridView();
//            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
//            npoGridView.dataTable = dtNew;
//            npoGridView.Keys.Add("Uid");
//            npoGridView.DisableColumn.Add("Uid");
//            npoGridView.ShowPage = false;
//            npoGridView.TableID = "ItemPeriod";

//            //-------------------------------------------------------------------------
//            NPOGridViewColumn col = new NPOGridViewColumn("奉獻項目");
//            col.ColumnType = NPOColumnType.NormalText;
//            col.ColumnName.Add("奉獻項目");
//            npoGridView.Columns.Add(col);
//            //-------------------------------------------------------------------------
//            col = new NPOGridViewColumn("奉獻內容");
//            col.ColumnType = NPOColumnType.NormalText;
//            col.ColumnName.Add("奉獻內容");
//            npoGridView.Columns.Add(col);
//            //-------------------------------------------------------------------------
//            col = new NPOGridViewColumn("奉獻金額");
//            col.ColumnType = NPOColumnType.NormalText;
//            col.ColumnName.Add("奉獻金額");
//            npoGridView.Columns.Add(col);
//            //-------------------------------------------------------------------------

//            lblDonateContent.Text += @"                
//                    <span align='center' style='color: Blue; font-weight: bold;'>
//                        ※　定　期　定　額　奉　獻　項　目　清　單
//                    </span>
//                " + npoGridView.Render();
//        }
    }
}