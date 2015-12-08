using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateReturnUrl : System.Web.UI.Page
{

    // 2014/11/3 增加Log
    static log4net.ILog logger = log4net.LogManager.GetLogger("RollingLogFileAppender");

    //DataTable dtOnce = new DataTable();
    //DataTable dtPeriod = new DataTable();
    string strOrderId;
    string strStatus;
    string strAmount;
    string strPayType;
    string strDonorId;
    //string Mailhead = "";
    string SmtpServer = System.Configuration.ConfigurationManager.AppSettings["MailServer"];
    string MailSubject = "";
    string MailBody = "";
    string MailTo = "";
    string MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
    string strBody = "";
    string gdonor_name = "";
    //string strSql = "";
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
    string strIpayHttp = System.Configuration.ConfigurationManager.AppSettings["ipay_http"];
    string StrEmailToDonations = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];
    string strPayTypeName = "";
    string strPayformat = "";
    string strAuthdate = "";
    string strErrcode = "";
    string InsertIEPay_flag = "";

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {

            logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", DonateReturnUrl.aspx.cs 程式內部: Page_Load 進入點開始, URL:" + HttpUtility.UrlDecode(Request.Form.ToString()));

            /* 測試用
            if (Request.Params["orderid"] == null)*/
            if (Request.Form["orderid"] == null)
            {
                logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", DonateReturnUrl.aspx.cs 程式內部: Post未帶參數");
                Response.Write("請勿直接輸入網址！");
                Response.End();
            }

            /**/
            strOrderId = Request.Form["orderid"].ToString();
            strStatus = Request.Form["status"].ToString();
            strAmount = Request.Form["account"].ToString();
            strPayType = Request.Form["paytype"].ToString();
            strDonorId = Request.Form["param"].ToString();
            strPayformat = Request.Form["payformat"].ToString();
            strAuthdate = Request.Form["authdate"].ToString();
            strErrcode = Request.Form["errcode"].ToString();

            HFD_UID.Value = Request.Form["param"].ToString();
            HFD_OrderId.Value = Request.Form["orderid"].ToString();
            /* 測試用
            strOrderId = Request.Params["orderid"].ToString();
            strStatus = Request.Params["status"].ToString();
            strAmount = Request.Params["account"].ToString();
            strPayType = Request.Params["paytype"].ToString();
            strDonorId = Request.Params["param"].ToString();
            strPayformat = Request.Params["payformat"].ToString();
            strAuthdate = Request.Params["authdate"].ToString();
            */
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

            lblTitle.Text += "<br/><br/>交易序號 Transaction No：" + strOrderId;
            if (strStatus == "0")
            {
                //this.btnBackPay.Visible = false;
                //20140710 更改授權狀態顯示內容
                switch (strPayType)
                {
                    case "1":
                        lblTitle.Text += "<br/><br/><font color='red'>授權狀態 Pledge Status：成功 Success</font>"; //信用卡
                        break;
                    case "8":
                        lblTitle.Text += "<br/><br/><font color='red'>授權狀態 Pledge Status：成功 Success</font>"; //WEB ATM
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
                lblTitle.Text += "<br/><br/>您的奉獻金額 Total Amount：NT $" + strAmount + " 元";

                logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", OrderId:" + strOrderId + ", DonateReturnUrl.aspx.cs 程式內部: 授權成功");

            }
            else
            {
                //this.btnBackPay.Visible = true;

                lblTitle.Text += "<br/><br/><font color='red'>授權狀態 Pledge Status：失敗 declined (" + strStatus + ")</font>";
                logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", OrderId:" + strOrderId + ", DonateReturnUrl.aspx.cs 程式內部: 授權失敗 授權狀態:" + strStatus);
            }

            // 2014/6/6 改變判斷邏輯方式
            switch (strPayType)
            {
                case "1":
                    lblTitle.Text += "<br/><br/>付款方式 Payment：信用卡 Credit Card";
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
            //lblName.Text = gdonor_name;
            //lblTitle.Text = gdonor_name + lblTitle.Text;

            //if (Session["IsRefresh"] == null)
            //{

            InsertIEPay();
            if (InsertIEPay_flag == "ok")
            {
                if (strStatus == "0")
                {
                    logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", OrderId:" + strOrderId + ", DonateReturnUrl.aspx.cs 程式內部: InsertIEPay()");

                    // 2014/7/9 只有線上信用卡轉帳授權成功才算是確定繳款
                    if (strPayType == "1")
                    {
                        InsertDonate();
                        logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", OrderId:" + strOrderId + ", DonateReturnUrl.aspx.cs 程式內部: InsertDonate()");
                        // 2014/7/3 增加重新統計捐款人捐款資訊
                        UpdateDonor();
                        // 2014/9/24 增加授權成功內部通知
                        InternalSendMail(true);
                    }
                    //dtOnce.Rows[0].Delete();
                    //Session["InsertPeriod"] = null;
                    // 2014/6/6 增加完成線上奉獻就清除Session
                    //Session.Clear();
                }
                else
                {
                    // 2014/8/11 增加奉獻天使授權失敗的內部通知信
                    InternalSendMail(false);

                }
                //}
                //Session["IsRefresh"] = "yes";
            }
            else
            {

                logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", OrderId:" + strOrderId + ", DonateReturnUrl.aspx.cs 程式內部: 使用者網頁重新整理");
                Util.ShowMsg("請勿重新整理網頁！");
            }
            if (strStatus == "0")//授權成功才填問卷
            {
                //this.btnQuestion.Visible = true;
                HFD_Question.Value = "Y";
            }
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

            DataTable dt = null;
            DataRow dr = null;
            Dictionary<string, object> dict1 = new Dictionary<string, object>();
            string strSql = @"select *
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
                ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
            }

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
            try
            {
                NpoDB.ExecuteSQLS(strSql, dict);
                InsertIEPay_flag = "ok";
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
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
        //errcode 2014/12/4 新增欄位
        list.Add(new ColumnData("errcode", strErrcode, true, false, false));
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

        // 2014/11/3 原從Session變數取得修改成從資料庫取得
        Dictionary<string, object> dict_web = new Dictionary<string, object>();
        dict_web.Add("OrderId", strOrderId);
        DataTable dt_web = new DataTable();
        string strSql_web = @"
            SELECT od_sob,[Donate_Purpose],[Donate_Amount]
            FROM [Donate_Web]
            WHERE [IsDelete] is null and [od_sob] = @OrderId
        ";
        dt_web = NpoDB.GetDataTableS(strSql_web, dict_web);

        //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
        for (int i = 0; i < dt_web.Rows.Count; i++)
        {
            Dictionary<string, object> dict = new Dictionary<string, object>();
            List<ColumnData> list = new List<ColumnData>();

            string purpose = dt_web.Rows[i]["Donate_Purpose"].ToString();
            string account = dt_web.Rows[i]["Donate_Amount"].ToString();
            
            SetColumnDataDonate(list,purpose,account);
            string strSql = "";
            strSql = Util.CreateInsertCommand("DONATE", list, dict);
            
            try
            {
                NpoDB.ExecuteSQLS(strSql, dict);
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }

        }

        //發送mail*****************************************

        DataTable dt = null;
        DataRow dr = null;
        Dictionary<string, object> dict1 = new Dictionary<string, object>();
        string strSql2 = @"select *
            from DONOR 
            where Donor_Id=@DonorID
            ";
        dict1.Add("DonorID", strDonorId);
        dt = NpoDB.GetDataTableS(strSql2, dict1);
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

        strBody = "親愛的 " + gdonor_name + " 平安！Dear " + gdonor_name + ",<BR><BR>";
        strBody += "感謝您對GOODTV的奉獻支持，以下是您本次奉獻資料： Thank you for your monetary gift:<BR>";
        strBody += "您的帳號 Your account(email)： " + gEmail + " <BR>";
        strBody += "收據編號 Receipt No.：" + strInvoice_No + "  <BR>";
        strBody += "收據抬頭 Name of Receipt：" + gTitle + "<BR>";
        strBody += "奉獻金額 Amount：NT$" + strAmount + "元<BR>";
        strBody += "收據地址 Mailing address：" + strAddr2 + " <BR>";
        if (gType == "是")
            strBody += "收據寄發 Send receipt：是 Yes <BR>";
        else
            strBody += "收據寄發 Send receipt：否 No <BR>";
        //strBody += "徵信錄 Credit list：" + gIsAnonymous + " <BR><BR>";
        if (gIsAnonymous == "刊登")
            strBody += "徵信錄 Credit list：刊登 Yes <BR>";
        else
            strBody += "徵信錄 Credit list：不刊登 No <BR>";
        //strBody += "信用卡授權時間：" + Util.GetDBDateTime() + "<P>";
        strBody += "若您有任何捐款相關問題敬請來電: 02-8024-3911 查詢 Any question you may call us at: 02-8024-3911.<BR>";
        strBody += "願神大大賜福給您! May God bless you!<BR>";
        strBody += "GOODTV捐款服務組敬上 Yours sincerely, GOODTV<BR>";
        SendEMailObject MailObject = new SendEMailObject();
        MailObject.SmtpServer = SmtpServer;
        MailSubject = " GOODTV線上奉獻確認信Donation Confirmation";
        MailBody = strBody;
        MailTo = gEmail;

        string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, strBody);

        if (MailObject.ErrorCode != 0)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        }

    }

    //-------------------------------------------------------------------------------------------------------------
    private void SetColumnDataDonate(List<ColumnData> list, string purpose, string account)
    {
      
        //新增捐款紀錄
        list.Add(new ColumnData("od_sob", strOrderId, true, false, false));
        list.Add(new ColumnData("Donor_Id", strDonorId, true, false, false));
        //2014/6/4 改正捐款日期不含時間
        list.Add(new ColumnData("Donate_Date", Util.GetDBDateTime().ToShortDateString(), true, false, false));
        //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
        list.Add(new ColumnData("Donate_Amt", account, true, false, false));
        list.Add(new ColumnData("Donate_AmtD", account, true, false, false));
        list.Add(new ColumnData("Donate_Fee", "0", true, false, false));
        list.Add(new ColumnData("Donate_FeeD", "0", true, false, false));
        list.Add(new ColumnData("Donate_Accou", account, true, false, false));
        list.Add(new ColumnData("Donate_AccouD", account, true, false, false));
        list.Add(new ColumnData("Donate_AmtM", "0", true, false, false));
        list.Add(new ColumnData("Donate_FeeM", "0", true, false, false));
        list.Add(new ColumnData("Donate_AccouM", "0", true, false, false));
        list.Add(new ColumnData("Donate_AmtA", "0", true, false, false));
        list.Add(new ColumnData("Donate_FeeA", "0", true, false, false));
        list.Add(new ColumnData("Donate_AccouA", "0", true, false, false));
        list.Add(new ColumnData("Donate_AmtS", "0", true, false, false));
        list.Add(new ColumnData("Donate_RateS", "0", true, false, false));
        list.Add(new ColumnData("Donate_FeeS", "0", true, false, false));
        list.Add(new ColumnData("Donate_AccouS", "0", true, false, false));
        list.Add(new ColumnData("Donate_Payment", "網路信用卡", true, false, false));
        //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
        //list.Add(new ColumnData("Donate_Purpose", purpose, true, false, false));
        list.Add(new ColumnData("Donate_Purpose", "經常費", true, false, false));
        list.Add(new ColumnData("Donate_Purpose_Type", "D", true, false, false));
        // 2014/6/5 修正收據開立判斷
        list.Add(new ColumnData("Donate_Type", "單次捐款", true, false, false));
        //list.Add(new ColumnData("Donate_Type", strDonate_Type, true, false, false));
        list.Add(new ColumnData("IsBeductible", "N", true, false, false));
        list.Add(new ColumnData("Donate_Amt2", "0", true, false, false));
        list.Add(new ColumnData("Dept_Id", "C001", true, false, false));
        string strInvoiceTitle = Util.GetDBValue("Donor", "Invoice_Title", "Donor_Id", strDonorId);
        list.Add(new ColumnData("Invoice_Title", strInvoiceTitle, true, false, false));
        string strInvoicePre = Util.GetDBValue("Dept", "Invoice_Pre", "Dept_Id", "C001");
        list.Add(new ColumnData("Invoice_Pre", strInvoicePre, true, false, false));
        list.Add(new ColumnData("Invoice_No", GetInvoiceNo(), true, false, false));
        list.Add(new ColumnData("Invoice_Print", "0", true, false, false));
        list.Add(new ColumnData("Invoice_Print_Add", "0", true, false, false));
        list.Add(new ColumnData("Invoice_Print_Yearly_Add", "0", true, false, false));
        list.Add(new ColumnData("Accoun_Bank", "臺灣銀行", true, false, false));
        // 2014/6/5 修正收據開立判斷
        //list.Add(new ColumnData("Invoice_type", "單次收據", true, false, false));
        list.Add(new ColumnData("Invoice_type", Util.GetDBValue("Donor", "Invoice_Type", "Donor_Id", strDonorId), true, false, false));
        list.Add(new ColumnData("Issue_Type", "", true, false, false));
        list.Add(new ColumnData("Issue_Type_Keep", "", true, false, false));
        list.Add(new ColumnData("Export", "N", true, false, false));
        list.Add(new ColumnData("Create_Date", Util.GetDBDateTime(), true, false, false));
        list.Add(new ColumnData("Create_DateTime", Util.GetDBDateTime(), true, false, false));
        list.Add(new ColumnData("Create_User", "線上金流", true, false, false));
        string strRemoteAddr = Request.ServerVariables["REMOTE_ADDR"];
        if (strRemoteAddr == "::1")
        {
            strRemoteAddr = "127.0.0.1";
        }
        list.Add(new ColumnData("Create_IP", strRemoteAddr, true, false, false));
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

    protected void btnBackPay_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("DonateInfoConfirm.aspx"));
    }

    protected void btnConfirm_Click(object sender, EventArgs e)
    {
        //Response.Redirect(Util.RedirectByTime("DonateOnlineAll.aspx"));
        ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>window.open('DonateOnlineAll.aspx','_self');</script>");
    }
    protected void btnQuestion_Click(object sender, EventArgs e)
    {
        //ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>if (confirm('感謝您的奉獻！您的寶貴意見將可以協助我們提供更優質的捐款服務。您願意現在填寫奉獻問卷調查嗎？')){window.open('OnlineQuestionnaire.aspx?OrderNumber=" + strOrderId + "&Donor_Id=" + strDonorId + "&Device=1', 'window','HEIGHT=500,WIDTH=540,top=50,left=50,toolbar=yes,scrollbars=yes,resizable=yes'); }else {alert('非常感謝您的奉獻！'); };</script>");
        //ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>if (confirm('感謝您的奉獻！您的寶貴意見將可以協助我們提供更優質的捐款服務。您願意現在填寫奉獻問卷調查嗎？')){window.open('OnlineQuestionnaire.aspx?OrderNumber=" + HFD_OrderId.Value + "&Donor_Id=" + HFD_UID.Value + "&Device=1','_blank'); }else {alert('非常感謝您的奉獻！'); };</script>");
        //HFD_Question.Value = "Y";
        //ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"感謝您的奉獻！您的寶貴意見將可以協助我們提供更優質的捐款服務。\");</script>");
    }
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        //20151112有填才新增問卷
        if (this.DonateMotive1.Checked || this.DonateMotive2.Checked || this.DonateMotive3.Checked || this.DonateMotive4.Checked || this.DonateMotive5.Checked || this.WatchMode1.Checked || this.WatchMode2.Checked || this.WatchMode3.Checked || this.WatchMode4.Checked || this.WatchMode5.Checked || this.WatchMode6.Checked || this.WatchMode7.Checked || this.WatchMode8.Checked || this.WatchMode9.Checked || txtToGoodTV.Text != "")
        {
            Dictionary<string, object> dict = new Dictionary<string, object>();
            string strSql = "insert into  Donate_OnlineQuestion\n";
            strSql += "( OrderNumber, Donor_Id, DonateMotive1, DonateMotive2, DonateMotive3, DonateMotive4, DonateMotive5, \n";
            strSql += " WatchMode1, WatchMode2, WatchMode3, WatchMode4, WatchMode5, WatchMode6, WatchMode7, WatchMode8, WatchMode9, Device, ToGOODTV,Create_Date,Create_DateTime,DonateWay) values\n";
            strSql += "(@OrderNumber,@Donor_Id,@DonateMotive1,@DonateMotive2,@DonateMotive3,@DonateMotive4,@DonateMotive5, \n";
            strSql += "@WatchMode1,@WatchMode2,@WatchMode3,@WatchMode4,@WatchMode5,@WatchMode6,@WatchMode7,@WatchMode8,@WatchMode9,@Device,@ToGOODTV,@Create_Date,@Create_DateTime,@DonateWay)\n";

            dict.Add("OrderNumber", HFD_OrderId.Value);
            dict.Add("Donor_Id", HFD_UID.Value);
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
            dict.Add("DonateWay", "線上單筆奉獻");

            NpoDB.ExecuteSQLS(strSql, dict);
            ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"感謝您撥出寶貴的時間來完成問卷調查。願神祝福您！\");</script>");
        }
        HFD_Question.Value = "";
    }
}
