using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Xml;
using System.Collections.Generic;
using System.Security.Cryptography;//Md5加密
using System.IO;
using System.Text;

public partial class Online_DonateToPledge : System.Web.UI.Page
{
    string strReturnUrl_OK = System.Configuration.ConfigurationManager.AppSettings["ReturnUrl_OK"];
    string strReturnUrl_Fail = System.Configuration.ConfigurationManager.AppSettings["ReturnUrl_Fail"];
    string strOTTReturnUrl_OK = System.Configuration.ConfigurationManager.AppSettings["OTTReturnUrl_OK"];
    string strOTTReturnUrl_Fail = System.Configuration.ConfigurationManager.AppSettings["OTTReturnUrl_Fail"];
    string StrCubkey = System.Configuration.ConfigurationManager.AppSettings["cubkey"];
    string DNS = System.Configuration.ConfigurationManager.AppSettings["dns"];
    string SmtpServer = System.Configuration.ConfigurationManager.AppSettings["MailServer"];
    string MailSubject = "";
    string MailBody = "";
    string MailTo = "";
    string MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
    string strBody = "";
    string StrEmailToDonations = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];
    string gdonor_name = "";
    string gEmail = "";
    string gOTTAddress = "";
    string InsertIEPay_flag = "";
    //取消Session改為變數 2015/11/19
    string strStoreID = "";
    string strOrderNumber = "";
    string strAmount = "";
    string strAuthStatus = "";
    string strAuthCode = "";
    string strDonorId = "";
    string strOTT = "";
    // 2015/11/18 增加Log
    static log4net.ILog logger = log4net.LogManager.GetLogger("RollingLogFileAppender");

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", DonateToPledge.aspx.cs 程式內部: Page_Load 進入點開始, URL:" + HttpUtility.UrlDecode(Request.Form.ToString()));
            //定義一個接收POST的XML
            //MD5 md5 = MD5.Create();
            //string str = "010970018" + "635696303232151" + "100" + "0000" + "0" + "f5600f43ff792a7a25274e5f210e48f2";

            //XmlDocument doc = new XmlDocument();
            //XmlDeclaration declaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            //doc.AppendChild(declaration);
            ////建立根節點
            //XmlElement root = doc.CreateElement("CUBXML");
            //doc.AppendChild(root);
            ////建立子節點
            //XmlElement CAVALUE = doc.CreateElement("CAVALUE");
            //XmlElement ORDERINFO = doc.CreateElement("ORDERINFO");
            //XmlElement AUTHINFO = doc.CreateElement("AUTHINFO");
            ////加入至CUBXML節點底下
            //root.AppendChild(CAVALUE);
            //root.AppendChild(ORDERINFO);
            //root.AppendChild(AUTHINFO);
            ////建立子節點
            //XmlElement STOREID = doc.CreateElement("STOREID");
            //XmlElement ORDERNUMBER = doc.CreateElement("ORDERNUMBER");
            //XmlElement AMOUNT = doc.CreateElement("AMOUNT");

            //XmlElement AUTHSTATUS = doc.CreateElement("AUTHSTATUS");
            //XmlElement AUTHCODE = doc.CreateElement("AUTHCODE");
            //XmlElement AUTHTIME = doc.CreateElement("AUTHTIME");
            //XmlElement AUTHMSG = doc.CreateElement("AUTHMSG");
            ////加入至ORDERINFO子節點底下
            //ORDERINFO.AppendChild(STOREID);
            //ORDERINFO.AppendChild(ORDERNUMBER);
            //ORDERINFO.AppendChild(AMOUNT);
            ////加入至AUTHINFO子節點底下
            //AUTHINFO.AppendChild(AUTHSTATUS);
            //AUTHINFO.AppendChild(AUTHCODE);
            //AUTHINFO.AppendChild(AUTHTIME);
            //AUTHINFO.AppendChild(AUTHMSG);
            ////建立文字
            //CAVALUE.InnerText = StrToMd5String(str);
            //STOREID.InnerText = "010970018";
            //ORDERNUMBER.InnerText = "635696303232151";
            //AMOUNT.InnerText = "100";

            //AUTHSTATUS.InnerText = "0000";  //成功
            //AUTHCODE.InnerText = "0";
            //AUTHTIME.InnerText = "20150611163711";
            //AUTHMSG.InnerText = "TEST";
            //接收國泰世華的授權結果XML
            string strRsXml = Request["strRsXml"].ToString();//doc.OuterXml;
            logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", 授權結果XML:" + strRsXml + ", DonateToPledge.aspx.cs 程式內部: 接收國泰世華的授權結果XML");
            XmlDocument XmlDoc = new XmlDocument();
            XmlDoc.LoadXml(strRsXml);
            XmlNode storeid = XmlDoc.SelectSingleNode("/CUBXML/ORDERINFO/STOREID");
            strStoreID = storeid.InnerText;
            XmlNode ordernumber = XmlDoc.SelectSingleNode("/CUBXML/ORDERINFO/ORDERNUMBER");
            strOrderNumber = ordernumber.InnerText;
            //Session["OrderNumber"] = strOrderNumber;
            XmlNode amount = XmlDoc.SelectSingleNode("/CUBXML/ORDERINFO/AMOUNT");
            strAmount = amount.InnerText;
            //Session["Amount"] = strAmount;
            XmlNode authstatus = XmlDoc.SelectSingleNode("/CUBXML/AUTHINFO/AUTHSTATUS");
            strAuthStatus = authstatus.InnerText;
            //Session["AuthStatus"] = strAuthStatus;
            XmlNode authcode = XmlDoc.SelectSingleNode("/CUBXML/AUTHINFO/AUTHCODE");
            strAuthCode = authcode.InnerText;
            //Response.Write("AUTHCODE：" + authcode.InnerText + "<br/ >");
            XmlNode authtime = XmlDoc.SelectSingleNode("/CUBXML/AUTHINFO/AUTHTIME");
            //Response.Write("AUTHTIME：" + authtime.InnerText + "<br/ >");
            XmlNode authmsg = XmlDoc.SelectSingleNode("/CUBXML/AUTHINFO/AUTHMSG");
            //Response.Write("AUTHMSG：" + authmsg.InnerText + "<br/ >");
            logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", OrderId:" + ordernumber.InnerText + ", Amount:" + amount.InnerText + ", AuthStatus:" + authstatus.InnerText + ", Authtime:" + authtime.InnerText + ", DonateToPledge.aspx.cs 程式內部: 接收國泰世華的授權結果XML資訊");
            
            try
            {
                // 取得Donor_Id
                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("OrderId", ordernumber.InnerText);
                DataTable dt = new DataTable();
                string strSql = @"  SELECT od_sob,[Donor_Id],ISNULL(OTT,'') as [OTT]
                                FROM [Donate_Web]
                                WHERE [IsDelete] is null and [od_sob] = @OrderId";
                dt = NpoDB.GetDataTableS(strSql, dict);
                if (dt.Rows.Count > 0)
                {
                    strDonorId = dt.Rows[0]["Donor_Id"].ToString();
                    strOTT = dt.Rows[0]["OTT"].ToString();
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }

            InsertIEPay();
            if (InsertIEPay_flag == "ok")
            {
                logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", OrderId:" + strOrderNumber + ", DonateToPledge.aspx.cs 程式內部: InsertIEPay()");
            }
            string strReturnUrl = "";
            if (strAuthStatus == "0000")//成功
            {
                logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", OrderId:" + strOrderNumber + ", DonateToPledge.aspx.cs 程式內部: 授權成功");

                if (strOTT == "Y")
                {
                    strReturnUrl = strOTTReturnUrl_OK;
                }
                else
                {
                    strReturnUrl = strReturnUrl_OK;
                }
                InsertDonate();
                logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", OrderId:" + strOrderNumber + ", DonateToPledge.aspx.cs 程式內部: InsertDonate()");
                // 增加重新統計捐款人捐款資訊
                UpdateDonor();
                // 2015/9/30 增加授權成功內部通知
                InternalSendMail(true);
            }
            else
            {
                logger.Info("IP:" + Request.ServerVariables["REMOTE_ADDR"] + ", OrderId:" + strOrderNumber + ", DonateToPledge.aspx.cs 程式內部: 授權失敗 授權狀態:" + strAuthStatus);
                if (strOTT == "Y")
                {
                    strReturnUrl = strOTTReturnUrl_Fail;
                }
                else
                {
                    strReturnUrl = strReturnUrl_Fail;
                }
                // 2015/9/30 增加授權失敗內部通知
                InternalSendMail(false);
            }

            //Check CA Value MD5(storeid+ordernumber+amount+authstatus+authcode+cubkey)
            string strSource = strStoreID + strOrderNumber + strAmount + strAuthStatus + strAuthCode + StrCubkey;
            string strMD5Hash = FormsAuthentication.HashPasswordForStoringInConfigFile(strSource, "MD5");
            StringComparer comparer = StringComparer.OrdinalIgnoreCase;

            //Generate Return Message
            //Generate Return CA Value MD5(domain name + cubkey)
            string strInput = DNS + StrCubkey;
            strMD5Hash = FormsAuthentication.HashPasswordForStoringInConfigFile(strInput, "MD5");
            string strXML = "<?xml version=\'1.0\' encoding=\'UTF-8\'?><MERCHANTXML><CAVALUE>" + strMD5Hash + "</CAVALUE><RETURL>" + strReturnUrl + "</RETURL></MERCHANTXML>";
            Response.ClearHeaders();
            Response.AddHeader("content-type", "text/xml");
            Response.Write(strXML);
            Response.End();
        }
    }
    public static string StrToMd5String(string str)
    {
        MD5 md = new MD5CryptoServiceProvider();

        byte[] b = md.ComputeHash(System.Text.Encoding.Default.GetBytes(str));

        var md5str = BitConverter.ToString(b).Replace("-", "");

        return md5str;
    }
    //-------------------------------------------------------------------------------------------------------------
    private void InsertDonate()
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();

        string account = strAmount;
        SetColumnDataDonate(list, account);
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


        //發送mail給奉獻天使*****************************************

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
            gEmail = dr["Email"].ToString().Trim();
            gOTTAddress = dr["OTT_Address"].ToString();

        }
        if (gEmail != "")
        {
            strBody = "親愛的 " + gdonor_name + " 平安！<BR><BR>";
            strBody += "感謝您對GOODTV的奉獻支持，以下是您本次奉獻資料： <BR>";
            strBody += "您的帳號： " + gEmail + " <BR>";
            strBody += "奉獻金額：NT$" + strAmount + "元<BR>";
            strBody += "聯絡地址：" + gOTTAddress + "<BR><BR>";

            //strBody += "信用卡授權時間：" + Util.GetDBDateTime() + "<P>";
            strBody += "若您有任何捐款相關問題敬請來電: 02-8024-3911 查詢 <BR>";
            strBody += "願神大大賜福給您！<BR>";
            strBody += "GOODTV捐款服務組敬上<BR>";
            SendEMailObject MailObject = new SendEMailObject();
            MailObject.SmtpServer = SmtpServer;
            MailSubject = " GOODTV線上奉獻確認信";
            MailBody = strBody;
            MailTo = gEmail;

            string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, strBody);

            if (MailObject.ErrorCode != 0)
            {
                ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
            }
        }
    }
    //-------------------------------------------------------------------------------------------------------------
    private void SetColumnDataDonate(List<ColumnData> list, string account)
    {

        //新增捐款紀錄
        list.Add(new ColumnData("od_sob", strOrderNumber, true, false, false));
        list.Add(new ColumnData("Donor_Id", strDonorId, true, false, false));
        //2014/6/4 改正捐款日期不含時間
        list.Add(new ColumnData("Donate_Date", Util.GetDBDateTime().ToShortDateString(), true, false, false));
        //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
        list.Add(new ColumnData("Donate_Amt", account, true, false, false));
        list.Add(new ColumnData("Donate_AmtD", "0", true, false, false));
        list.Add(new ColumnData("Donate_Fee", "0", true, false, false));
        list.Add(new ColumnData("Donate_FeeD", "0", true, false, false));
        list.Add(new ColumnData("Donate_Accou", account, true, false, false));
        list.Add(new ColumnData("Donate_AccouD", "0", true, false, false));
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
    //-------------------------------------------------------------------------------------------------------------
    private void InsertIEPay()
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();
        string AuthStatus,Paytype = "";
        if (strAuthStatus == "0000")
        {
            AuthStatus = "0";
        }
        else
        {
            AuthStatus = "1";
        }
        if (strOTT == "Y")
        {
            Paytype = "51";
        }
        else
        {
            Paytype = "41";
        }
        SetColumnDataIEPay(list, AuthStatus, Paytype);
        string strSql = "";
        strSql = Util.CreateInsertCommand("DONATE_IEPAY", list, dict);
        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        String strSql_check = "select orderid from DONATE_IEPAY where orderid = '" + strOrderNumber + "'";
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
    private void SetColumnDataIEPay(List<ColumnData> list, string AuthStatus, string Paytype)
    {
        //orderid
        list.Add(new ColumnData("orderid", strOrderNumber, true, false, false));
        //status
        list.Add(new ColumnData("status", AuthStatus, true, false, false));
        //account
        list.Add(new ColumnData("account", strAmount, true, false, false));
        //paytype
        list.Add(new ColumnData("paytype", Paytype, true, false, false));
        //payformat
        list.Add(new ColumnData("payformat", "1" /*Request.Form["payformat"].ToString()*/, true, false, false));
        //authdate
        list.Add(new ColumnData("authdate", Util.GetToday(DateType.yyyyMMddHHmmss) /*Request.Form["authdate"].ToString()*/, true, false, false));
        //param
        list.Add(new ColumnData("param", strDonorId /*Request.Form["param"].ToString()*/, true, false, false));
        //errcode 2014/12/4 新增欄位
        list.Add(new ColumnData("errcode", strAuthStatus, true, false, false));
        ////authcode
        //list.Add(new ColumnData("authcode", Util.DateTime2String(txtBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnNull), true, true, false));
        ////return_date
        //list.Add(new ColumnData("return_date", txtIDNo.Text, true, true, false));
        ////return_ip
        //list.Add(new ColumnData("return_ip", Util.GetControlValue("RDO_LiveRegion") == "台灣" ? "N" : "Y", true, true, false));
        ////return_url1
        //list.Add(new ColumnData("return_url1", Util.GetToday(DateType.yyyyMMdd), true, false, false));
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
            gOTTAddress = dr["OTT_Address"].ToString();
        }
        string strPhone;
        strPhone = dr["Cellular_Phone"].ToString();

        //發送內部通知mail*****************************************
        string strBody = gdonor_name + " 奉獻資料如下：<BR/><BR/>";
        strBody += "捐款方式：單筆奉獻 <BR/>";
        strBody += "付款方式：信用卡 <BR/>";
        strBody += "訂單編號：" + strOrderNumber + " <BR/>";
        strBody += "付款狀態：" + (YesNo == true ? "授權成功" : "授權失敗") + "<BR/>";
        strBody += String.Format("捐款金額：{0:#,0} 元 <BR/>", strAmount);
        strBody += "E-mail：" + gEmail + "  <BR/>";
        strBody += "電話：" + strPhone + "<BR/>";
        strBody += "聯絡地址：" + gOTTAddress + "<BR/>";

        SendEMailObject MailObject = new SendEMailObject();
        string MailSubject = " GOODTV線上奉獻_新奉獻通知(" + gdonor_name + "_單筆奉獻_" + (YesNo == true ? "授權成功" : "授權失敗") + ")";
        string MailBody = strBody;

        string result = MailObject.SendEmail(StrEmailToDonations, MailFrom, MailSubject, strBody);

        if (MailObject.ErrorCode != 0)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        }

    }
}
