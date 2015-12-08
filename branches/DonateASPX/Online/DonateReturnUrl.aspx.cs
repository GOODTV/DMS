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
    string MailFrom = "";
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
    string strInvoice_No = "";
    string strInvoice_Pre = "";
   
    protected void Page_Load(object sender, EventArgs e)
    {
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
            //LoadFormData();
            if (Session["DonorName"] != null)
            {
                lblTitle.Text = Session["DonorName"].ToString() + lblTitle.Text;
            }
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

            lblTitle.Text += "<br/><br/>交易序號：" + Request.Form["orderid"].ToString();
            if (strStatus == "0")
            {
                this.btnBackPay.Visible = false;

                lblTitle.Text += "<br/><br/>授權狀態：成功";
                Session["sta"] = "OK";

                InsertIEPay();
                InsertDonate();
                dtOnce.Rows[0].Delete();
                Session["InsertPeriod"] = null;                           
            }
            else
            {
                this.btnBackPay.Visible = true;

                lblTitle.Text += "<br/><br/>授權狀態：失敗" + strStatus;
                Session["sta"] = "";
            }
            lblTitle.Text += "<br/><br/>您的奉獻金額：" + strAmount + " 元";
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
            //else
            //{
             //   lblTitle.Text += "<br/><br/>付款方式：" + strPayType;
           // }
            //lblTitle.Text += "<br/>param：" + Request.Form["param"].ToString();

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
        NpoDB.ExecuteSQLS(strSql, dict);
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
        list.Add(new ColumnData("payformat", Request.Form["payformat"].ToString(), true, false, false));
        //authdate
        list.Add(new ColumnData("authdate", Request.Form["authdate"].ToString(), true, false, false));
        //param
        list.Add(new ColumnData("param", Request.Form["param"].ToString(), true, false, false));
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
            dict1.Add("DonorID", Session["DonorID"].ToString());
            dt = NpoDB.GetDataTableS(strSql, dict1);
            if (dt.Rows.Count != 0)
            {
                dr = dt.Rows[0];
                gdonor_name = dr["donor_name"].ToString();
                gEmail = dr["Email"].ToString();
                gTitle = dr["Invoice_Title"].ToString();
                gType = dr["Invoice_Type"].ToString() == "Y" ? "是" : "否";
                gIsAnonymous = dr["IsAnonymous"].ToString() == "Y" ? "是" : "否";
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
                strAddr += gHouseNo + "號";
            }
            if (gHouseNoSub != "")
            {
                strAddr += "之" + gHouseNoSub;
            }
            if (gFloor != "")
            {
                strAddr += "," + gFloor + "樓";
            }
            if (gFloorSub != "")
            {
                strAddr += "之" + gFloorSub;
            }
            if (gRoom != "")
            {
                strAddr += "," + gRoom + "室";
            }
            string Sname = "";
            strAddr2 = gAbroad + gCity + gArea + strAddr;
            Sname = Util.GetDBValue("donor", "Donor_Name", "donor_ID", strDonorId);
            strInvoice_No = Util.GetDBValue("DONATE", "Invoice_No", "od_sob", strOrderId);
            strInvoice_Pre = Util.GetDBValue("DONATE", "Invoice_Pre", "od_sob", strOrderId); 
            
            strBody = "親愛的 " + Session["DonorName"].ToString() + " 感謝您對GOODTV的奉獻支持，以下是您本次奉獻資料：<BR>";
            strBody += "您的帳號: " + gEmail + " <BR>";
            strBody += "奉獻編號：" + strInvoice_No + "  <BR>";
            strBody += "收據抬頭：" + gTitle + "<BR>";
            strBody += "奉獻金額 新台幣" + strAmount + "<BR>";
            strBody += "收據地址：" + strAddr2 + " <BR>";
            strBody += "收據寄發：" + gType + " <BR>";
            strBody += "徵信錄：" + gIsAnonymous + " (您的姓名將不刊登於徵信錄) <BR>";
            strBody += "信用卡授權時間：" + Util.GetDBDateTime() + "<P>";
            strBody += "若您有任何捐款相關問題敬請來電: 02-8024-3911 查詢<BR>";
            strBody += "願神大大賜福給您!<BR>";
            strBody += "GOODTV捐款管理組敬上<BR>";
            SendEMailObject MailObject = new SendEMailObject();
            MailObject.SmtpServer = SmtpServer;
            MailSubject = " GOODTV線上奉獻確認信";
            MailBody = strBody;
            MailTo =  gEmail;
            MailFrom = MailFrom = System.Configuration.ConfigurationSettings.AppSettings["MailFrom"];

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
        list.Add(new ColumnData("Donate_Date", Util.GetDBDateTime(), true, false, false));
        //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
        //RS3("Donate_Amt")=Request.Form("amount")
        list.Add(new ColumnData("Donate_Amt", account, true, false, false));
        //RS3("Donate_AmtD")=Request.Form("amount")
        list.Add(new ColumnData("Donate_AmtD", account, true, false, false));
        //RS3("Donate_Fee")=Donate_Fee
        //RS3("Donate_FeeD")=Donate_Fee
        //RS3("Donate_Accou")=Donate_Accou
        //RS3("Donate_AccouD")=Donate_Accou
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
        list.Add(new ColumnData("Donate_Purpose", purpose, true, false, false));
        //RS3("Donate_Purpose_Type")="D"
        list.Add(new ColumnData("Donate_Purpose_Type", "D", true, false, false));
        //RS3("Donate_Type")="單次捐款"
        list.Add(new ColumnData("Donate_Type", "單次捐款", true, false, false));
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
        //RS3("Accoun_Date")=null
        //RS3("Invoice_type")=RS1("Donate_Invoice_Type")  
        list.Add(new ColumnData("Invoice_type", "單次收據", true, false, false));
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
        string strSql = "select top 1 invoice_no from donate where invoice_no like @InvoiceNo";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("InvoiceNo", strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd") + "%");
        
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            //取得當日最大流水號
            string strSN = dt.Rows[0]["invoice_no"].ToString().Replace(strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd"), "");
            //取得當日最新流水號
            strSN = (Convert.ToInt16(strSN) + 1).ToString("0000");
            strRet = strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd") + strSN;
        }
        else
        {
            strRet = strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd") + "0001";
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
        //Response.Redirect(Util.RedirectByTime("DonateSingle.aspx"));

        //for 單筆奉獻金流使用
        //account:　金額(正整數)
        account.Value = dtOnce.Rows[0]["奉獻金額"].ToString();
        //orderid:　訂單編號(不得重複,勿超過15碼)
        orderid.Value = DateTime.Now.Ticks.ToString().Substring(0, 15);
        //param :　提供客戶自行運用(可不填)
        param.Value = strDonorId;

        string script = "<script> document.forms[0].action='http://gateway.demo.linkuswell.com.tw/cardfinance.php'; document.forms[0].submit(); </script>";
        Page.ClientScript.RegisterStartupScript(this.GetType(), "postform", script);
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
                Response.Redirect(Util.RedirectByTime("DonateOnlineAll.aspx"));
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