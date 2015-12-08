using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Web.SessionState;
using System.Net.Mail;
using System.Xml;
using System.Text;
using System.Security.Cryptography;//Md5加密

public partial class Online_DonateAtOnce : System.Web.UI.Page
{

    string strCreateDateTime = "";      //for get new donor_id
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
    string strInvoiceType;
    string strIsAnonymous;
    string strCity;
    string strArea;
    string strIsAbroad;
    string strOverseasCountry;
    string strOverseasAddress;
    string strInvoice_Title;
    string strReport = ""; //紀錄異常的
    
    // 2015/6/11 增加國泰世華銀行信用卡授權網站變數
    string strIpayHttp = System.Configuration.ConfigurationManager.AppSettings["cathay_epos"];
    string StrCubkey = System.Configuration.ConfigurationManager.AppSettings["cubkey"];
    string StrStoreId = System.Configuration.ConfigurationManager.AppSettings["storeid2"];

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
            //新增帳號註記 Y="尚未註冊過，新增帳號"
            HFD_AddNew.Value = Request.Params["HFD_AddNew"];
            //已登入帳號的捐款人編號
            //HFD_DonorID.Value = Request.Params["HFD_DonorID"];
            //已登入帳號註記
            //HFD_LoginOK.Value = Request.Params["HFD_LoginOK"];

            

            Util.FillCityData(ddlLiveCity);
            LoadDropDownList();

            //預設值欄位
            //同意上傳國稅局申報
            //HFD_IsFdc.Value = "否";
            //居住地區
            HFD_LiveRegion.Value = "台灣";
            HFD_Receipt.Value = "年度證明";   //單次收據
            //訂閱紙本月刊
            //HFD_Subscribe.Value = "否";  //原"是"，2015/2/26改成"否"
            //是否刊登在捐款人名錄
            //HFD_Anony.Value = "不刊登";

            //載入下拉式選單
            LoadFormData();

            //if (HFD_LoginOK.Value == "ok")
            //{
            //    //已登入帳號
            //    txtEmail.ReadOnly = true;
            //    txtEmail.CssClass = "readonly";
            //    LoadData(HFD_DonorID.Value);
            //    Util.SetDdlIndex(ddlLiveArea, HFD_LiveArea.Value);
            //}
            //else
            //{
                //txtEmail.ReadOnly = false;
                //txtEmail.CssClass = "";
                //txtEmail.Text = Request.Params["txtAccount"];
            //}

        }

    }

    //-------------------------------------------------------------------------------------------------------------
    protected void btnCheckOut_Click(object sender, EventArgs e)
    {
        //保留備用此處登入
        //bool bolRet = CheckLogin(txtEmail.Text, txtPassword.Text);

        //儲存捐款人資料
        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();
        SetColumnData(list);

        //if (HFD_LoginOK.Value == "ok")
        //{
        //    //已登入帳號
        //    string strSql = Util.CreateUpdateCommand("Donor", list, dict);
        //    NpoDB.ExecuteSQLS(strSql, dict);
        //}
        //else
        //{
            //新增帳號與不註冊直接奉獻

            string strSql = Util.CreateInsertCommand("Donor", list, dict);
            strSql += " SELECT SCOPE_IDENTITY() ";
            HFD_DonorID.Value = NpoDB.ExecuteSQLScalar(strSql, dict);

            //if (HFD_AddNew.Value == "Y")
            //{
            //    HFD_LoginOK.Value = "ok";
            //}
        //}

        //20150611 Post XML內容
        //storeid
        storeid.Value = StrStoreId;
        //ordernumber:　訂單編號(不得重複,勿超過15碼)
        ordernumber.Value = DateTime.Now.Ticks.ToString().Substring(0, 15);
        Session["OrderNumber"] = ordernumber.Value;
        //cubkey
        cubkey.Value = StrCubkey;
        //cavalue 使用Md5加密
        MD5 md5 = MD5.Create();
        string str = StrStoreId + ordernumber.Value + txtAmount.Text + StrCubkey;
        cavalue.Value = StrToMd5String(str);
        //postxml內容
        XmlDocument doc = new XmlDocument();
        XmlDeclaration declaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
        doc.AppendChild(declaration);
        //建立根節點
        XmlElement root = doc.CreateElement("MERCHANTXML");
        doc.AppendChild(root);
        //建立子節點
        XmlElement CAVALUE = doc.CreateElement("CAVALUE");
        XmlElement ORDERINFO = doc.CreateElement("ORDERINFO");
        //加入至MERCHANTXML節點底下
        root.AppendChild(CAVALUE);
        root.AppendChild(ORDERINFO);
        //建立子節點
        XmlElement STOREID = doc.CreateElement("STOREID");
        XmlElement ORDERNUMBER = doc.CreateElement("ORDERNUMBER");
        XmlElement AMOUNT = doc.CreateElement("AMOUNT");
        //加入至ORDERINFO子節點底下
        ORDERINFO.AppendChild(STOREID);
        ORDERINFO.AppendChild(ORDERNUMBER);
        ORDERINFO.AppendChild(AMOUNT);
        //建立文字
        CAVALUE.InnerText = cavalue.Value;
        STOREID.InnerText = StrStoreId;
        ORDERNUMBER.InnerText = ordernumber.Value;
        AMOUNT.InnerText = txtAmount.Text;

        strRqXML.Value = doc.OuterXml;

        //新增交易資料(DONATE_WEB)
        InsertDonateWeb();

        ClientScript.RegisterStartupScript(this.GetType(), "postform", "<script>document.forms[0].action='" + strIpayHttp + "'; document.forms[0].submit();</script>");
        //ClientScript.RegisterStartupScript(this.GetType(), "postform", "<script>document.forms[0].action='DonateMobile.aspx'; document.forms[0].submit();</script>");
        //Response.Redirect(Util.RedirectByTime("DonateInfoConfirm.aspx"));
    }
    public static string StrToMd5String(string str)
    {
        MD5 md = new MD5CryptoServiceProvider();

        byte[] b = md.ComputeHash(System.Text.Encoding.Default.GetBytes(str));

        var md5str = BitConverter.ToString(b).Replace("-", "");

        return md5str;
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
        string strSql = @"select * from Donor where Donor_Id=@Donor_Id and DeleteDate is null ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", HFD_DonorID.Value);

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            DataRow dr = dt.Rows[0];
            strDonorName = dr["Donor_Name"].ToString();
            strSex = dr["Sex"].ToString();
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
        //新增交易資料(DONATE_WEB)
        list.Add(new ColumnData("Donate_Type", "單次捐款", true, false, false));
        //list.Add(new ColumnData("Donate_Type", strDonate_Type, true, false, false));
        list.Add(new ColumnData("od_sob", Session["OrderNumber"].ToString(), true, false, false));
        list.Add(new ColumnData("Dept_Id", "C001", true, false, false));
        list.Add(new ColumnData("Donor_Id", HFD_DonorID.Value, true, false, false));
        list.Add(new ColumnData("Donate_CreateDate", Util.GetDBDateTime().ToShortDateString(), true, false, false));
        list.Add(new ColumnData("Donate_CreateDateTime", Util.GetDBDateTime(), true, false, false));
        string strRemoteAddr = Request.ServerVariables["REMOTE_ADDR"];
        if (strRemoteAddr == "::1")
        {
            strRemoteAddr = "127.0.0.1";
        }
        list.Add(new ColumnData("Donate_CreateIP", strRemoteAddr, true, false, false));
        list.Add(new ColumnData("Donate_Amount", txtAmount.Text, true, false, false));
        list.Add(new ColumnData("Donate_DonorName", strDonorName, true, false, false));
        list.Add(new ColumnData("Donate_Sex", strSex, true, false, false));
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
        list.Add(new ColumnData("Donate_Purpose", "經常費", true, false, false));
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
    //-------------------------------------------------------------------------------------------------------------
    private void LoadDropDownList()
    {
        ddlSection.Items.Clear();
        ddlSection.Items.Add(new ListItem("", ""));
        ddlSection.Items.Add(new ListItem("一", "一"));
        ddlSection.Items.Add(new ListItem("二", "二"));
        ddlSection.Items.Add(new ListItem("三", "三"));
        ddlSection.Items.Add(new ListItem("四", "四"));
        ddlSection.Items.Add(new ListItem("五", "五"));
        ddlSection.Items.Add(new ListItem("六", "六"));
        ddlSection.Items.Add(new ListItem("七", "七"));
        ddlSection.Items.Add(new ListItem("八", "八"));
        ddlSection.Items.Add(new ListItem("九", "九"));
        ddlSection.Items.Add(new ListItem("十", "十"));

        Util.FillDropDownList(ddlOverseasCountry, Util.GetDataTable("CaseCode", "GroupName", "國家", "CaseID", ""), "CaseName", "CaseName", false);
        ddlOverseasCountry.Items.Insert(0, new ListItem("", ""));
        ddlOverseasCountry.SelectedIndex = 0;
    }

    //-------------------------------------------------------------------------------------------------------------
    private void LoadFormData()
    {
        //性別
        List<ControlData> list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_Gender", "rdoSex1", "男", "男", false, HFD_Gender.Value));
        list.Add(new ControlData("Radio", "RDO_Gender", "rdoSex2", "女", "女", false, HFD_Gender.Value));
        lblGender.Text = HtmlUtil.RenderControl(list);
        //寄發收據
        //list = new List<ControlData>();
        //list.Add(new ControlData("Radio", "RDO_Receipt", "rdoReceipt1", "年度證明", "年度證明", false, HFD_Receipt.Value));
        //list.Add(new ControlData("Radio", "RDO_Receipt", "rdoReceipt2", "單次收據", "逐次寄發", false, HFD_Receipt.Value));
        //list.Add(new ControlData("Radio", "RDO_Receipt", "rdoReceipt3", "不寄", "不寄", false, HFD_Receipt.Value));
        //lblReceipt.Text = HtmlUtil.RenderControl(list);
        //徵信錄
        //list = new List<ControlData>();
        //list.Add(new ControlData("Radio", "RDO_Anony", "rdoAnony1", "不刊登", "不刊登", false, HFD_Anony.Value));
        //list.Add(new ControlData("Radio", "RDO_Anony", "rdoAnony2", "刊登", "刊登", false, HFD_Anony.Value));
        //lblAnony.Text = HtmlUtil.RenderControl(list);
        //訂閱月刊
        //list = new List<ControlData>();
        //list.Add(new ControlData("Radio", "RDO_Subscribe", "RdoSubscribe1", "是", "是", false, HFD_Subscribe.Value));
        //list.Add(new ControlData("Radio", "RDO_Subscribe", "RdoSubscribe2", "否", "否", false, HFD_Subscribe.Value));
        //lblSubscribe.Text = HtmlUtil.RenderControl(list);
        //是否上傳給國稅局
        //list = new List<ControlData>();
        //list.Add(new ControlData("Radio", "RDO_IsFdc", "rdoIsFdc1", "是", "是", false, HFD_IsFdc.Value));
        //list.Add(new ControlData("Radio", "RDO_IsFdc", "rdoIsFdc2", "否", "否", false, HFD_IsFdc.Value));
        //lblIsFdc.Text = HtmlUtil.RenderControl(list);
        //居住地區
        list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_LiveRegion", "RdoLiveRegion1", "台灣", "台灣", false, HFD_LiveRegion.Value));
        list.Add(new ControlData("Radio", "RDO_LiveRegion", "RdoLiveRegion2", "海外地區", "海外地區", false, HFD_LiveRegion.Value));
        lblLiveRegion.Text = HtmlUtil.RenderControl(list);
    }

    //-------------------------------------------------------------------------------------------------------------
    private void SetColumnData(List<ColumnData> list)
    {

        //已登入帳號取得資料準備比對
        //if (HFD_LoginOK.Value == "ok")
        //{
        //    Dictionary<string, object> dict = new Dictionary<string, object>();
        //    string strSql = @"select * from DONOR where Donor_Id=@DonorID and DeleteDate is null ";
        //    dict.Add("DonorID", HFD_DonorID.Value);
        //    DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        //    if (dt.Rows.Count > 0)
        //    {
        //        DataRow dr = dt.Rows[0];
        //        gDonorID = dr["Donor_Id"].ToString(); //捐款人編號
        //        gName = dr["Donor_Name"].ToString(); //奉獻者姓名
        //        Session["DonorName"] = dr["Donor_Name"].ToString();
        //        gSex = dr["Sex"].ToString();
        //        gEmail = dr["Email"].ToString();
        //        gTitle = dr["Invoice_Title"].ToString(); //收據抬頭

        //        //mod by geo 20131226 for 寄發收據always否
        //        //gType = dr["Invoice_Type"].ToString() == "Y" ? "是" : "否"; // 寄發收據
        //        gType = dr["Invoice_Type"].ToString(); // 寄發收據

        //        gIsAnonymous = HFD_Anony.Value; //徵信錄
        //        gBirthday = Util.DateTime2String(dr["Birthday"], DateType.yyyyMMdd, EmptyType.ReturnEmpty);
        //        gIDNo = dr["IDNo"].ToString(); //身份證字號
        //        gAbroad = dr["IsAbroad"].ToString() == "Y" ? "海外地區" : "台灣";//居住地區
        //        gCity = Util.GetCityName(dr["City"].ToString());  //聯絡地址-縣市
        //        gArea = Util.GetAreaName(dr["Area"].ToString());  //鄉鎮 Area
        //        gZipcode = dr["ZipCode"].ToString();
        //        gStreet = dr["Street"].ToString();//街道
        //        gSection = dr["Section"].ToString(); //段
        //        gLane = dr["Lane"].ToString();  //巷
        //        gAlley = dr["Alley"].ToString();//弄
        //        gHouseNo = dr["HouseNo"].ToString(); //號
        //        gHouseNoSub = dr["HouseNoSub"].ToString(); //號之
        //        gFloor = dr["Floor"].ToString();  //樓
        //        gFloorSub = dr["FloorSub"].ToString();  //樓之
        //        gRoom = dr["Room"].ToString(); //室
        //        gAttn = dr["Attn"].ToString(); //Attn
        //        gOverseasCountry = dr["OverseasCountry"].ToString();//海外
        //        gOverseasAddress = dr["OverseasAddress"].ToString();//海外地址
        //        gPhoneCode = dr["Tel_Office_Loc"].ToString();//電話區碼
        //        gCorpPhone = dr["Tel_Office"].ToString();//電話
        //        gtExt = dr["Tel_Office_Ext"].ToString();//分機
        //        gCellPhone = dr["Cellular_Phone"].ToString(); //手機
        //        gFaxCode = dr["Fax_Loc"].ToString(); //傳真區碼
        //        gFax = dr["Fax"].ToString(); //傳真
        //        gSendNews = dr["IsSendNews"].ToString() == "Y" ? "是" : "否";//訂閱月刊
        //        gIsFdc = dr["IsFdc"].ToString() == "Y" ? "是" : "否"; //是否上傳給國稅局
        //        gToGoodTV = dr["ToGoodTV"].ToString(); // 想要對GOOD TV說的話            
        //        gChangePwd = dr["Donor_Pwd"].ToString(); // 密碼            
        //    }
        //}
        //**************************************************************

        //第一個欄位給 update 用
        //捐款人編號
        list.Add(new ColumnData("Donor_Id", txtDonorID.Text, false, false, true));
        //奉獻者姓名
        list.Add(new ColumnData("Donor_Name", txtDonorName.Text, true, true, false));
        //性別
        list.Add(new ColumnData("Sex", Util.GetControlValue("RDO_Gender"), true, true, false));
        //稱謂
        list.Add(new ColumnData("Title", Util.GetControlValue("RDO_Gender") == "男" ? "先生" : "小姐", true, true, false));
        //身份別
        list.Add(new ColumnData("Donor_Type", "個人", true, false, false));
        //Email
        list.Add(new ColumnData("Email", txtEmail.Text, true, false, false));
        //密碼  若有新增帳號與輸入密碼則存密碼，反之則否
        if (HFD_AddNew.Value == "Y" && !String.IsNullOrEmpty(txtPassword.Text))
        {
            list.Add(new ColumnData("Donor_Pwd", txtPassword.Text, true, false, false));
        }
        //收據抬頭
        //if (txtTitle.Text.Trim() != "")
        //{
        //    list.Add(new ColumnData("Invoice_Title", txtTitle.Text, true, true, false));
        //}
        //else //帶入捐款人姓名
        //{
            list.Add(new ColumnData("Invoice_Title", txtDonorName.Text, true, true, false));
        //}
        //寄發收據
        //mod by geo 20131226 for 寄發收據always否
        //list.Add(new ColumnData("Invoice_Type", Util.GetControlValue("RDO_Receipt") == "" ? "不需要" : "年證明及收據", true, true, false));
        // 2014/6/5 修正收據開立判斷
        //list.Add(new ColumnData("Invoice_Type", Util.GetControlValue("RDO_Receipt") == "是" ? "單次收據" : "不寄", true, true, false));

        //20150519 捐款人記錄當年度第一次收據寄送方式
        //string Sql = "select Year(ISNULL(IsInvoiceTypeExist,'')) as IsInvoiceTypeExist from Donor\n";
        //Sql += "where DeleteDate is null and Email=@Email and Donor_Pwd is not NULL \n";

        //Dictionary<string, object> dict_check = new Dictionary<string, object>();
        //dict_check.Add("Email", txtEmail.Text);
        //DataTable dtCheck = NpoDB.GetDataTableS(Sql, dict_check);

        //if (dtCheck.Rows.Count != 0)
        //{
        //    DataRow drcheck = dtCheck.Rows[0];
        //    if (drcheck["IsInvoiceTypeExist"].ToString().Trim() != DateTime.Now.Year.ToString())
        //    {
        //        list.Add(new ColumnData("IsInvoiceTypeExist", Util.GetToday(DateType.yyyyMMdd), true, true, false));
        //        list.Add(new ColumnData("Invoice_Type", Util.GetControlValue("RDO_Receipt"), true, true, false));
        //    }
        //}
        //else
        //{
        //    list.Add(new ColumnData("IsInvoiceTypeExist", Util.GetToday(DateType.yyyyMMdd), true, true, false));
        //    list.Add(new ColumnData("Invoice_Type", Util.GetControlValue("RDO_Receipt"), true, true, false));
        //}
        //20150610 WEB版線上捐款預設值
        list.Add(new ColumnData("Invoice_Type", "年度證明", true, true, false));
        //徵信錄
        //list.Add(new ColumnData("IsAnonymous", Util.GetControlValue("RDO_Anony") == "不刊登" ? "Y" : "N", true, true, false));
        //20150610 WEB版線上捐款預設值
        list.Add(new ColumnData("IsAnonymous", "Y", true, true, false));
        //生日
        //list.Add(new ColumnData("Birthday", Util.DateTime2String(txtBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnNull), true, true, false));
        //身份證字號
        //list.Add(new ColumnData("IDNo", txtIDNo.Text.ToUpper(), true, true, false));
        //居住地區
        list.Add(new ColumnData("IsAbroad", Util.GetControlValue("RDO_LiveRegion") == "台灣" ? "N" : "Y", true, true, false));
        //收據地區
        list.Add(new ColumnData("IsAbroad_Invoice", Util.GetControlValue("RDO_LiveRegion") == "台灣" ? "N" : "Y", true, true, false));

        string strLiveAddress = "";
        string strLiveAddress2 = "";
        if (Util.GetControlValue("RDO_LiveRegion") == "台灣")
        {
            //聯絡地址-縣市
            list.Add(new ColumnData("City", ddlLiveCity.SelectedValue, true, true, false));
            //聯絡地址-鄉鎮
            list.Add(new ColumnData("Area", HFD_LiveArea.Value, true, true, false));
            //聯絡地址-郵遞區號
            list.Add(new ColumnData("ZipCode", HFD_LiveArea.Value, true, true, false));
            //聯絡地址-地址
            list.Add(new ColumnData("Street", txtLiveStreet.Text.Trim(), true, true, false));
            list.Add(new ColumnData("Section", ddlSection.SelectedValue, true, true, false));
            list.Add(new ColumnData("Lane", txtLane.Text.Trim(), true, true, false));
            list.Add(new ColumnData("Alley", txtAlley.Text.Trim(), true, true, false));
            list.Add(new ColumnData("HouseNo", txtHouseNo.Text.Trim(), true, true, false));
            //list.Add(new ColumnData("HouseNoSub", txtHouseNoSub.Text.Trim(), true, true, false));
            list.Add(new ColumnData("Floor", txtFloor.Text.Trim(), true, true, false));
            list.Add(new ColumnData("FloorSub", txtFloorSub.Text.Trim(), true, true, false));
            list.Add(new ColumnData("Room", txtRoom.Text.Trim(), true, true, false));
            list.Add(new ColumnData("Attn", txtAttn.Text.Trim(), true, true, false));

            if (txtLiveStreet.Text.Trim() != "")
            {
                strLiveAddress += txtLiveStreet.Text.Trim();
            }
            if (ddlSection.SelectedIndex != 0)
            {
                strLiveAddress += ddlSection.SelectedValue + "段";
            }
            if (txtLane.Text.Trim() != "")
            {
                strLiveAddress += txtLane.Text.Trim() + "巷";
            }
            if (txtAlley.Text.Trim() != "")
            {
                strLiveAddress += txtAlley.Text.Trim() + "弄";
            }
            if (txtHouseNo.Text.Trim() != "")
            {
                strLiveAddress += txtHouseNo.Text.Trim();
            }
            //if (txtHouseNoSub.Text.Trim() != "")
            //{
            //    strLiveAddress += "之" + txtHouseNoSub.Text.Trim();
            //}
            if (txtHouseNo.Text.Trim() != "")
            {
                strLiveAddress += "號";
            }
            if (txtFloor.Text.Trim() != "")
            {
                strLiveAddress += txtFloor.Text.Trim() + "樓";
            }
            if (txtFloorSub.Text.Trim() != "")
            {
                strLiveAddress += "之" + txtFloorSub.Text.Trim();
            }
            if (txtRoom.Text.Trim() != "")
            {
                strLiveAddress += " - " + txtRoom.Text.Trim() + "室";
            }
            //marked out by Samuel for removing duplicated attn shown in both invoice and mailing address 2014/08/11 
            //if (txtAttn.Text.Trim() != "")
            //{
            // strLiveAddress += "(" + txtAttn.Text.Trim() + ")";
            //}
            strLiveAddress2 = Util.GetControlValue("RDO_LiveRegion") + ddlLiveCity.SelectedItem.Text + Util.GetAreaName(HFD_LiveArea.Value) + strLiveAddress;
            //if (Util.GetControlValue("RDO_Receipt") != "")
            //{
                //聯絡地址-縣市
                list.Add(new ColumnData("Invoice_City", ddlLiveCity.SelectedValue, true, true, false));
                //聯絡地址-鄉鎮
                list.Add(new ColumnData("Invoice_Area", HFD_LiveArea.Value, true, true, false));
                //聯絡地址-郵遞區號
                list.Add(new ColumnData("Invoice_ZipCode", HFD_LiveArea.Value, true, true, false));
                //聯絡地址-地址
                list.Add(new ColumnData("Invoice_Address", strLiveAddress, true, true, false));
                list.Add(new ColumnData("Invoice_Street", txtLiveStreet.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Section", ddlSection.SelectedValue, true, true, false));
                list.Add(new ColumnData("Invoice_Lane", txtLane.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Alley", txtAlley.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_HouseNo", txtHouseNo.Text.Trim(), true, true, false));
                //list.Add(new ColumnData("Invoice_HouseNoSub", txtHouseNoSub.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Floor", txtFloor.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_FloorSub", txtFloorSub.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Room", txtRoom.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Attn", txtAttn.Text.Trim(), true, true, false));
            //}
            //訂購紙本月刊 20140623修改
            //if (Util.GetControlValue("RDO_Subscribe") == "是")
            //{
            //    if (strLiveAddress == "") //若未填地址，自動存不訂購月刊
            //    {
            //        list.Add(new ColumnData("IsSendNews", "N", true, true, false));
            //        list.Add(new ColumnData("IsSendNewsNum", "NULL", true, true, false));
            //    }
            //    else
            //    {
            //        list.Add(new ColumnData("IsSendNews", "Y", true, true, false));
            //        list.Add(new ColumnData("IsSendNewsNum", "1", true, true, false));
            //    }
            //}
            //else
            //{
                list.Add(new ColumnData("IsSendNews", "N", true, true, false));
                list.Add(new ColumnData("IsSendNewsNum", "NULL", true, true, false));
            //}
        }
        else
        {
            //海外地區
            //國家/省/城市/區
            list.Add(new ColumnData("OverseasCountry", ddlOverseasCountry.SelectedValue, true, true, false));
            //地址
            list.Add(new ColumnData("OverseasAddress", txtOverseasAddress.Text.Trim(), true, true, false));
            //strLiveAddress = txtOverseasCountry.Text.Trim() + " " + txtOverseasAddress.Text.Trim();
            strLiveAddress = txtOverseasAddress.Text.Trim() + " " + ddlOverseasCountry.SelectedValue;
            strLiveAddress2 = Util.GetControlValue("RDO_LiveRegion") + strLiveAddress;
            //if (Util.GetControlValue("RDO_Receipt") != "")
            //{
                //收據國家/省/城市/區
                list.Add(new ColumnData("Invoice_OverseasCountry", ddlOverseasCountry.SelectedValue, true, true, false));
                //收據地址
                list.Add(new ColumnData("Invoice_OverseasAddress", txtOverseasAddress.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Address", strLiveAddress, true, true, false));
            //}
            //訂購紙本月刊 20140623修改
            list.Add(new ColumnData("IsSendNews", "N", true, true, false));
            list.Add(new ColumnData("IsSendNewsNum", "NULL", true, true, false));
        }
        list.Add(new ColumnData("Address", strLiveAddress, true, true, false));
        //住家電話
        //list.Add(new ColumnData("Tel_Home", txtLocalPhone.Text, true, true, false));
        //電話區碼 Tel_Office_Loc
        list.Add(new ColumnData("Tel_Office_Loc", txtPhoneCode.Text, true, true, false));
        //電話
        list.Add(new ColumnData("Tel_Office", txtCorpPhone.Text, true, true, false));
        //電話分機 Tel_Office_Ext
        list.Add(new ColumnData("Tel_Office_Ext", txtExt.Text, true, true, false));
        //手機電話
        list.Add(new ColumnData("Cellular_Phone", txtCellPhone.Text, true, true, false));
        //傳真區碼
        //list.Add(new ColumnData("Fax_Loc", txtFaxCode.Text, true, true, false));
        //傳真
        //list.Add(new ColumnData("Fax", txtFax.Text, true, true, false));
        //訂閱月刊
        //list.Add(new ColumnData("IsSendNews", Util.GetControlValue("RDO_Subscribe") == "是" ? "Y" : "N", true, true, false));
        //2014/6/11 User需求暫時取消，2015/1/15恢復
        //是否上傳給國稅局
        //list.Add(new ColumnData("IsFdc", Util.GetControlValue("RDO_IsFdc") == "是" ? "Y" : "N", true, true, false));
        //20150610 WEB版線上捐款預設值
        list.Add(new ColumnData("IsFdc", "N", true, true, false));
        //想要對GOOD TV說的話
        //改成有填寫「想對GOOD TV說的話」欄位就寄送Email與內送Email
        //if (!String.IsNullOrEmpty(txtToGoodTV.Text))
        //{
        //    list.Add(new ColumnData("ToGoodTV", gToGoodTV + "\r\n" + Util.GetToday(DateType.yyyyMMddHHmmss) + " 在線上奉獻填寫如下：\r\n" + txtToGoodTV.Text + "\r\n", true, true, false));
        //}
        //20130930 Add by GoodTV Tanya
        //隸屬組織
        list.Add(new ColumnData("Dept_Id", "C001", true, true, false));
        //是否為讀者
        list.Add(new ColumnData("IsMember", "N", true, true, false));

        //新增日期
        strCreateDateTime = Util.GetToday(DateType.yyyyMMddHHmmss); //for get new donor_id
        list.Add(new ColumnData("Create_Date", Util.GetToday(DateType.yyyyMMdd), true, false, false));
        list.Add(new ColumnData("Create_DateTime", strCreateDateTime, true, false, false));
        list.Add(new ColumnData("Create_User", "線上金流", true, false, false));
        //最後更新日期
        list.Add(new ColumnData("LastUpdate_Date", Util.GetToday(DateType.yyyyMMdd), false, true, false));
        list.Add(new ColumnData("LastUpdate_DateTime", Util.GetToday(DateType.yyyyMMddHHmmss), false, true, false));
        list.Add(new ColumnData("LastUpdate_User", "線上金流", false, true, false));

        //寄送Email用 for 新增帳號與已登入帳號用
        string Mailhead = "";
        string MailFrom = "";
        string MailSubject = "";
        string MailBody = "";
        string MailTo = "";
        string strReportToDonations = ""; //紀錄異常的(內部Email使用)

        //if (HFD_LoginOK.Value == "ok")
        //{

        //    //已登入帳號 與原資料比對是否有異動
        //    string strAddr = "";
        //    string strAddr2 = "";

        //    if (gStreet != "")
        //    {
        //        strAddr += gStreet;
        //    }
        //    if (gSection != "")
        //    {
        //        strAddr += gSection + "段";
        //    }
        //    if (gLane != "")
        //    {
        //        strAddr += gLane + "巷";
        //    }
        //    if (gAlley != "")
        //    {
        //        strAddr += gAlley + "弄";
        //    }
        //    if (gHouseNo != "")
        //    {
        //        strAddr += gHouseNo;
        //    }
        //    if (gHouseNoSub != "")
        //    {
        //        strAddr += "之" + gHouseNoSub;
        //    }
        //    if (gHouseNo != "")
        //    {
        //        strAddr += "號";
        //    }
        //    if (gFloor != "")
        //    {
        //        strAddr += gFloor + "樓"; ;
        //    }
        //    if (gFloorSub != "")
        //    {
        //        strAddr += "之" + gFloorSub;
        //    }
        //    if (gRoom != "")
        //    {
        //        strAddr += " -" + gRoom + "室";
        //    }
        //    //marked out by Samuel for removing duplicated attn shown in both invoice and mailing address 2014/08/11 
        //    //if (gAttn != "")
        //    //{
        //    //strAddr += "(" + gAttn + ")";
        //    //}
        //    // 2014/6/12 更正地址邏輯(台灣與海外地區) by 詩儀
        //    //strAddr = (Util.GetControlValue("RDO_LiveRegion") == "台灣") ? strAddr : txtOverseasCountry.Text.Trim() + " " + txtOverseasAddress.Text.Trim();
        //    if (Util.GetControlValue("RDO_LiveRegion") != gAbroad) //海外地區<==>台灣
        //    {
        //        if (gAbroad == "台灣")
        //        {
        //            strAddr2 = gAbroad + gCity + gArea + strAddr;  //異動前地址
        //        }
        //        else
        //        {
        //            // 2014/7/7 看到應該是address+Country
        //            //strAddr2 = gAbroad + gOverseasCountry + " " + ;//異動前海外地址
        //            strAddr2 = gAbroad + gOverseasAddress + " " + gOverseasCountry;//異動前海外地址
        //        }
        //    }
        //    else  //台灣<==>台灣 or 海外<==>海外
        //    {
        //        if (Util.GetControlValue("RDO_LiveRegion") == "台灣")
        //        {
        //            strAddr2 = gAbroad + gCity + gArea + strAddr;  //異動前地址
        //        }
        //        else
        //        {
        //            // 2014/7/7 看到應該是address+Country
        //            //strAddr2 = gAbroad + gOverseasCountry + " " + gOverseasAddress;//異動前海外地址
        //            strAddr2 = gAbroad + gOverseasAddress + " " + gOverseasCountry;//異動前海外地址
        //        }
        //    }

        //    //20140616 修正只顯示更改後的資料。
        //    if (txtDonorName.Text != gName) strReport = strReport + "姓名：" + txtDonorName.Text + "<br/><br/>";
        //    if (Util.GetControlValue("RDO_Gender") != gSex) strReport = strReport + "性別：" + Util.GetControlValue("RDO_Gender") + "<br/><br/>";
        //    //if (txtEmail.Text != gEmail) strReport = strReport + gEmail + "-->更改為：" + txtEmail.Text + "</BR>";
        //    //if (txtTitle.Text != gTitle)
        //    //{
        //    //    strReport = strReport + "收據抬頭：" + txtTitle.Text + "<br/><br/>";
        //    //    strReportToDonations = strReportToDonations + "收據抬頭：" + txtTitle.Text + "<br/><br/>";
        //    //}
        //    //20150519 停止收據寄送異動通知
        //    //if (Util.GetControlValue("RDO_Receipt") != gType)
        //    //{
        //    //    strReport = strReport + "寄發收據：" + Util.GetControlValue("RDO_Receipt") + "<br/><br/>";
        //    //    strReportToDonations = strReportToDonations + "寄發收據：" + Util.GetControlValue("RDO_Receipt") + "<br/><br/>";
        //    //}
        //    //if (Util.GetControlValue("RDO_Anony") != gIsAnonymous) strReport = strReport + "徵信錄：" + Util.GetControlValue("RDO_Anony") + "<br/><br/>";
        //    //if (Util.DateTime2String(txtBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnEmpty) != gBirthday) strReport = strReport + "生日：" + Util.DateTime2String(txtBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnNull) + "<br/><br/>";
        //    //if (txtIDNo.Text != gIDNo) strReport = strReport + "身分證字號：" + txtIDNo.Text + "<br/><br/>";

        //    if (strLiveAddress2 != strAddr2)
        //    {
        //        strReport = strReport + "地址：" + strLiveAddress2 + "<br/><br/>";
        //        strReportToDonations = strReportToDonations + "地址：" + strLiveAddress2 + "<br/><br/>";
        //    }

        //    // 2014/7/8 電話區碼+電話+電話分機
        //    if ((txtPhoneCode.Text + txtCorpPhone.Text + txtExt.Text) != (gPhoneCode + gCorpPhone + gtExt))
        //    {
        //        strReport = strReport + "電話：" + txtPhoneCode.Text + "-" + txtCorpPhone.Text + (txtExt.Text != "" ? " 分機：" + txtExt.Text : "") + "<br/><br/>";
        //        strReportToDonations = strReportToDonations + "電話：" + txtPhoneCode.Text + "-" + txtCorpPhone.Text + (txtExt.Text != "" ? " 分機：" + txtExt.Text : "") + "<br/><br/>";
        //    }
        //    //if (txtPhoneCode.Text != gPhoneCode)
        //    //{
        //    //    strReport = strReport + "電話區碼：" + txtPhoneCode.Text + "<BR><BR>";
        //    //    strReportToDonations = strReportToDonations + "電話區碼：" + txtPhoneCode.Text + "<BR><BR>";
        //    //}
        //    //if (txtCorpPhone.Text != gCorpPhone)
        //    //{
        //    //    strReport = strReport + "電話：(" + txtPhoneCode.Text + ")" + txtCorpPhone.Text + "<BR><BR>";
        //    //    strReportToDonations = strReportToDonations + "電話：(" + txtPhoneCode.Text + ")" + txtCorpPhone.Text + "<BR><BR>";
        //    //}
        //    //if (txtExt.Text != gtExt)
        //    //{
        //    //    strReport = strReport + "電話分機：" + txtExt.Text + "<BR><BR>";
        //    //    strReportToDonations = strReportToDonations + "電話分機：" + txtExt.Text + "<BR><BR>";
        //    //}
        //    if (txtCellPhone.Text != gCellPhone)
        //    {
        //        strReport = strReport + "手機：" + txtCellPhone.Text + "<br/><br/>";
        //        strReportToDonations = strReportToDonations + "手機：" + txtCellPhone.Text + "<br/><br/>";
        //    }
        //    // 2014/7/8 傳真區碼+傳真
        //    //if (txtFaxCode.Text + txtFax.Text != gFaxCode + gFax) strReport = strReport + "傳真：" + txtFaxCode.Text + "-" + txtFax.Text + "<br/><br/>";
        //    //if (txtFaxCode.Text != gFaxCode) strReport = strReport + "傳真區碼：" + txtFaxCode.Text + "<BR><BR>";
        //    //if (txtFax.Text != gFax) strReport = strReport + "傳真：" + txtFax.Text + "<BR><BR>"; if (Util.GetControlValue("RDO_Subscribe") != gSendNews) strReport = strReport + "訂閱月刊：" + Util.GetControlValue("RDO_Subscribe") + "<BR><BR>";
        //    //if (Util.GetControlValue("RDO_IsFdc") != gIsFdc) strReport = strReport + "上傳給國稅局：" + gIsFdc + "-->更改為：" + Util.GetControlValue("RDO_IsFdc") + "<BR>";
        //    //if (Util.GetControlValue("RDO_IsFdc") != gIsFdc) strReport = strReport + "上傳給國稅局：" + Util.GetControlValue("RDO_IsFdc") + "<br/><br/>";
        //    if (!String.IsNullOrEmpty(txtPassword.Text) && txtPassword.Text != gChangePwd) strReport += "您的密碼已變更為：" + txtPassword.Text + "<br/><br/>";
        //    //if (!String.IsNullOrEmpty(txtToGoodTV.Text))
        //    //{
        //    //    strReport += "想要對GOOD TV說的話：<br/>" + txtToGoodTV.Text.Replace("\r\n", "<br/>") + "<br/><br/>";
        //    //}

        //    if (strReport != "")  //若有異動資料則發信
        //    {

        //        Mailhead = "親愛的 " + txtDonorName.Text + " 平安！<br/><br/>以下是您在GOODTV網站的異動內容如下，敬請比對查閱，謝謝您。<br/><br/>";
        //        strReport += "<br/>";
        //        strReport += "若您有任何資料異動相關問題敬請來電: 02-8024-3911 查詢<br/>";
        //        strReport += "願神祝福您!<br/>";
        //        strReport += "GOODTV捐款服務組敬上<br/>";

        //        SendEMailObject MailObject = new SendEMailObject();
        //        //MailObject.SmtpServer = SmtpServer;
        //        MailSubject = "GOODTV異動通知信";
        //        MailBody = Mailhead + strReport;
        //        MailTo = txtEmail.Text;

        //        MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
        //        string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, MailBody);
        //        if (MailObject.ErrorCode != 0)
        //        {
        //            ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        //        }

        //    }

        //    if (strReportToDonations != "")  //若有指定欄位異動則發內部信
        //    {

        //        Mailhead = txtDonorName.Text + " 異動個人資料如下：<br/><br/>";

        //        SendEMailObject MailObject = new SendEMailObject();
        //        //MailObject.SmtpServer = SmtpServer;
        //        MailSubject = "GOODTV線上奉獻_異動通知";
        //        MailBody = Mailhead + strReportToDonations;
        //        MailTo = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];

        //        MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
        //        string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, MailBody);
        //        if (MailObject.ErrorCode != 0)
        //        {
        //            ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        //        }

        //    }

        //}
        //else 
        if (HFD_AddNew.Value == "Y")
        {
            //新增帳號
            string strBody = "";
            strBody = "親愛的 " + txtDonorName.Text + " 平安！<br/><br/>您的GOODTV帳號已於 " + strCreateDateTime + " 註冊成功，以下是您的帳號及密碼, 請妥善保管喔!<br/>";
            strBody += "您的帳號：" + txtEmail.Text + "<br/>";
            strBody += "您的密碼：" + txtPassword.Text + "<br/>";
            strBody += "願神祝福您!<br/><br/>GOODTV捐款服務組敬上";

            SendEMailObject MailObject = new SendEMailObject();
            //MailObject.SmtpServer = SmtpServer;
            MailSubject = " GOODTV線上奉獻註冊成功通知信";
            MailBody = strBody;
            MailTo = txtEmail.Text;
            MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
            string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, strBody);
            if (MailObject.ErrorCode != 0)
            {
                ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
            }

        }

        //if (!String.IsNullOrEmpty(txtToGoodTV.Text))  //若有想要對GOODTV說的話則發內部信
        //{

        //    Mailhead = "奉獻天使：" + txtDonorName.Text + "<br/>Email: " + txtEmail.Text + "<br/>";
        //    if (!String.IsNullOrEmpty(txtCellPhone.Text))
        //    {
        //        Mailhead += "手機：" + txtCellPhone.Text + "<br/>";
        //    }
        //    if (!String.IsNullOrEmpty(txtCorpPhone.Text))
        //    {
        //        Mailhead += "電話：" + txtPhoneCode.Text + "-" + txtCorpPhone.Text + (txtExt.Text != "" ? " 分機：" + txtExt.Text : "") + "<br/>";
        //    }
        //    Mailhead += "<br/>想要對GOODTV說的話如下：<br/>";

        //    SendEMailObject MailObject = new SendEMailObject();
        //    //MailObject.SmtpServer = SmtpServer;
        //    MailSubject = "GOODTV線上奉獻_奉獻天使想要對GOODTV說的話_通知";
        //    MailBody = Mailhead + txtToGoodTV.Text.Replace("\r\n", "<br/>") + "<br/>";
        //    MailTo = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];

        //    MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
        //    string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, MailBody);
        //    if (MailObject.ErrorCode != 0)
        //    {
        //        ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        //    }

        //}

    }

}
