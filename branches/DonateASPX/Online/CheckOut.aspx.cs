using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Web.SessionState;
using System.Net.Mail;
public partial class Online_CheckOut : System.Web.UI.Page
{
    //public static HttpSessionState Session { get { return HttpContext.Current.Session; } }
    DataTable dtOnce = new DataTable();
    DataTable dtPeriod = new DataTable();
    string strCreateDateTime = "";      //for get new donor_id
   //暫存**********************************
    string gDonorID;
    string gName;
    string gSex;
    string gEmail;
    string gTitle;
    string gType;
    string gIsAnonymous;
    string gBirthday;
    string gIDNo;
    string gAbroad;
    string gCity;
    string gArea;
    string gZipcode;
    string gStreet;
    string gSection;
    string gLane;
    string gAlley;
    string gHouseNo;
    string gHouseNoSub;
    string gFloor;
    string gFloorSub;
    string gRoom;
    string gOverseasCountry;
    string gOverseasAddress;
    string gPhoneCode;
    string gCorpPhone;
    string gtExt;
    string gCellPhone;
    string gFaxCode;
    string gFax;
    string gSendNews;
    string gIsFdc;
    string gToGoodTV;
    string strReport =""; //紀錄異常的
    //************************************
    protected void Page_Load(object sender, EventArgs e)
    {
        //有 npoGridView 時才需要
        //Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            Util.FillCityData(ddlLiveCity);
            LoadDropDownList();
            
            //若已登入或輸入過資料，則不show登入
            if (Session["DonorID"] == null)
            {
                HFD_ShowLogin.Value = "Y";
                HFD_showAddNewCheckbox.Value = "Y";
                btnOpenLogin.Visible = true;
                txtEmail.ReadOnly = false;
                txtEmail.CssClass = "";
            }
            else
            {
                HFD_ShowLogin.Value = "N";
                HFD_showAddNewCheckbox.Value = "N";
                HFD_AddNewChecked.Value = "N";                
                btnOpenLogin.Visible = false;
                txtEmail.ReadOnly = true;
                txtEmail.CssClass = "readonly";                
            }
            HFD_IsFdc.Value = "是";
            HFD_LiveRegion.Value = "台灣";
            HFD_Receipt.Value = "是";
            HFD_Subscribe.Value = "是";
            HFD_Anony.Value = "不匿名";
            LoadFormData();

            if (Session["DonorID"] != null)
            {
                //Response.Redirect(Util.RedirectByTime("DonateInfoConfirm.aspx"));
                LoadData(Session["DonorID"].ToString());
            }
            //for 註冊新帳號用
            HFD_AddNewChecked.Value = Util.GetQueryString("AddNewChecked");
        }
        if (Session["ItemOnce"] != null)
        {
            dtOnce = (DataTable)Session["ItemOnce"];
        }
        if (Session["ItemPeriod"] != null)
        {
            dtPeriod = (DataTable)Session["ItemPeriod"];
        }

        int iOnce = dtOnce.Rows.Count;
        int iPeriod= dtPeriod.Rows.Count;
        //lblCart.ForeColor = System.Drawing.Color.DarkRed;
        //lblCart.ForeColor = System.Drawing.Color.White;
        //lblCart.Text = "我的奉獻清單(單筆奉獻：" + iOnce.ToString() + "項；定期定額奉獻：" + iPeriod.ToString() + "項)";
        //if (Session["DonorName"] != null)
        //{
         //   lblCart.Text = lblCart.Text.Replace("我", Session["DonorName"].ToString() + " ");
        //}

        Util.SetDdlIndex(ddlLiveArea, HFD_LiveArea.Value);


    }
    //-------------------------------------------------------------------------------------------------------------
    private void LoadData(string strDonorID)
    {
        string strSql = "";
        DataTable dt = null;
        DataRow dr = null;
        Dictionary<string, object> dict = new Dictionary<string, object>();


        strSql = @"select *
                from DONOR 
                where Donor_Id=@DonorID
                ";
        dict.Add("DonorID", strDonorID);
        dt = NpoDB.GetDataTableS(strSql, dict);
        //資料異常
        if (dt.Rows.Count != 0)
        {
            dr = dt.Rows[0];
            //捐款人編號 Donor_Id
            txtDonorID.Text = dr["Donor_Id"].ToString();
            //奉獻者姓名 Donor_Name
            txtDonorName.Text = dr["Donor_Name"].ToString();
            //性別 Sex
            HFD_Gender.Value = dr["Sex"].ToString();
            //Email Email
            txtEmail.Text = dr["Email"].ToString();
            //密碼  若有勾選新增帳號則存密碼，反之則否
            //收據抬頭 Invoice_Title
            txtTitle.Text = dr["Invoice_Title"].ToString();
            //寄發收據 Invoice_Type
            HFD_Receipt.Value = dr["Invoice_Type"].ToString() == "不需要" ? "是" : "否";
            //徵信錄 IsAnonymous
            if (dr["IsAnonymous"].ToString() == "")
            {
                HFD_Anony.Value = "";
            }
            else
            {
                HFD_Anony.Value = dr["IsAnonymous"].ToString() == "Y" ? "匿名" : "不匿名";
            }
            //生日 Birthday
            txtBirthday.Text = Util.DateTime2String(dr["Birthday"], DateType.yyyyMMdd, EmptyType.ReturnEmpty);
            //身份證字號 IDNo
            txtIDNo.Text = dr["IDNo"].ToString();
            //居住地區 IsAbroad
            HFD_LiveRegion.Value = dr["IsAbroad"].ToString() == "Y" ? "海外地區" : "台灣";
            //聯絡地址-縣市 City
            Util.FillCityData(ddlLiveCity);
            Util.SetDdlIndex(ddlLiveCity, dr["City"].ToString());
            //聯絡地址-鄉鎮 Area
            Util.FillAreaData(ddlLiveArea, dr["City"].ToString());
            Util.SetDdlIndex(ddlLiveArea, dr["Area"].ToString());
            HFD_LiveArea.Value = dr["Area"].ToString();
            //聯絡地址-郵遞區號 ZipCode
            txtLiveZip.Text = dr["ZipCode"].ToString();
            //聯絡地址-地址
            //道路街名 Street
            txtLiveStreet.Text = dr["Street"].ToString();
            //段 Section
            Util.SetDdlIndex(ddlSection, dr["Section"].ToString());
            //巷 Lane
            txtLane.Text = dr["Lane"].ToString();
            //弄 Alley
            txtAlley.Text = dr["Alley"].ToString();
            //號 HouseNo
            txtHouseNo.Text = dr["HouseNo"].ToString();
            //號之 HouseNoSub
            txtHouseNoSub.Text = dr["HouseNoSub"].ToString();
            //樓 Floor
            txtFloor.Text = dr["Floor"].ToString();
            //樓之 FloorSub
            txtFloorSub.Text = dr["FloorSub"].ToString();
            //室 Room
            txtRoom.Text = dr["Room"].ToString();

            //海外地區
            //國家/省/城市/區
            txtOverseasCountry.Text = dr["OverseasCountry"].ToString();
            //地址
            txtOverseasAddress.Text = dr["OverseasAddress"].ToString();

            //住家電話 Tel_Home
            //txtLocalPhone.Text = dr["Tel_Home"].ToString();
            //電話區碼 Tel_Office_Loc
            txtPhoneCode.Text = dr["Tel_Office_Loc"].ToString();
            //電話 Tel_Office
            txtCorpPhone.Text = dr["Tel_Office"].ToString();
            //電話分機 Tel_Ext
            txtExt.Text = dr["Tel_Office_Ext"].ToString();
            //手機電話 Cellular_Phone
            txtCellPhone.Text = dr["Cellular_Phone"].ToString();
            //傳真區碼 Fax_Loc
            txtFaxCode.Text = dr["Fax_Loc"].ToString();
            //傳真 Fax
            txtFax.Text = dr["Fax"].ToString();
            //訂閱月刊 IsSendNews
            HFD_Subscribe.Value = dr["IsSendNews"].ToString() == "Y" ? "是" : "否";
            //是否上傳給國稅局 IsFdc
            HFD_IsFdc.Value = dr["IsFdc"].ToString() == "Y" ? "是" : "否";
            //想要對GOOD TV說的話
            txtToGoodTV.Text = dr["ToGoodTV"].ToString();

            LoadFormData();



            
          
        }
    }
    //-------------------------------------------------------------------------
    protected void btnContinue_Click(object sender, EventArgs e)
    {
        Response.Redirect("DonateOnlineAll.aspx");
    }
    //-------------------------------------------------------------------------------------------------------------
    protected void btnCheckOut_Click(object sender, EventArgs e)
    {
        //Response.Redirect("CheckOut.aspx");
        //Util.ShowMsg(Util.GetControlValue("CHK_AddNew"));

        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();
        SetColumnData(list);
        string strSql = "";
        if (txtDonorID.Text.Trim() == "")
        {
            strSql = Util.CreateInsertCommand("Donor", list, dict);
            //bool bolRet = CheckLogin(txtEmail.Text, txtPassword.Text);
        }
        else
        {
            strSql = Util.CreateUpdateCommand("Donor", list, dict);
        }
        NpoDB.ExecuteSQLS(strSql, dict);

        //Util.ShowMsg("OK");
        //Session["DonorID"] = txtDonorID.Text;
        //Session["DonorName"] = txtDonorName.Text;
        //Response.Redirect(Util.RedirectByTime("DonateType.aspx"));

        if (txtDonorID.Text.Trim() == "")
        {
            strSql = @"select * from donor 
                    where Donor_Name=@DonorName 
                    and Create_DateTime=@CreateDateTime
                   ";

            dict.Clear();
            dict.Add("DonorName", txtDonorName.Text.Trim());
            dict.Add("CreateDateTime", strCreateDateTime);

            DataTable dtDonor = NpoDB.GetDataTableS(strSql, dict);

            if (dtDonor.Rows.Count > 0)
            {
                Session["DonorID"] = dtDonor.Rows[0]["Donor_Id"].ToString();
                Session["DonorName"] = txtDonorName.Text.Trim();
            }

            // 點選「完成填寫」後，顯示提示訊息 by Hilty 2013/12/18
            Session["Msg"] = "E-mail為帳號經註冊後，若需變更請電洽-奉獻服務專線：(02)8024-3911 捐款服務組";
        }

        Response.Redirect(Util.RedirectByTime("DonateInfoConfirm.aspx"));
    }
    //-------------------------------------------------------------------------------------------------------------
    private void ShowRegister()
    {
        txtAlley.Text = "";
        txtBirthday.Text = "";
        txtCellPhone.Text = "";
        txtConfirmPwd.Text = "";
        txtPhoneCode.Text = "";
        txtCorpPhone.Text = "";
        txtExt.Text = "";
        txtDonorID.Text = "";
        txtDonorName.Text = "";
        txtEmail.Text = "";
        txtFaxCode.Text = "";
        txtFax.Text = "";
        txtFloor.Text = "";
        txtFloorSub.Text = "";
        txtHouseNo.Text = "";
        txtHouseNoSub.Text = "";
        txtIDNo.Text = "";
        txtLane.Text = "";
        txtLiveStreet.Text = "";
        txtLiveZip.Text = "";
        //txtLocalPhone.Text = "";
        txtOverseasAddress.Text = "";
        txtOverseasCountry.Text = "";
        txtPassword.Text = "";
        txtRoom.Text = "";
        txtTitle.Text = "";
        txtToGoodTV.Text = "";
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
        list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_Receipt", "rdoReceipt1", "是", "是", false, HFD_Receipt.Value));
        list.Add(new ControlData("Radio", "RDO_Receipt", "rdoReceipt2", "否", "否", false, HFD_Receipt.Value));
        lblReceipt.Text = HtmlUtil.RenderControl(list);
        //徵信錄
        list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_Anony", "rdoAnony1", "匿名", "匿名", false, HFD_Anony.Value));
        list.Add(new ControlData("Radio", "RDO_Anony", "rdoAnony2", "不匿名", "不匿名", false, HFD_Anony.Value));
        lblAnony.Text = HtmlUtil.RenderControl(list);
        //訂閱月刊
        list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_Subscribe", "RdoSubscribe1", "是", "是", false, HFD_Subscribe.Value));
        list.Add(new ControlData("Radio", "RDO_Subscribe", "RdoSubscribe2", "否", "否", false, HFD_Subscribe.Value));
        lblSubscribe.Text = HtmlUtil.RenderControl(list);
        //是否上傳給國稅局
        list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_IsFdc", "rdoIsFdc1", "是", "是", false, HFD_IsFdc.Value));
        list.Add(new ControlData("Radio", "RDO_IsFdc", "rdoIsFdc2", "否", "否", false, HFD_IsFdc.Value));
        lblIsFdc.Text = HtmlUtil.RenderControl(list);
        //新增帳號
        list = new List<ControlData>();
        list.Add(new ControlData("Checkbox", "CHK_AddNew", "ChkAddNew", "新增帳號", "新增帳號", false, ""));
        lblAddNew.Text = HtmlUtil.RenderControl(list);
        //居住地區
        list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_LiveRegion", "RdoLiveRegion1", "台灣", "台灣", false, HFD_LiveRegion.Value));
        list.Add(new ControlData("Radio", "RDO_LiveRegion", "RdoLiveRegion2", "海外地區", "海外地區", false, HFD_LiveRegion.Value));
        lblLiveRegion.Text = HtmlUtil.RenderControl(list);
    }
    //-------------------------------------------------------------------------------------------------------------
    protected void btnLogin_Click(object sender, EventArgs e)
    {
        if (Session["CheckCode"] != null)
        {
            //測試時暫時驗證碼取消
            //txtCheckCode.Text = Session["CheckCode"].ToString().ToUpper();

            if (txtCheckCode.Text.ToUpper() == Session["CheckCode"].ToString().ToUpper())
            {
                if (CheckLogin(txtAccount.Text, txtPwd.Text))
                {
                    //Util.ShowMsg("OK");
                    HFD_ShowLogin.Value = "N";
                    HFD_showAddNewCheckbox.Value = "N";
                    HFD_AddNewChecked.Value = "N";

                    Session["DonorID"] = txtDonorID.Text;
                    Session["DonorName"] = txtDonorName.Text;
                    //Response.Redirect(Util.RedirectByTime("DonateType.aspx"));
                }
                else
                {
                    Util.ShowMsg("【帳號或密碼】錯誤，請重新輸入！");
                    HFD_ShowLogin.Value = "Y";
                    HFD_showAddNewCheckbox.Value = "Y";
                }
            }
            else
            {
                Util.ShowMsg("【驗證碼】錯誤，請重新輸入！");
                HFD_ShowLogin.Value = "Y";
            }
        }
        else
        {
            Util.ShowMsg("請重新登入！");
        }
    }
    //-------------------------------------------------------------------------------------------------------------
    private bool CheckLogin(string strAccount,string strPwd)
    {
        bool bolRet = false;
        string strSql = "";
        DataTable dt = null;
        DataRow dr = null;
        Dictionary<string, object> dict = new Dictionary<string, object>();

        
        strSql = @"select *
                from DONOR 
                where Email=@Email
                and Donor_Pwd=@Password
                ";
        dict.Add("Email", strAccount);
        dict.Add("Password", strPwd);
        dt = NpoDB.GetDataTableS(strSql, dict);
        //資料異常
        if (dt.Rows.Count != 0)
        {
            //SetSysMsg("查無資料!");
            bolRet = true;

            dr = dt.Rows[0];
            //捐款人編號 Donor_Id
            txtDonorID.Text = dr["Donor_Id"].ToString();
            //奉獻者姓名 Donor_Name
            txtDonorName.Text = dr["Donor_Name"].ToString();
            //性別 Sex
            HFD_Gender.Value = dr["Sex"].ToString();
            //Email Email
            txtEmail.Text = dr["Email"].ToString();
            //add by geo 20131226 for lock email when login
            txtEmail.ReadOnly = true;
            txtEmail.CssClass = "readonly";

            //密碼  若有勾選新增帳號則存密碼，反之則否
            //收據抬頭 Invoice_Title
            txtTitle.Text = dr["Invoice_Title"].ToString();
            //寄發收據 Invoice_Type
            HFD_Receipt.Value = dr["Invoice_Type"].ToString() == "不需要" ? "否" : "是";
            //徵信錄 IsAnonymous
            if (dr["IsAnonymous"].ToString() == "")
            {
                HFD_Anony.Value = "";
            }
            else
            {
                HFD_Anony.Value = dr["IsAnonymous"].ToString() == "Y" ? "匿名" : "不匿名";
            }
            //生日 Birthday
            txtBirthday.Text = Util.DateTime2String(dr["Birthday"], DateType.yyyyMMdd, EmptyType.ReturnEmpty);
            //身份證字號 IDNo
            txtIDNo.Text = dr["IDNo"].ToString();
            //居住地區 IsAbroad
            HFD_LiveRegion.Value = dr["IsAbroad"].ToString() == "Y" ? "海外地區" : "台灣";
            //聯絡地址-縣市 City
            Util.FillCityData(ddlLiveCity);
            Util.SetDdlIndex(ddlLiveCity, dr["City"].ToString());
            //聯絡地址-鄉鎮 Area
            Util.FillAreaData(ddlLiveArea, dr["City"].ToString());
            Util.SetDdlIndex(ddlLiveArea, dr["Area"].ToString());
            HFD_LiveArea.Value = dr["Area"].ToString();
            //聯絡地址-郵遞區號 ZipCode
            txtLiveZip.Text = dr["ZipCode"].ToString();
            //聯絡地址-地址
            //道路街名 Street
            txtLiveStreet.Text = dr["Street"].ToString();
            //段 Section
            Util.SetDdlIndex(ddlSection, dr["Section"].ToString());
            //巷 Lane
            txtLane.Text = dr["Lane"].ToString();
            //弄 Alley
            txtAlley.Text = dr["Alley"].ToString();
            //號 HouseNo
            txtHouseNo.Text = dr["HouseNo"].ToString();
            //號之 HouseNoSub
            txtHouseNoSub.Text = dr["HouseNoSub"].ToString();
            //樓 Floor
            txtFloor.Text = dr["Floor"].ToString();
            //樓之 FloorSub
            txtFloorSub.Text = dr["FloorSub"].ToString();
            //室 Room
            txtRoom.Text = dr["Room"].ToString();

            //海外地區
            //國家/省/城市/區
            txtOverseasCountry.Text = dr["OverseasCountry"].ToString();
            //地址
            txtOverseasAddress.Text = dr["OverseasAddress"].ToString();

            //住家電話 Tel_Home
            //txtLocalPhone.Text = dr["Tel_Home"].ToString();
            //電話區碼 Tel_Office_Loc
            txtPhoneCode.Text = dr["Tel_Office_Loc"].ToString();
            //電話 Tel_Office
            txtCorpPhone.Text = dr["Tel_Office"].ToString();
            //電話分機 Tel_Ext
            txtExt.Text = dr["Tel_Office_Ext"].ToString();
            //手機電話 Cellular_Phone
            txtCellPhone.Text = dr["Cellular_Phone"].ToString();
            //傳真區碼 Fax_Loc
            txtFaxCode.Text = dr["Fax_Loc"].ToString();
            //傳真 Fax
            txtFax.Text = dr["Fax"].ToString();
            //訂閱月刊 IsSendNews
            HFD_Subscribe.Value = dr["IsSendNews"].ToString() == "Y" ? "是" : "否";
            //是否上傳給國稅局 IsFdc
            HFD_IsFdc.Value = dr["IsFdc"].ToString() == "Y" ? "是" : "否";
            //想要對GOOD TV說的話
            txtToGoodTV.Text = dr["ToGoodTV"].ToString();

            LoadFormData();




            
            //把資料先暫存********************************************************
            gDonorID = dr["Donor_Id"].ToString(); //捐款人編號
            gName = dr["Donor_Name"].ToString(); //奉獻者姓名
            gSex = dr["Sex"].ToString();
            gEmail = dr["Email"].ToString();
            gTitle = dr["Invoice_Title"].ToString(); //收據抬頭
            
            //mod by geo 20131226 for 寄發收據always否
            //gType = dr["Invoice_Type"].ToString() == "Y" ? "是" : "否"; // 寄發收據
            gType = dr["Invoice_Type"].ToString() == "不需要" ? "否" : "是"; // 寄發收據

            gIsAnonymous = HFD_Anony.Value; //徵信錄
            gBirthday = Util.DateTime2String(dr["Birthday"], DateType.yyyyMMdd, EmptyType.ReturnEmpty);
            gIDNo = dr["IDNo"].ToString(); //身份證字號
            gAbroad = dr["IsAbroad"].ToString() == "Y" ? "海外地區" : "台灣";//居住地區
            gCity = Util.GetCityName(dr["City"].ToString());  //聯絡地址-縣市
            Util.SetDdlIndex(ddlLiveCity, dr["City"].ToString());
            gArea =Util.GetAreaName(dr["Area"].ToString());  //鄉鎮 Area
            gZipcode = dr["ZipCode"].ToString();
            gStreet = dr["Street"].ToString();//街道
            gSection =  dr["Section"].ToString(); //段
            gLane = dr["Lane"].ToString();  //巷
            gAlley = dr["Alley"].ToString();//弄
            gHouseNo = dr["HouseNo"].ToString(); //號
            gHouseNoSub = dr["HouseNoSub"].ToString(); //號之
            gFloor = dr["Floor"].ToString();  //樓
            gFloorSub = dr["FloorSub"].ToString();  //樓之
            gRoom = dr["Room"].ToString(); //室
            gOverseasCountry = dr["OverseasCountry"].ToString();//海外
            gOverseasAddress = dr["OverseasAddress"].ToString();//海外地址
            gPhoneCode = dr["Tel_Office_Loc"].ToString();//電話區碼
            gCorpPhone = dr["Tel_Office"].ToString();//電話
            gtExt = dr["Tel_Office_Ext"].ToString();//分機
            gCellPhone = dr["Cellular_Phone"].ToString(); //手機
            gFaxCode = dr["Fax_Loc"].ToString(); //傳真區碼
            gFax = dr["Fax"].ToString(); //傳真
            gSendNews = dr["IsSendNews"].ToString() == "Y" ? "是" : "否";//訂閱月刊
            gIsFdc = dr["IsFdc"].ToString() == "Y" ? "是" : "否"; //是否上傳給國稅局
            gToGoodTV = dr["ToGoodTV"].ToString(); // 想要對GOOD TV說的話
            //  *******************************************************************
        }
        return bolRet;
    }
    //-------------------------------------------------------------------------------------------------------------
    private void SetColumnData(List<ColumnData> list)
    {

        //********************************************************
        string strSql = "";
        int  Cnt =0;
  
        DataTable dt = null;
        DataRow dr = null;
        Dictionary<string, object> dict = new Dictionary<string, object>();


        // Session["DonorID"] 沒有值的時候用的===============================================
        if (Session["DonorID"] == "" || Session["DonorID"] == null) 
        {
        DataRow dr1 = null;
        string strSql1 = " select max(donor_id) as Cnt from donor";
        Dictionary<string, object> dict1 = new Dictionary<string, object>();
        dict.Add("1", 1);
        DataTable dt1 = NpoDB.GetDataTableS(strSql1, dict1);
               //群組名稱
        dr1 = dt1.Rows[0];
        Cnt = int.Parse(dr1["Cnt"].ToString());
        Cnt = Cnt + 1;
        Session["DonorID"] = Cnt;
       }
        //===================================================

        strSql = @"select *
                from DONOR 
                where Donor_Id=@DonorID
                ";

        dict.Add("DonorID", Session["DonorID"].ToString());
        dt = NpoDB.GetDataTableS(strSql, dict);
        //資料異常
        if (dt.Rows.Count != 0)
        {
            dr = dt.Rows[0];
            gDonorID = dr["Donor_Id"].ToString(); //捐款人編號
            gName = dr["Donor_Name"].ToString(); //奉獻者姓名
            gSex = dr["Sex"].ToString();
            gEmail = dr["Email"].ToString();
            gTitle = dr["Invoice_Title"].ToString(); //收據抬頭
            
            //mod by geo 20131226 for 寄發收據always否
            //gType = dr["Invoice_Type"].ToString() == "Y" ? "是" : "否"; // 寄發收據
            gType = dr["Invoice_Type"].ToString() == "不需要" ? "否" : "是"; // 寄發收據

            gIsAnonymous = HFD_Anony.Value; //徵信錄
            gBirthday = Util.DateTime2String(dr["Birthday"], DateType.yyyyMMdd, EmptyType.ReturnEmpty);
            gIDNo = dr["IDNo"].ToString(); //身份證字號
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
            gOverseasCountry = dr["OverseasCountry"].ToString();//海外
            gOverseasAddress = dr["OverseasAddress"].ToString();//海外地址
            gPhoneCode = dr["Tel_Office_Loc"].ToString();//電話區碼
            gCorpPhone = dr["Tel_Office"].ToString();//電話
            gtExt = dr["Tel_Office_Ext"].ToString();//分機
            gCellPhone = dr["Cellular_Phone"].ToString(); //手機
            gFaxCode = dr["Fax_Loc"].ToString(); //傳真區碼
            gFax = dr["Fax"].ToString(); //傳真
            gSendNews = dr["IsSendNews"].ToString() == "Y" ? "是" : "否";//訂閱月刊
            gIsFdc = dr["IsFdc"].ToString() == "Y" ? "是" : "否"; //是否上傳給國稅局
            gToGoodTV = dr["ToGoodTV"].ToString(); // 想要對GOOD TV說的話            
        }
        //**************************************************************

        //第一個欄位給 update 用
        //list.Add(new ColumnData("Uid", HFD_DonorID.Value, false, false, true));
        //捐款人編號
        list.Add(new ColumnData("Donor_Id", txtDonorID.Text, false, false, true));
        //奉獻者姓名
        list.Add(new ColumnData("Donor_Name", txtDonorName.Text, true, true, false));
        //性別
        list.Add(new ColumnData("Sex", Util.GetControlValue("RDO_Gender"), true, true, false));
        //Email
        list.Add(new ColumnData("Email", txtEmail.Text, true, false, false));
        //密碼  若有勾選新增帳號則存密碼，反之則否
        if (Util.GetControlValue("CHK_AddNew") != "")
        {
            list.Add(new ColumnData("Donor_Pwd", txtPassword.Text, true, true, false));
        }
        //收據抬頭
        list.Add(new ColumnData("Invoice_Title", txtTitle.Text, true, true, false));
        //寄發收據
        //mod by geo 20131226 for 寄發收據always否
        //list.Add(new ColumnData("Invoice_Type", Util.GetControlValue("RDO_Receipt") == "" ? "不需要" : "年證明及收據", true, true, false));
        list.Add(new ColumnData("Invoice_Type", Util.GetControlValue("RDO_Receipt") == "是" ? "單次收據" : "不需要", true, true, false));


        //徵信錄
        list.Add(new ColumnData("IsAnonymous", Util.GetControlValue("RDO_Anony") == "匿名" ? "Y" : "N", true, true, false));
        //生日
        list.Add(new ColumnData("Birthday", Util.DateTime2String(txtBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnNull), true, true, false));
        //身份證字號
        list.Add(new ColumnData("IDNo", txtIDNo.Text, true, true, false));
        //居住地區
        list.Add(new ColumnData("IsAbroad", Util.GetControlValue("RDO_LiveRegion") == "台灣" ? "N" : "Y", true, true, false));

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
            list.Add(new ColumnData("HouseNoSub", txtHouseNoSub.Text.Trim(), true, true, false));
            list.Add(new ColumnData("Floor", txtFloor.Text.Trim(), true, true, false));
            list.Add(new ColumnData("FloorSub", txtFloorSub.Text.Trim(), true, true, false));
            list.Add(new ColumnData("Room", txtRoom.Text.Trim(), true, true, false));

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
                strLiveAddress += txtHouseNo.Text.Trim() + "號";
            }
            if (txtHouseNoSub.Text.Trim() != "")
            {
                strLiveAddress += "之" + txtHouseNoSub.Text.Trim();
            }
            if (txtFloor.Text.Trim() != "")
            {
                strLiveAddress += "," + txtFloor.Text.Trim() + "樓";
            }
            if (txtFloorSub.Text.Trim() != "")
            {
                strLiveAddress += "之" + txtFloorSub.Text.Trim();
            }
            if (txtRoom.Text.Trim() != "")
            {
                strLiveAddress += "," + txtRoom.Text.Trim() + "室";
            }
            strLiveAddress2 = Util.GetControlValue("RDO_LiveRegion") + ddlLiveCity.SelectedItem.Text + Util.GetAreaName(HFD_LiveArea.Value) + strLiveAddress;
            if (Util.GetControlValue("RDO_Receipt") != "")
            {
                //聯絡地址-縣市
                list.Add(new ColumnData("Invoice_City", ddlLiveCity.SelectedValue, true, true, false));
                //聯絡地址-鄉鎮
                list.Add(new ColumnData("Invoice_Area", HFD_LiveArea.Value, true, true, false));
                //聯絡地址-郵遞區號
                list.Add(new ColumnData("Invoice_ZipCode", HFD_LiveArea.Value, true, true, false));
                //聯絡地址-地址
                list.Add(new ColumnData("Invoice_Address", strLiveAddress, true, true, false));
            }
        }
        else
        {


            //海外地區
            //國家/省/城市/區
            list.Add(new ColumnData("OverseasCountry", txtOverseasCountry.Text.Trim(), true, true, false));
            //地址
            list.Add(new ColumnData("OverseasAddress", txtOverseasAddress.Text.Trim(), true, true, false));
            strLiveAddress = txtOverseasAddress.Text.Trim() + " " + txtOverseasCountry.Text.Trim();
            if (Util.GetControlValue("RDO_Receipt") != "")
            {
                list.Add(new ColumnData("Invoice_Address", strLiveAddress, true, true, false));
            }
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
        list.Add(new ColumnData("Fax_Loc", txtFaxCode.Text, true, true, false));
        //傳真
        list.Add(new ColumnData("Fax", txtFax.Text, true, true, false));
        //訂閱月刊
        list.Add(new ColumnData("IsSendNews", Util.GetControlValue("RDO_Subscribe") == "是" ? "Y" : "N", true, true, false));
        //是否上傳給國稅局
        list.Add(new ColumnData("IsFdc", Util.GetControlValue("RDO_IsFdc") == "是" ? "Y" : "N", true, true, false));
        //想要對GOOD TV說的話
        list.Add(new ColumnData("ToGoodTV", txtToGoodTV.Text, true, true, false));
        //20130930 Add by GoodTV Tanya
        //隸屬組織
        list.Add(new ColumnData("Dept_Id","C001", true, true, false));
        //是否為讀者
        list.Add(new ColumnData("IsMember", "N", true, true, false));

        //新增日期
        strCreateDateTime = Util.GetToday(DateType.yyyyMMddHHmmss); //for get new donor_id
        list.Add(new ColumnData("Create_Date", Util.GetToday(DateType.yyyyMMdd), true, false, false));
        list.Add(new ColumnData("Create_DateTime", strCreateDateTime, true, false, false));
        list.Add(new ColumnData("Create_User", "Online", true, false, false));
        //最後更新日期
        list.Add(new ColumnData("LastUpdate_Date", Util.GetToday(DateType.yyyyMMdd), false, true, false));
        list.Add(new ColumnData("LastUpdate_DateTime", Util.GetToday(DateType.yyyyMMddHHmmss), false, true, false));
        list.Add(new ColumnData("LastUpdate_User", "Online", false, true, false));



        //異動資料*****************************************************************

        string Mailhead = "";
        string SmtpServer = System.Configuration.ConfigurationSettings.AppSettings["MailServer"];
        string MailFrom = "";
        string MailSubject = "";
        string MailBody = "";
        string MailTo = "";
        
        if (HFD_AddNewChecked.Value != "Y")
        {
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
            strAddr2 = gAbroad + gCity + gArea + strAddr;  //異動前地址
            if (txtDonorName.Text != gName) strReport = strReport + "原姓名：" + gName + "-->更改為：" + txtDonorName.Text + "</BR>";
            if (Util.GetControlValue("RDO_Gender") != gSex) strReport = strReport + "原性別：" + gSex + "-->更改為：" + Util.GetControlValue("RDO_Gender") + "</BR>";
            //if (txtEmail.Text != gEmail) strReport = strReport + gEmail + "-->更改為：" + txtEmail.Text + "</BR>";
            if (txtTitle.Text != gTitle) strReport = strReport + "抬頭：" + gTitle + "-->更改為：" + txtTitle.Text + "</BR>";
            if (Util.GetControlValue("RDO_Receipt") != gType) strReport = strReport + "寄發收據：" + gType + "-->更改為：" + Util.GetControlValue("RDO_Receipt") + "</BR>";
            if (Util.GetControlValue("RDO_Anony") != gIsAnonymous) strReport = strReport + "徵信錄：" + gIsAnonymous + "-->更改為：" + Util.GetControlValue("RDO_Anony") + "</BR>";
            if (Util.DateTime2String(txtBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnEmpty ) != gBirthday) strReport = strReport + "原生日：" + gBirthday + "-->更改為：" + Util.DateTime2String(txtBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnNull) + "</BR>";
            if (txtIDNo.Text != gIDNo) strReport = strReport + "身分證字號：" + gIDNo + "-->更改為：" + txtIDNo.Text + "</BR>";
            if (strLiveAddress2 != strAddr2) strReport = strReport + "原地址：" + strAddr2 + "-->更改為：" + strLiveAddress2 + "</BR>";
            if (txtOverseasCountry.Text.Trim() != gOverseasCountry) strReport = strReport + gOverseasCountry + "-->更改為：" + txtOverseasCountry.Text.Trim() + "</BR>";
            if (txtOverseasAddress.Text.Trim() != gOverseasAddress) strReport = strReport + gOverseasAddress + "-->更改為：" + txtOverseasAddress.Text.Trim() + "</BR>";
            if (txtPhoneCode.Text != gPhoneCode) strReport = strReport + "電話區碼：" + gPhoneCode + "-->更改為：" + txtPhoneCode.Text + "</BR>";
            if (txtCorpPhone.Text != gCorpPhone) strReport = strReport + "電話：" + gCorpPhone + "-->更改為：" + txtCorpPhone.Text + "</BR>";
            if (txtExt.Text != gtExt) strReport = strReport + "分機：" + gtExt + "-->更改為：" + txtExt.Text + "</BR>";
            if (txtCellPhone.Text != gCellPhone) strReport = strReport + "手機：" + gCellPhone + "-->更改為：" + txtCellPhone.Text + "</BR>";
            if (txtFaxCode.Text != gFaxCode) strReport = strReport + "傳真區碼：" + gFaxCode + "-->更改為：" + txtFaxCode.Text + "</BR>";
            if (txtFax.Text != gFax) strReport = strReport + "傳真：" + gFax + "-->更改為：" + txtFax.Text + "</BR>";
            if (Util.GetControlValue("RDO_Subscribe") != gSendNews) strReport = strReport + "訂閱月刊：" + gSendNews + "-->更改為：" + Util.GetControlValue("RDO_Subscribe") + "</BR>";
            if (Util.GetControlValue("RDO_IsFdc") != gIsFdc) strReport = strReport + "上傳給國稅局：" + gIsFdc + "-->更改為：" + Util.GetControlValue("RDO_IsFdc") + "</BR>";
            if (txtToGoodTV.Text != gToGoodTV) strReport = strReport + "想要對GOOD TV說的話：" + gToGoodTV + "-->更改為：" + txtToGoodTV.Text;

            if (strReport != "")  //若有異動資料則發信
            {

              
                Mailhead = "親愛的 " + txtDonorName.Text + " 平安！<BR> 您在GOODTV網站上奉獻時，更改了您的個人資料，異動內容如下，敬請比對查閱，謝謝您。<BR>";
                strReport += "<BR><BR><BR>";
                strReport += "若您有任何資料異動相關問題敬請來電: 02-8024-3911 查詢<BR>";
                strReport += "願神祝福您!<BR>";
                strReport += "GOODTV捐款管理組敬上<BR>";
             
                SendEMailObject MailObject = new SendEMailObject();
                MailObject.SmtpServer = SmtpServer;
                MailSubject = "GOODTV異動通知信";
                 MailBody = Mailhead + strReport;
                 MailTo = txtEmail.Text;
                 
                 MailFrom = System.Configuration.ConfigurationSettings.AppSettings["MailFrom"];
                string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, MailBody);
                if (MailObject.ErrorCode != 0)
                {
                    this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
                }
             
            }
        }
        else
        {
            string strBody = "";
            strBody = "親愛的 " + txtDonorName.Text + " 平安！<BR> 您的GOODTV帳號已於 " + strCreateDateTime  + " 註冊成功，以下是您的帳號及密碼, 請妥善保管喔!<BR>";
            strBody += "您的帳號：" + txtEmail.Text + "<BR>";
            strBody += "您的密碼：" + txtPassword.Text + "<BR>";
            strBody += "願神祝福您!<BR><BR>GOODTV捐款管理組敬上";
            
            
       

            SendEMailObject MailObject = new SendEMailObject();
            MailObject.SmtpServer = SmtpServer;
            MailSubject = " GOODTV線上奉獻註冊成功通知信";
            MailBody = strBody;
            MailTo = txtEmail.Text;
            MailFrom = System.Configuration.ConfigurationSettings.AppSettings["MailFrom"];
            string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, strBody);
            if (MailObject.ErrorCode != 0)
            {
                this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
            }
        }
        //********************************************************************
    }

    protected void btnRegister_Click(object sender, EventArgs e)
    {
        Response.Redirect("CheckOut.aspx?AddNewChecked=Y");
    }
}