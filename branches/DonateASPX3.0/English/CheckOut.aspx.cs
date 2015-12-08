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
    string gAttn;
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
    //string strInvoice_Type;

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
                //註冊新帳號用
                if (Util.GetQueryString("AddNewChecked") == "Y")
                {
                    HFD_ShowLogin.Value = "N";
                    HFD_AddNewChecked.Value = "Y";
                }
                else
                {
                    HFD_ShowLogin.Value = "Y";  
                }
                HFD_showAddNewCheckbox.Value = "Y";
                //btnOpenLogin.Visible = true;
                txtEmail.ReadOnly = false;
                txtEmail.CssClass = "";
            }
            else
            {
                HFD_ShowLogin.Value = "N";
                HFD_showAddNewCheckbox.Value = "N";
                HFD_AddNewChecked.Value = "N";
                //btnOpenLogin.Visible = false;
                txtEmail.ReadOnly = true;
                txtEmail.CssClass = "readonly";                
            }
            //HFD_IsFdc.Value = "是";
            HFD_LiveRegion.Value = "海外地區";
            HFD_Receipt.Value = "Not necessary";
            HFD_Subscribe.Value = "否";
            HFD_Anony.Value = "不刊登";
            LoadFormData();

            if (Session["DonorID"] != null)
            {
                //Response.Redirect(Util.RedirectByTime("DonateInfoConfirm.aspx"));
                LoadData(Session["DonorID"].ToString());
            }
            //for 註冊新帳號用
            //HFD_AddNewChecked.Value = Util.GetQueryString("AddNewChecked");

        }

        if (Session["ItemOnce"] != null)
        {
            dtOnce = (DataTable)Session["ItemOnce"];
            //strInvoice_Type = "單次收據";
        }
        if (Session["ItemPeriod"] != null)
        {
            dtPeriod = (DataTable)Session["ItemPeriod"];
            //strInvoice_Type = "年度證明";
        }

        int iOnce = dtOnce.Rows.Count;
        int iPeriod = dtPeriod.Rows.Count;
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
                where Donor_Id=@DonorID and DeleteDate is null
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
            // 
            //HFD_Receipt.Value = dr["Invoice_Type"].ToString() == "不寄" ? "是" : "否";
            //HFD_Receipt.Value = dr["Invoice_Type"].ToString();
            if (dr["Invoice_Type"].ToString() == "不寄")
            { 
                HFD_Receipt.Value = "Not necessary";
            }
            else if (dr["Invoice_Type"].ToString() == "年度證明")
            {
                HFD_Receipt.Value = "One receipt for your total annual donation";
            }
            else
            {
                HFD_Receipt.Value = "After each donation payment";
            }
            //徵信錄 IsAnonymous
            if (dr["IsAnonymous"].ToString() == "")
            {
                HFD_Anony.Value = "";
            }
            else
            {
                HFD_Anony.Value = dr["IsAnonymous"].ToString() == "Y" ? "不刊登" : "刊登";
            }
            //生日 Birthday
            txtBirthday.Text = Util.DateTime2String(dr["Birthday"], DateType.yyyyMMdd, EmptyType.ReturnEmpty);
            //身份證字號 IDNo
            //txtIDNo.Text = dr["IDNo"].ToString();
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
            //Attn
            txtAttn.Text = dr["Attn"].ToString();

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
            //HFD_IsFdc.Value = dr["IsFdc"].ToString() == "Y" ? "是" : "否";
            //想要對GOOD TV說的話
            txtToGoodTV.Text = dr["ToGoodTV"].ToString();

            LoadFormData();

        }
    }
    //-------------------------------------------------------------------------
    protected void btnPrev_Click(object sender, EventArgs e)
    {
        //Response.Redirect("DonateOnlineAll.aspx");
        Response.Redirect("ShoppingCart.aspx?Mode=FromDefault");
    }
    //-------------------------------------------------------------------------------------------------------------
    protected void btnCheckOut_Click(object sender, EventArgs e)
    {
        //Response.Redirect("CheckOut.aspx");
        //Util.ShowMsg(Util.GetControlValue("CHK_AddNew"));
        if (HFD_AddNewChecked.Value != "N")
        {
            string strSql2 = "";

            DataTable dt2 = null;
            Dictionary<string, object> dict2 = new Dictionary<string, object>();
            strSql2 = "select * from DONOR where Email=@Email and Donor_Pwd Is not NULL and DeleteDate is null";

            dict2.Add("Email", txtEmail.Text.Trim());
            dt2 = NpoDB.GetDataTableS(strSql2, dict2);
            if (dt2.Rows.Count != 0)
            {
                this.Page.RegisterStartupScript("s", "<script>alert('此Email帳號已存在，請先登入！');</script>");
                return;
            }
        }

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

        // 2014/6/6 取消此邏輯問題
        /*
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
        */
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
        //txtIDNo.Text = "";
        txtLane.Text = "";
        txtLiveStreet.Text = "";
        txtLiveZip.Text = "";
        //txtLocalPhone.Text = "";
        txtOverseasAddress.Text = "";
        txtOverseasCountry.Text = "";
        txtPassword.Text = "";
        txtRoom.Text = "";
        txtAttn.Text = "";
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
        list.Add(new ControlData("Radio", "RDO_Gender", "rdoSex1", "男", "Male", false, HFD_Gender.Value));
        list.Add(new ControlData("Radio", "RDO_Gender", "rdoSex2", "女", "Female", false, HFD_Gender.Value));
        lblGender.Text = HtmlUtil.RenderControl(list);
        //寄發收據
        list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_Receipt", "rdoReceipt1", "One receipt for your total annual donation", "One receipt for your total annual donation", false, HFD_Receipt.Value));
        list.Add(new ControlData("Radio", "RDO_Receipt", "rdoReceipt2", "After each donation payment", "After each donation payment", false, HFD_Receipt.Value));
        list.Add(new ControlData("Radio", "RDO_Receipt", "rdoReceipt3", "Not necessary", "Not necessary", false, HFD_Receipt.Value));
        lblReceipt.Text = HtmlUtil.RenderControl(list);
        //徵信錄
        /*list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_Anony", "rdoAnony1", "不刊登", "不刊登", false, HFD_Anony.Value));
        list.Add(new ControlData("Radio", "RDO_Anony", "rdoAnony2", "刊登", "刊登", false, HFD_Anony.Value));
        lblAnony.Text = HtmlUtil.RenderControl(list);*/
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
        list.Add(new ControlData("Checkbox", "CHK_AddNew", "ChkAddNew", "新增帳號", "Create Account", false, ""));
        lblAddNew.Text = HtmlUtil.RenderControl(list);
        //居住地區
        list = new List<ControlData>();
        list.Add(new ControlData("Radio", "RDO_LiveRegion", "RdoLiveRegion1", "台灣", "Taiwan", false, HFD_LiveRegion.Value));
        list.Add(new ControlData("Radio", "RDO_LiveRegion", "RdoLiveRegion2", "海外地區", "Overseas", false, HFD_LiveRegion.Value));
        lblLiveRegion.Text = HtmlUtil.RenderControl(list);
    }
    //-------------------------------------------------------------------------------------------------------------
    protected void btnLogin_Click(object sender, EventArgs e)
    {
        //if (Session["CheckCode"] != null)
        //{
            //測試時暫時驗證碼取消
            //txtCheckCode.Text = Session["CheckCode"].ToString().ToUpper();

            //if (txtCheckCode.Text.ToUpper() == Session["CheckCode"].ToString().ToUpper())
            //{
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
            //}
            //else
            //{
            //    Util.ShowMsg("【驗證碼】錯誤，請重新輸入！");
            //    HFD_ShowLogin.Value = "Y";
            //}
        //}
        //else
        //{
        //    Util.ShowMsg("請重新登入！");
        //}
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
                and Donor_Pwd=@Password and DeleteDate is null
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
            //HFD_Receipt.Value = dr["Invoice_Type"].ToString();
            if (dr["Invoice_Type"].ToString() == "不寄")
            {
                HFD_Receipt.Value = "Not necessary";
            }
            else if (dr["Invoice_Type"].ToString() == "年度證明")
            {
                HFD_Receipt.Value = "One receipt for your total annual donation";
            }
            else
            {
                HFD_Receipt.Value = "After each donation payment";
            }
            //徵信錄 IsAnonymous
            if (dr["IsAnonymous"].ToString() == "")
            {
                HFD_Anony.Value = "";
            }
            else
            {
                HFD_Anony.Value = dr["IsAnonymous"].ToString() == "Y" ? "不刊登" : "刊登";
            }
            //生日 Birthday
            txtBirthday.Text = Util.DateTime2String(dr["Birthday"], DateType.yyyyMMdd, EmptyType.ReturnEmpty);
            //身份證字號 IDNo
            //txtIDNo.Text = dr["IDNo"].ToString();
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
            //Attn
            txtAttn.Text = dr["Attn"].ToString();

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
            // 2014/6/10 User需求暫時取消
            //是否上傳給國稅局 IsFdc
            //HFD_IsFdc.Value = dr["IsFdc"].ToString() == "Y" ? "是" : "否";
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
            gType = dr["Invoice_Type"].ToString(); // 寄發收據

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
            gAttn = dr["Attn"].ToString(); //Attn
            gOverseasCountry = dr["OverseasCountry"].ToString();//海外
            gOverseasAddress = dr["OverseasAddress"].ToString();//海外地址
            gPhoneCode = dr["Tel_Office_Loc"].ToString();//電話區碼
            gCorpPhone = dr["Tel_Office"].ToString();//電話
            gtExt = dr["Tel_Office_Ext"].ToString();//分機
            gCellPhone = dr["Cellular_Phone"].ToString(); //手機
            gFaxCode = dr["Fax_Loc"].ToString(); //傳真區碼
            gFax = dr["Fax"].ToString(); //傳真
            gSendNews = dr["IsSendNews"].ToString() == "Y" ? "是" : "否";//訂閱月刊
            //gIsFdc = dr["IsFdc"].ToString() == "Y" ? "是" : "否"; //是否上傳給國稅局
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
        if (Session["DonorID"] == null)
        {
            DataRow dr1 = null;
            string strSql1 = " select max(donor_id) as Cnt from donor";
            Dictionary<string, object> dict1 = new Dictionary<string, object>();
            //dict.Add("1", 1);
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
                where Donor_Id=@DonorID and DeleteDate is null
                ";

        dict.Add("DonorID", Session["DonorID"].ToString());
        dt = NpoDB.GetDataTableS(strSql, dict);
        //資料異常
        if (dt.Rows.Count != 0)
        {
            dr = dt.Rows[0];
            gDonorID = dr["Donor_Id"].ToString(); //捐款人編號
            gName = dr["Donor_Name"].ToString(); //奉獻者姓名
            Session["DonorName"] = dr["Donor_Name"].ToString();
            gSex = dr["Sex"].ToString();
            gEmail = dr["Email"].ToString();
            gTitle = dr["Invoice_Title"].ToString(); //收據抬頭

            //mod by geo 20131226 for 寄發收據always否
            //gType = dr["Invoice_Type"].ToString() == "Y" ? "是" : "否"; // 寄發收據
            gType = dr["Invoice_Type"].ToString(); // 寄發收據

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
            gAttn = dr["Attn"].ToString(); //Attn
            gOverseasCountry = dr["OverseasCountry"].ToString();//海外
            gOverseasAddress = dr["OverseasAddress"].ToString();//海外地址
            gPhoneCode = dr["Tel_Office_Loc"].ToString();//電話區碼
            gCorpPhone = dr["Tel_Office"].ToString();//電話
            gtExt = dr["Tel_Office_Ext"].ToString();//分機
            gCellPhone = dr["Cellular_Phone"].ToString(); //手機
            gFaxCode = dr["Fax_Loc"].ToString(); //傳真區碼
            gFax = dr["Fax"].ToString(); //傳真
            gSendNews = dr["IsSendNews"].ToString() == "Y" ? "是" : "否";//訂閱月刊
            //gIsFdc = dr["IsFdc"].ToString() == "Y" ? "是" : "否"; //是否上傳給國稅局
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
        //稱謂
        list.Add(new ColumnData("Title", Util.GetControlValue("RDO_Gender") == "男" ? "先生" : "小姐", true, true, false));
        //身份別
        list.Add(new ColumnData("Donor_Type", "個人", true, false, false));
        //Email
        list.Add(new ColumnData("Email", txtEmail.Text, true, false, false));
        //密碼  若有勾選新增帳號則存密碼，反之則否
        if (Util.GetControlValue("CHK_AddNew") != "")
        {
            list.Add(new ColumnData("Donor_Pwd", txtPassword.Text, true, true, false));
        }
        //收據抬頭
        if (txtTitle.Text.Trim() != "")
        {
            list.Add(new ColumnData("Invoice_Title", txtTitle.Text, true, true, false));
        }
        else //帶入捐款人姓名
        {
            list.Add(new ColumnData("Invoice_Title", txtDonorName.Text, true, true, false));
        }
        //寄發收據
        //mod by geo 20131226 for 寄發收據always否
        //list.Add(new ColumnData("Invoice_Type", Util.GetControlValue("RDO_Receipt") == "" ? "不需要" : "年證明及收據", true, true, false));
        // 2014/6/5 修正收據開立判斷
        //list.Add(new ColumnData("Invoice_Type", Util.GetControlValue("RDO_Receipt") == "是" ? "單次收據" : "不寄", true, true, false));
        //list.Add(new ColumnData("Invoice_Type", Util.GetControlValue("RDO_Receipt"), true, true, false));
        string Receipt_type = "";
        switch (Util.GetControlValue("RDO_Receipt"))
        {
            case "One receipt for your total annual donation":
                Receipt_type = "年度證明";
                break;
            case "After each donation payment":
                Receipt_type = "單次收據";
                break;
            case "Not necessary":
                Receipt_type = "不寄";
                break;
        }
        Session["Receipt_type"] = Receipt_type;
        list.Add(new ColumnData("Invoice_Type", Receipt_type, true, true, false));

        //徵信錄
        //list.Add(new ColumnData("IsAnonymous", Util.GetControlValue("RDO_Anony") == "不刊登" ? "Y" : "N", true, true, false));
        //生日
        list.Add(new ColumnData("Birthday", Util.DateTime2String(txtBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnNull), true, true, false));
        //身份證字號
        //list.Add(new ColumnData("IDNo", txtIDNo.Text, true, true, false));
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
            list.Add(new ColumnData("HouseNoSub", txtHouseNoSub.Text.Trim(), true, true, false));
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
            if (txtHouseNoSub.Text.Trim() != "")
            {
                strLiveAddress += "之" + txtHouseNoSub.Text.Trim();
            }
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
                strLiveAddress += " -" + txtRoom.Text.Trim() + "室";
            }
            //marked out by Samuel for removing duplicated attn shown in both invoice and mailing address 2014/08/11 
			//if (txtAttn.Text.Trim() != "")
            //{
               // strLiveAddress += "(" + txtAttn.Text.Trim() + ")";
            //}
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

                list.Add(new ColumnData("Invoice_Street", txtLiveStreet.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Section", ddlSection.SelectedValue, true, true, false));
                list.Add(new ColumnData("Invoice_Lane", txtLane.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Alley", txtAlley.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_HouseNo", txtHouseNo.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_HouseNoSub", txtHouseNoSub.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Floor", txtFloor.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_FloorSub", txtFloorSub.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Room", txtRoom.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Attn", txtAttn.Text.Trim(), true, true, false));
            }
            //訂購紙本月刊 20140623修改
            if (Util.GetControlValue("RDO_Subscribe") == "是")
            {
                if (strLiveAddress == "") //若未填地址，自動存不訂購月刊
                {
                    list.Add(new ColumnData("IsSendNews", "N", true, true, false));
                    list.Add(new ColumnData("IsSendNewsNum", "NULL", true, true, false));
                }
                else
                {
                    list.Add(new ColumnData("IsSendNews", "Y", true, true, false));
                    list.Add(new ColumnData("IsSendNewsNum", "1", true, true, false));
                }
            }
            else
            {
                list.Add(new ColumnData("IsSendNews", "N", true, true, false));
                list.Add(new ColumnData("IsSendNewsNum", "NULL", true, true, false));
            }
        }
        else
        {
            //海外地區
            //國家/省/城市/區
            list.Add(new ColumnData("OverseasCountry", txtOverseasCountry.Text.Trim(), true, true, false));
            //地址
            list.Add(new ColumnData("OverseasAddress", txtOverseasAddress.Text.Trim(), true, true, false));
            //strLiveAddress = txtOverseasCountry.Text.Trim() + " " + txtOverseasAddress.Text.Trim();
            strLiveAddress = txtOverseasAddress.Text.Trim() + " " + txtOverseasCountry.Text.Trim();
            strLiveAddress2 = Util.GetControlValue("RDO_LiveRegion") + strLiveAddress;
            if (Util.GetControlValue("RDO_Receipt") != "")
            {
                //收據國家/省/城市/區
                list.Add(new ColumnData("Invoice_OverseasCountry", txtOverseasCountry.Text.Trim(), true, true, false));
                //收據地址
                list.Add(new ColumnData("Invoice_OverseasAddress", txtOverseasAddress.Text.Trim(), true, true, false));
                list.Add(new ColumnData("Invoice_Address", strLiveAddress, true, true, false));
            }
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
        list.Add(new ColumnData("Fax_Loc", txtFaxCode.Text, true, true, false));
        //傳真
        list.Add(new ColumnData("Fax", txtFax.Text, true, true, false));
        //訂閱月刊
        //list.Add(new ColumnData("IsSendNews", Util.GetControlValue("RDO_Subscribe") == "是" ? "Y" : "N", true, true, false));
        //6/11 User需求暫時取消
        //是否上傳給國稅局
        //list.Add(new ColumnData("IsFdc", Util.GetControlValue("RDO_IsFdc") == "是" ? "Y" : "N", true, true, false));
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
        list.Add(new ColumnData("Create_User", "線上金流", true, false, false));
        //最後更新日期
        list.Add(new ColumnData("LastUpdate_Date", Util.GetToday(DateType.yyyyMMdd), false, true, false));
        list.Add(new ColumnData("LastUpdate_DateTime", Util.GetToday(DateType.yyyyMMddHHmmss), false, true, false));
        list.Add(new ColumnData("LastUpdate_User", "線上金流", false, true, false));



        //異動資料*****************************************************************

        string Mailhead = "";
        //string SmtpServer = System.Configuration.ConfigurationSettings.AppSettings["MailServer"];
        string MailFrom = "";
        string MailSubject = "";
        string MailBody = "";
        string MailTo = "";
        string strReportToDonations = ""; //紀錄異常的(內部Email使用)

        if (HFD_AddNewChecked.Value != "Y")
        {
            if (HFD_ShowLogin.Value == "N")
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
				//marked out by Samuel for removing duplicated attn shown in both invoice and mailing address 2014/08/11 
                //if (gAttn != "")
                //{
                    //strAddr += "(" + gAttn + ")";
                //}
                // 2014/6/12 更正地址邏輯(台灣與海外地區) by 詩儀
                //strAddr = (Util.GetControlValue("RDO_LiveRegion") == "台灣") ? strAddr : txtOverseasCountry.Text.Trim() + " " + txtOverseasAddress.Text.Trim();
                if (Util.GetControlValue("RDO_LiveRegion") != gAbroad) //海外地區<==>台灣
                {
                    if (gAbroad == "台灣")
                    {
                        strAddr2 = gAbroad + gCity + gArea + strAddr;  //異動前地址
                    }
                    else
                    {
                        // 2014/7/7 看到應該是address+Country
                        //strAddr2 = gAbroad + gOverseasCountry + " " + ;//異動前海外地址
                        strAddr2 = gAbroad + gOverseasAddress + " " + gOverseasCountry;//異動前海外地址
                    }
                }
                else  //台灣<==>台灣 or 海外<==>海外
                {
                    if (Util.GetControlValue("RDO_LiveRegion") == "台灣")
                    {
                        strAddr2 = gAbroad + gCity + gArea + strAddr;  //異動前地址
                    }
                    else 
                    {
                        // 2014/7/7 看到應該是address+Country
                        //strAddr2 = gAbroad + gOverseasCountry + " " + gOverseasAddress;//異動前海外地址
                        strAddr2 = gAbroad + gOverseasAddress + " " + gOverseasCountry;//異動前海外地址
                    }
                }

                //20140616 修正只顯示更改後的資料。
                if (txtDonorName.Text != gName) strReport = strReport + "Name：" + txtDonorName.Text + "<BR><BR>";
                if (Util.GetControlValue("RDO_Gender") != gSex) strReport = strReport + "Gender：" + Util.GetControlValue("RDO_Gender") + "<BR><BR>";
                //if (txtEmail.Text != gEmail) strReport = strReport + gEmail + "-->更改為：" + txtEmail.Text + "</BR>";
                if (txtTitle.Text != gTitle)
                {
                    strReport = strReport + "Name of receipt：" + txtTitle.Text + "<BR><BR>";
                    strReportToDonations = strReportToDonations + "Name of receipt：" + txtTitle.Text + "<BR><BR>";
                }
                if (Session["Receipt_type"].ToString() != gType)
                {
                    strReport = strReport + "Send receipt：" + Session["Receipt_type"].ToString() + "<BR><BR>";
                    strReportToDonations = strReportToDonations + "Send receipt：" + Session["Receipt_type"].ToString() + "<BR><BR>";
                }
                //if (Util.GetControlValue("RDO_Anony") != gIsAnonymous) strReport = strReport + "徵信錄：" + Util.GetControlValue("RDO_Anony") + "<BR><BR>";
                if (Util.DateTime2String(txtBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnEmpty) != gBirthday) strReport = strReport + "Birthday：" + Util.DateTime2String(txtBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnNull) + "<BR><BR>";
                //if (txtIDNo.Text != gIDNo) strReport = strReport + "身分證字號：" + txtIDNo.Text + "<BR><BR>";

                if (strLiveAddress2 != strAddr2)
                {
                    strReport = strReport + "Mailing address：" + strLiveAddress2 + "<BR><BR>";
                    strReportToDonations = strReportToDonations + "Mailing address：" + strLiveAddress2 + "<BR><BR>";
                }

                // 2014/7/8 電話區碼+電話+電話分機
                if ((txtPhoneCode.Text + txtCorpPhone.Text + txtExt.Text) != (gPhoneCode + gCorpPhone + gtExt))
                {
                    strReport = strReport + "Tel：" + txtPhoneCode.Text + "-" + txtCorpPhone.Text + (txtExt.Text!=""?" Ext.：" + txtExt.Text:"") + "<BR><BR>";
                    strReportToDonations = strReportToDonations + "Tel：" + txtPhoneCode.Text + "-" + txtCorpPhone.Text + (txtExt.Text != "" ? " Ext.：" + txtExt.Text : "") + "<BR><BR>";
                }
                //if (txtPhoneCode.Text != gPhoneCode)
                //{
                //    strReport = strReport + "電話區碼：" + txtPhoneCode.Text + "<BR><BR>";
                //    strReportToDonations = strReportToDonations + "電話區碼：" + txtPhoneCode.Text + "<BR><BR>";
                //}
                //if (txtCorpPhone.Text != gCorpPhone)
                //{
                //    strReport = strReport + "電話：(" + txtPhoneCode.Text + ")" + txtCorpPhone.Text + "<BR><BR>";
                //    strReportToDonations = strReportToDonations + "電話：(" + txtPhoneCode.Text + ")" + txtCorpPhone.Text + "<BR><BR>";
                //}
                //if (txtExt.Text != gtExt)
                //{
                //    strReport = strReport + "電話分機：" + txtExt.Text + "<BR><BR>";
                //    strReportToDonations = strReportToDonations + "電話分機：" + txtExt.Text + "<BR><BR>";
                //}
                if (txtCellPhone.Text != gCellPhone)
                {
                    strReport = strReport + "CellPhone：" + txtCellPhone.Text + "<BR><BR>";
                    strReportToDonations = strReportToDonations + "CellPhone：" + txtCellPhone.Text + "<BR><BR>";
                }
                // 2014/7/8 傳真區碼+傳真
                if (txtFaxCode.Text + txtFax.Text != gFaxCode + gFax) strReport = strReport + "Fax：" + txtFaxCode.Text + "-" + txtFax.Text + "<BR><BR>";
                //if (txtFaxCode.Text != gFaxCode) strReport = strReport + "傳真區碼：" + txtFaxCode.Text + "<BR><BR>";
                //if (txtFax.Text != gFax) strReport = strReport + "傳真：" + txtFax.Text + "<BR><BR>"; if (Util.GetControlValue("RDO_Subscribe") != gSendNews) strReport = strReport + "訂閱月刊：" + Util.GetControlValue("RDO_Subscribe") + "<BR><BR>";
                //if (Util.GetControlValue("RDO_IsFdc") != gIsFdc) strReport = strReport + "上傳給國稅局：" + gIsFdc + "-->更改為：" + Util.GetControlValue("RDO_IsFdc") + "<BR>";
                if (txtToGoodTV.Text != gToGoodTV) strReport = strReport + "Words to GOOD TV：" + txtToGoodTV.Text + "<BR>";

                if (strReport != "")  //若有異動資料則發信
                {


                    Mailhead = "Dear " + txtDonorName.Text + " <BR><BR>You have recently made change(s) to your personal information:<BR><BR>";
                    strReport += "<BR><BR><BR>";
                    strReport += "Feel free to call us if you have questions concering your change(s).TEL: 02-8024-3911<BR>";
                    strReport += "May God bless you.<BR>";
                    strReport += "Yours sincerely,<BR>"; 
                    strReport += "GOODTV<BR>";

                    SendEMailObject MailObject = new SendEMailObject();
                    //MailObject.SmtpServer = SmtpServer;
                    MailSubject = "GOODTV - you have just made a change";
                    MailBody = Mailhead + strReport;
                    MailTo = txtEmail.Text;

                    MailFrom = System.Configuration.ConfigurationSettings.AppSettings["MailFrom"];
                    string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, MailBody);
                    if (MailObject.ErrorCode != 0)
                    {
                        this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
                    }

                }

                if (strReportToDonations != "")  //若有指定欄位異動則發內部信
                {


                    Mailhead = txtDonorName.Text + " 異動個人資料如下<BR/><BR/>";

                    SendEMailObject MailObject = new SendEMailObject();
                    //MailObject.SmtpServer = SmtpServer;
                    MailSubject = "GOODTV線上奉獻_異動通知";
                    MailBody = Mailhead + strReportToDonations;
                    MailTo = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];

                    MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
                    string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, MailBody);
                    if (MailObject.ErrorCode != 0)
                    {
                        this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
                    }

                }

            }
        }
        else
        {
            string strBody = "";
            strBody = "親愛的 " + txtDonorName.Text + " 平安！<BR><BR>您的GOODTV帳號已於 " + strCreateDateTime + " 註冊成功，以下是您的帳號及密碼, 請妥善保管喔!<BR>";
            strBody += "您的帳號：" + txtEmail.Text + "<BR>";
            strBody += "您的密碼：" + txtPassword.Text + "<BR>";
            strBody += "願神祝福您!<BR><BR>GOODTV捐款服務組敬上";
            
            SendEMailObject MailObject = new SendEMailObject();
            //MailObject.SmtpServer = SmtpServer;
            MailSubject = " GOODTV線上奉獻註冊成功通知信";
            MailBody = strBody;
            MailTo = txtEmail.Text;
            MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
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
        //20140820 清空所有Textbox資料、Session by詩儀
        foreach (object ctrl in Page.Controls)
        {
            if (ctrl is System.Web.UI.HtmlControls.HtmlForm)
            {
                System.Web.UI.HtmlControls.HtmlForm form = (System.Web.UI.HtmlControls.HtmlForm)ctrl;
                foreach (object subctrl in form.Controls)
                {
                    if (subctrl is System.Web.UI.WebControls.TextBox)
                    {
                        TextBox textctrl = (TextBox)subctrl;
                        textctrl.Text = "";
                    }
                }
            }
        }
        Session["DonorID"] = null;
        Response.Redirect("CheckOut.aspx?AddNewChecked=Y");
        
    }
    //檢查Email是否重複 20140611新增 by 詩儀
    protected void txtIs_Email_Changed(object sender, EventArgs e)
    {
        string strSql = "";

        DataTable dt = null;
        Dictionary<string, object> dict = new Dictionary<string, object>();
        strSql = "select * from DONOR where Email=@Email and Donor_Pwd Is not NULL and DeleteDate is null";

        dict.Add("Email", txtEmail.Text.Trim());
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count != 0)
        {
            this.Page.RegisterStartupScript("s", "<script>alert('此Email已有人使用，請填寫其他Email！！');</script>");
            return;
        }
    }

    protected void btnForgotPassword(object sender, EventArgs e)
    {
        Response.Redirect("ResetPassword.aspx");

    }

}
