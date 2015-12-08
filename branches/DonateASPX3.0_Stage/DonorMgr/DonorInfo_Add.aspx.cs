using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class DonorMgr_DonorInfo_Add : BasePage
{

    bool flag = false;
    // 2014/4/8 增加捐款人編號變數，可在導向到收據維護紀錄【新增】頁面時帶入參數的變數
    string Donor_Id;
    string ShowStrMsg = "";

    protected void Page_Load(object sender, EventArgs e)
    {
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            LoadDropDownListData();
            lblCheckBoxList.Text = LoadCheckBoxListData();
            // 2014/5/22 修改若輸入「捐款人」查無此捐款人資料，按「新增」按鈕後，應可直接帶入「捐款人」和「收據抬頭」欄位。
            txtDonor_Name.Text = Util.GetQueryString("donor_name");
            tbxInvoice_Title.Text = txtDonor_Name.Text;
        }
        else//20140513 新增 by Ian_Kao 儲存原本以勾選的身分別選項 以防遺失
        {
            lblCheckBoxList.Text = LoadCheckBoxListData(Util.GetControlValue("GN_DonorType"));
        }
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //性別
        // 2014/4/7 加入排序
        //Util.FillDropDownList(ddlSex, Util.GetDataTable("CaseCode", "GroupName", "性別", "", ""), "CaseName", "CaseID", false);
        Util.FillDropDownList(ddlSex, Util.GetDataTable("CaseCode", "GroupName", "性別", "CaseID", ""), "CaseName", "CaseID", false);
        ddlSex.Items.Insert(0, new ListItem("", ""));
        ddlSex.SelectedIndex = 0;

        //稱謂
        // 2014/4/7 加入排序
        //Util.FillDropDownList(ddlTitle, Util.GetDataTable("CaseCode", "GroupName", "稱謂", "", ""), "CaseName", "CaseID", false);
        Util.FillDropDownList(ddlTitle, Util.GetDataTable("CaseCode", "GroupName", "稱謂", "CaseID", ""), "CaseName", "CaseID", false);
        ddlTitle.Items.Insert(0, new ListItem("", ""));
        ddlTitle.SelectedIndex = 3;

        //20140509 註解
        ////身分別  2014/4/18新增
        //Util.FillCheckBoxList(cblDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "", ""), "CaseName", "CaseName", false);
        //cblDonor_Type.Items[0].Selected = false;

        //縣市
        Util.FillDropDownList(ddlCity, Util.GetDataTable("CodeCity", "ParentCityID", "0", "", ""), "Name", "ZipCode", false);
        ddlCity.Items.Insert(0, new ListItem("縣 市", "縣 市"));
        ddlCity.SelectedIndex = 0;

        //鄉鎮市區
        //if (ddlArea.Items.FindByText("") == null)
        ddlArea.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea.SelectedIndex = 0;

        //段
        Util.FillDropDownList(ddlSection, Util.GetDataTable("CaseCode", "GroupName", "段別", "", ""), "CaseName", "CaseName", false);
        ddlSection.Items.Insert(0, new ListItem("", ""));
        ddlSection.SelectedIndex = 0;

        //弄衖
        Util.FillDropDownList(ddlAlley, Util.GetDataTable("CaseCode", "GroupName", "弄衖別", "", ""), "CaseName", "CaseID", false);
        ddlAlley.Items.Insert(0, new ListItem("", ""));
        ddlAlley.SelectedIndex = 1;

        //收據開立
        Util.FillDropDownList(ddlInvoice_Type, Util.GetDataTable("CaseCode", "GroupName", "收據開立", "", ""), "CaseName", "CaseID", false);
        ddlInvoice_Type.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Type.SelectedIndex = 2;

        //收據稱謂
        // 2014/4/9 加入排序
        //Util.FillDropDownList(ddlTitle2, Util.GetDataTable("CaseCode", "GroupName", "稱謂", "", ""), "CaseName", "CaseID", false);
        Util.FillDropDownList(ddlTitle2, Util.GetDataTable("CaseCode", "GroupName", "稱謂", "CaseID", ""), "CaseName", "CaseID", false);
        ddlTitle2.Items.Insert(0, new ListItem("", ""));
        ddlTitle2.SelectedIndex = 0;

        //縣市(收據地址)
        Util.FillDropDownList(ddlInvoice_City, Util.GetDataTable("CodeCity", "ParentCityID", "0", "", ""), "Name", "ZipCode", false);
        ddlInvoice_City.Items.Insert(0, new ListItem("縣 市", "縣 市"));
        ddlInvoice_City.SelectedIndex = 0;

        //鄉鎮市區(收據地址)
        ddlInvoice_Area.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlInvoice_Area.SelectedIndex = 0;

        //段(收據地址)
        Util.FillDropDownList(ddlInvoice_Section, Util.GetDataTable("CaseCode", "GroupName", "段別", "", ""), "CaseName", "CaseName", false);
        ddlInvoice_Section.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Section.SelectedIndex = 0;

        //弄衖(收據地址)
        Util.FillDropDownList(ddlInvoice_Alley, Util.GetDataTable("CaseCode", "GroupName", "弄衖別", "", ""), "CaseName", "CaseID", false);
        ddlInvoice_Alley.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Alley.SelectedIndex = 1;

        //徵信錄原則
        //Util.FillDropDownList(ddlIsAnonymous, Util.GetDataTable("CaseCode", "GroupName", "徵信錄", "", ""), "CaseName", "CaseID", false);
        //ddlIsAnonymous.SelectedIndex = 1;

        //國家 20150617新增
        Util.FillDropDownList(ddlOverseasCountry, Util.GetDataTable("CaseCode", "GroupName", "國家", "CaseID", ""), "CaseName", "CaseName", false);
        ddlOverseasCountry.Items.Insert(0, new ListItem("", ""));
        ddlOverseasCountry.SelectedIndex = 0;

        Util.FillDropDownList(ddlInvoice_OverseasCountry, Util.GetDataTable("CaseCode", "GroupName", "國家", "CaseID", ""), "CaseName", "CaseName", false);
        ddlInvoice_OverseasCountry.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_OverseasCountry.SelectedIndex = 0;
    }
    //----------------------------------------------------------------------
    //20140509 新增 by Ian_Kao
    public string LoadCheckBoxListData()
    {
        string strSql = "select CaseName from CaseCode where CodeType = 'DonorType' order by CaseID";
        DataTable dt = NpoDB.GetDataTableS(strSql, null);

        List<ControlData> list = new List<ControlData>();
        list.Clear();
        foreach (DataRow dr in dt.Rows)
        {
            string CaseName = dr["CaseName"].ToString();
            bool ShowBR = false;
            list.Add(new ControlData("Checkbox", "GN_DonorType", "CB_DonorType", CaseName, CaseName, ShowBR, ""));
        }
        return HtmlUtil.RenderControl(list);
    }
    //----------------------------------------------------------------------
    //20140513 新增 by Ian_Kao 儲存原本以勾選的身分別選項 以防遺失
    public string LoadCheckBoxListData(string DonorTypeValues)
    {
        string strSql = "select CaseName from CaseCode where CodeType = 'DonorType' order by CaseID";
        DataTable dt = NpoDB.GetDataTableS(strSql, null);

        List<ControlData> list = new List<ControlData>();
        list.Clear();
        foreach (DataRow dr in dt.Rows)
        {
            string CaseName = dr["CaseName"].ToString();
            bool ShowBR = false;
            list.Add(new ControlData("Checkbox", "GN_DonorType", "CB_DonorType", CaseName, CaseName, ShowBR, DonorTypeValues));
        }
        return HtmlUtil.RenderControl(list);
    }
    //----------------------------------------------------------------------
    public void CityToArea(DropDownList ddlTO, DropDownList ddlFrom)
    {
        Util.FillDropDownList(ddlTO, Util.GetDataTable("CodeCity", "ParentCityID", ddlFrom.SelectedValue, "", ""), "Name", "ZipCode", false);
        ddlTO.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlTO.SelectedIndex = 0;
    }
    //----------------------------------------------------------------------
    //dropdownlist選縣市自動帶入地區和電話區碼
    protected void ddlCity_SelectedIndexChanged(object sender, EventArgs e)
    {
        CityToArea(ddlArea, ddlCity);
        //20140127增加縣市與電話區碼連動
        if (ddlCity.SelectedItem.Text == "基隆市" || ddlCity.SelectedItem.Text == "台北市" || ddlCity.SelectedItem.Text == "新北市")
        {
            tbxTel_Office_Loc.Text = "02";
            tbxFax_Loc.Text = "02";
        }
        else if (ddlCity.SelectedItem.Text == "桃園縣" || ddlCity.SelectedItem.Text == "新竹縣" || ddlCity.SelectedItem.Text == "新竹市" || ddlCity.SelectedItem.Text == "宜蘭縣" || ddlCity.SelectedItem.Text == "花蓮縣")
        {
            tbxTel_Office_Loc.Text = "03";
            tbxFax_Loc.Text = "03";
        }
        else if (ddlCity.SelectedItem.Text == "苗栗縣")
        {
            tbxTel_Office_Loc.Text = "037";
            tbxFax_Loc.Text = "037";
        }
        else if (ddlCity.SelectedItem.Text == "台中市" || ddlCity.SelectedItem.Text == "彰化縣")
        {
            tbxTel_Office_Loc.Text = "04";
            tbxFax_Loc.Text = "04";
        }
        else if (ddlCity.SelectedItem.Text == "南投縣")
        {
            tbxTel_Office_Loc.Text = "049";
            tbxFax_Loc.Text = "049";
        }
        else if (ddlCity.SelectedItem.Text == "雲林縣" || ddlCity.SelectedItem.Text == "嘉義市" || ddlCity.SelectedItem.Text == "嘉義縣")
        {
            tbxTel_Office_Loc.Text = "05";
            tbxFax_Loc.Text = "05";
        }
        else if (ddlCity.SelectedItem.Text == "台南市" || ddlCity.SelectedItem.Text == "澎湖縣")
        {
            tbxTel_Office_Loc.Text = "06";
            tbxFax_Loc.Text = "06";
        }
        else if (ddlCity.SelectedItem.Text == "高雄市")
        {
            tbxTel_Office_Loc.Text = "07";
            tbxFax_Loc.Text = "07";
        }
        else if (ddlCity.SelectedItem.Text == "屏東縣")
        {
            tbxTel_Office_Loc.Text = "08";
            tbxFax_Loc.Text = "08";
        }
        else if (ddlCity.SelectedItem.Text == "台東縣")
        {
            tbxTel_Office_Loc.Text = "089";
            tbxFax_Loc.Text = "089";
        }
        else if (ddlCity.SelectedItem.Text == "金門縣")
        {
            tbxTel_Office_Loc.Text = "082";
            tbxFax_Loc.Text = "082";
        }
        else if (ddlCity.SelectedItem.Text == "連江縣")
        {
            tbxTel_Office_Loc.Text = "0836";
            tbxFax_Loc.Text = "0836";
        }
        else
        {
            tbxTel_Office_Loc.Text = "";
            tbxFax_Loc.Text = "";
        }
    }
    //----------------------------------------------------------------------
    //dropdownlist選地區自動填入郵遞區號
    protected void ddlArea_SelectedIndexChanged(object sender, EventArgs e)
    {
        tbxZipCode.Text = Util.GetCityCode(ddlArea.SelectedItem.Text, ddlCity.SelectedValue).Substring(0, 3);
    }
    //----------------------------------------------------------------------
    //dropdownlist選城市自動帶入地區
    protected void ddlInvoice_City_SelectedIndexChanged(object sender, EventArgs e)
    {
        CityToArea(ddlInvoice_Area,ddlInvoice_City);
    }
    //----------------------------------------------------------------------
    //dropdownlist選地區自動填入郵遞區號
    protected void ddlInvoice_Area_SelectedIndexChanged(object sender, EventArgs e)
    {
        tbxInvoice_ZipCode.Text = Util.GetCityCode(ddlInvoice_Area.SelectedItem.Text, ddlInvoice_City.SelectedValue).Substring(0,3);
    }
    //----------------------------------------------------------------------
    //台灣本島和海外地址擇一
    protected void cbxIsLocal_CheckedChanged(object sender, EventArgs e)
    {
        if (cbxIsLocal.Checked == true && cbxIsAbroad.Checked == true)
        {
            cbxIsAbroad.Checked = false;
        }
    }
    //----------------------------------------------------------------------
    protected void cbxIsAbroad_CheckedChanged(object sender, EventArgs e)
    {
        if (cbxIsAbroad.Checked == true && cbxIsLocal.Checked == true)
        {
            cbxIsLocal.Checked = false;
        }
        if (cbxIsAbroad.Checked == true)
        {
            // 2014/6/24 需求修改
            tbxIsSendNewsNum.Text = "";
            cbxIsDVD.Checked = false;
            cbxIsSendEpaper.Checked = false;
            cbxIsGift.Checked = false;
            cbxIsBigAmtThank.Checked = false;
            cbxIsPost.Checked = false;
        }

    }
    //----------------------------------------------------------------------
    //台灣本島和海外地址擇一(收據)
    protected void cbxInvoice_IsLocal_CheckedChanged(object sender, EventArgs e)
    {
        if (cbxInvoice_IsLocal.Checked == true && cbxInvoice_IsAbroad.Checked == true)
        {
            cbxInvoice_IsAbroad.Checked = false;
        }
    }
    protected void cbxInvoice_IsAbroad_CheckedChanged(object sender, EventArgs e)
    {
        if (cbxInvoice_IsAbroad.Checked == true && cbxInvoice_IsLocal.Checked == true)
        {
            cbxInvoice_IsLocal.Checked = false;
        }
    }
    //----------------------------------------------------------------------
    //勾選或取消勾選收據資料同上
    protected void cbxInvoice_same_CheckedChanged(object sender, EventArgs e)
    {
        if (cbxInvoice_same.Checked)
        {
            if (cbxIsLocal.Checked)
            {
                cbxInvoice_IsLocal.Checked = true;
                tbxInvoice_ZipCode.Text = tbxZipCode.Text;

                ddlInvoice_City.SelectedIndex = ddlCity.SelectedIndex;
                if (ddlInvoice_City.Items.FindByText("縣 市") == null)
                {
                    ddlInvoice_City.Items.Insert(0, new ListItem("縣 市", "縣 市"));
                }

                ddlInvoice_City.SelectedValue = ddlCity.SelectedValue;
                CityToArea(ddlInvoice_Area, ddlInvoice_City);
                ddlInvoice_Area.SelectedIndex = ddlArea.SelectedIndex;

                tbxInvoice_Street.Text = tbxStreet.Text;
                ddlInvoice_Section.SelectedIndex = ddlSection.SelectedIndex;
                tbxInvoice_Lane.Text = tbxLane.Text;
                tbxInvoice_Alley0.Text = tbxAlley.Text;
                ddlInvoice_Alley.SelectedIndex = ddlAlley.SelectedIndex;
                tbxInvoice_No1.Text = tbxNo1.Text;
                tbxInvoice_No2.Text = tbxNo2.Text;
                tbxInvoice_Floor1.Text = tbxFloor1.Text;
                tbxInvoice_Floor2.Text = tbxFloor2.Text;
                tbxInvoice_Room.Text = tbxRoom.Text;
                tbxInvoice_Attn.Text = tbxAttn.Text;

                //海外地址全部清空
                cbxInvoice_IsAbroad.Checked = false;
                ddlInvoice_OverseasCountry.SelectedIndex = 0;
                tbxInvoice_OverseasAddress.Text = "";
            }
            if (cbxIsAbroad.Checked)
            {
                cbxInvoice_IsAbroad.Checked = true;
                //tbxInvoice_OverseasCountry.Text = tbxOverseasCountry.Text;
                ddlInvoice_OverseasCountry.SelectedValue = ddlOverseasCountry.SelectedValue;
                tbxInvoice_OverseasAddress.Text = tbxOverseasAddress.Text;

                //台灣本島全部清空
                cbxInvoice_IsLocal.Checked = false;
                tbxInvoice_ZipCode.Text = "";
                ddlInvoice_City.SelectedIndex = 0;
                if (ddlInvoice_City.Items.FindByText("縣 市") == null)
                {
                    ddlInvoice_City.Items.Insert(0, new ListItem("縣 市", "縣 市"));
                }
                ddlInvoice_City.SelectedIndex = 0;
                ddlInvoice_Area.SelectedIndex = 0;
                tbxInvoice_Street.Text = "";
                ddlInvoice_Section.SelectedIndex = 0;
                tbxInvoice_Lane.Text = "";
                tbxInvoice_Alley0.Text = "";
                ddlInvoice_Alley.SelectedIndex = 0;
                tbxInvoice_No1.Text = "";
                tbxInvoice_No2.Text = "";
                tbxInvoice_Floor1.Text = "";
                tbxInvoice_Floor2.Text = "";
                tbxInvoice_Room.Text = "";
                tbxInvoice_Attn.Text = "";
            }

            // 2014/4/7 修改收據稱謂欄位可帶入捐款人稱謂欄位值
            ddlTitle2.SelectedItem.Text = ddlTitle.SelectedItem.Text;
            //2014/5/21 修改可帶入收據抬頭
            tbxInvoice_Title.Text = txtDonor_Name.Text;
        }
    }
    //----------------------------------------------------------------------
    protected void btnAdd_Click(object sender, EventArgs e)
    {

        //20140509 修改 by Ian_Kao
        ////20140425 新增 by Ian_Kao 另外新增驗證項來驗證是否身分別有填入
        //if (!IsValid)
        //{
        //    AjaxShowMessage("身分別欄位不得為空白");
        //    return;
        //}

        // 2014/6/25 移到前端處理 by Samson
        //if (Util.GetControlValue("GN_DonorType") == "")
        //{
        //    AjaxShowMessage("身分別欄位不得為空白");
        //    return;
        //}

        try
        {

            Donor_AddNew();
            // 2014/4/9 修改顯示訊息執行順序
            //SetSysMsg("新增資料成功");
            //flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }

        if (flag == true)
        {

            // 2014/4/9 修改用另外的顯示訊息函式
            AjaxShowMessage("捐款人資料已建立您可以新增捐款資料！");
            // 2014/4/7 新增後導向到收據維護紀錄【新增】頁面
            //Response.Redirect(Util.RedirectByTime("DonorInfo.aspx"));
            Response.Redirect(Util.RedirectByTime("../DonateMgr/Donate_Add.aspx", "Donor_Id=" + Donor_Id));

        }
    }

    //----------------------------------------------------------------------
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("DonorInfo.aspx"));
    }
    //----------------------------------------------------------------------
    public void Donor_AddNew()
    {
        string strSql = "insert into  DONOR\n";
        strSql += "( Donor_Name, Sex, Title, Donor_Type, IDNo, Birthday, Cellular_Phone, Tel_Office_Loc, Tel_Office, Tel_Office_Ext,\n";
        strSql += "  Fax_Loc, Fax, Email,Contactor, IsAbroad, ZipCode, City, Area, Address,\n";
        strSql += " Invoice_Type, Title2, Invoice_Title,Invoice_IDNo, IsAbroad_Invoice, IsAnonymous, Remark, ToGOODTV, IsFdc, \n";
        strSql += " Street,Section,Lane,Alley,HouseNo,HouseNoSub,Floor,FloorSub,Room,OverseasCountry,\n";
        strSql += " OverseasAddress,Invoice_Street,Invoice_Section,Invoice_Lane,Invoice_Alley,Invoice_HouseNo,\n";
        strSql += " Invoice_HouseNoSub,Invoice_Floor,Invoice_FloorSub,Invoice_Room,Invoice_OverseasCountry,\n";
        strSql += " Invoice_OverseasAddress,Invoice_ZipCode,Invoice_City,Invoice_Area,Invoice_Address, Dept_Id, IsMember,  \n";
        strSql += " IsSendNews, IsSendNewsNum, IsDVD, IsSendEpaper, IsGift, IsBigAmtThank, IsPost, IsContact, IsErrAddress, Attn, Invoice_Attn, \n";
        strSql += " Donate_No, Donate_Total, Donate_NoD, Donate_TotalD, Donate_NoC, Donate_TotalC, \n";
        strSql += " Donate_NoM, Donate_TotalM, Donate_NoS, Donate_TotalS, Donate_NoA, Donate_TotalA, Donate_NoND, Donate_TotalND, \n";
        strSql += " Create_Date, Create_DateTime, Create_User, Create_IP) values\n";
        strSql += "(@Donor_Name,@Sex,@Title,@Donor_Type,@IDNo,@Birthday,@Cellular_Phone,@Tel_Office_Loc,@Tel_Office,@Tel_Office_Ext,\n";
        // 2014/4/9 修改SQL 欄位指定錯誤
        strSql += "@Fax_Loc,@Fax,@Email,@Contactor,@IsAbroad,@ZipCode,@City,@Area,@Address,@Invoice_Type,@Title2,@Invoice_Title,@Invoice_IDNo, \n";
        strSql += "@IsAbroad_Invoice,@IsAnonymous,@Remark,@ToGOODTV,@IsFdc,@Street,@Section,@Lane,@Alley,@HouseNo,@HouseNoSub,@Floor, \n";
        //strSql += "@Fax_Loc,@Fax,@Email,@Contactor,@IsAbroad,@ZipCode,@City,@Area,@Address,@Invoice_Type,@Title2,@Invoice_Title,@IsAbroad_Invoice, \n";
        //strSql += "@IsAnonymous,@Remark,@ToGOODTV,@IsFdc,@Street,@Section,@Lane,@Alley,@HouseNo,@HouseNoSub,@Floor, \n";
        strSql += "@FloorSub,@Room,@OverseasCountry,@OverseasAddress,@Invoice_Street,@Invoice_Section,@Invoice_Lane,@Invoice_Alley, \n";
        strSql += "@Invoice_HouseNo,@Invoice_HouseNoSub,@Invoice_Floor,@Invoice_FloorSub,@Invoice_Room,@Invoice_OverseasCountry,@Invoice_OverseasAddress, \n";
        strSql += "@Invoice_ZipCode,@Invoice_City,@Invoice_Area,@Invoice_Address,@Dept_Id, @IsMember, \n";
        strSql += "@IsSendNews,@IsSendNewsNum,@IsDVD,@IsSendEpaper,@IsGift,@IsBigAmtThank,@IsPost,@IsContact,@IsErrAddress,@Attn,@Invoice_Attn, \n";
        strSql += "@Donate_No,@Donate_Total,@Donate_NoD,@Donate_TotalD,@Donate_NoC,@Donate_TotalC, \n";
        strSql += "@Donate_NoM,@Donate_TotalM,@Donate_NoS,@Donate_TotalS,@Donate_NoA,@Donate_TotalA,@Donate_NoND,@Donate_TotalND, \n";
        strSql += "@Create_Date,@Create_DateTime,@Create_User,@Create_IP)";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Name", txtDonor_Name.Text.Trim());
        dict.Add("Sex", ddlSex.SelectedItem.Text);
        dict.Add("Title", ddlTitle.SelectedItem.Text);
        //20140509 修改 by Ian_Kao
        //dict.Add("Donor_Type", cblDonor_Type.SelectedValue);
        dict.Add("Donor_Type", Util.GetControlValue("GN_DonorType"));
        dict.Add("IDNo", tbxIDNo.Text.Trim());
        if (txtBirthday.Text.Trim() != "")
        {
            dict.Add("Birthday", DateTime.Parse(txtBirthday.Text).ToShortDateString().ToString());
        }
        else
        {
            dict.Add("Birthday", null);
        }
        dict.Add("Cellular_Phone", tbxCellular_Phone.Text.Trim());
        dict.Add("Tel_Office_Loc", tbxTel_Office_Loc.Text.Trim());
        dict.Add("Tel_Office", tbxTel_Office.Text.Trim());
        dict.Add("Tel_Office_Ext", tbxTel_Office_Ext.Text.Trim());
        dict.Add("Fax_Loc", tbxFax_Loc.Text.Trim());
        dict.Add("Fax", tbxFax.Text.Trim());
        dict.Add("Email", tbxEMail.Text.Trim());
        dict.Add("Contactor", tbxContactor.Text.Trim());
        //dict.Add("JobTitle", tbxJobTitle.Text.Trim());
        //國內
        if (this.cbxIsLocal.Checked)
        {
            dict.Add("IsAbroad", "N");
            //地址
            dict.Add("ZipCode", tbxZipCode.Text.Trim());
            if (ddlCity.SelectedIndex != 0)
            {
                dict.Add("City", ddlCity.SelectedValue);
            }
            else
            {
                dict.Add("City", "");
            }
            if (ddlArea.SelectedIndex != 0)
            {
                dict.Add("Area", ddlArea.SelectedValue);
            }
            else
            {
                dict.Add("Area", "");
            }
            string Address = "";
            if (tbxStreet.Text.Trim() != "")
            {
                Address += tbxStreet.Text.Trim();
                dict.Add("Street", tbxStreet.Text.Trim());
            }
            else
            {
                dict.Add("Street", "");
            }
            if (ddlSection.SelectedIndex != 0)//有無段別
            {
                Address += ddlSection.SelectedItem.Text + "段";
                dict.Add("Section", ddlSection.SelectedItem.Text);
            }
            else
            {
                dict.Add("Section", "");
            }
            if (tbxLane.Text.Trim() != "")//有無巷
            {
                Address += tbxLane.Text.Trim() + "巷";
                dict.Add("Lane", tbxLane.Text.Trim());
            }
            else
            {
                dict.Add("Lane", "");
            }
            if (tbxAlley.Text.Trim() != "")//有無弄衖
            {
                Address += tbxAlley.Text.Trim() + ddlAlley.SelectedItem.Text;
                dict.Add("Alley", tbxAlley.Text.Trim());
            }
            else
            {
                dict.Add("Alley", "");
            }
            //2015/08/12 更正 <No1>號之<No2>
            if (tbxNo1.Text.Trim() != "")//有無號碼
            {
                Address += tbxNo1.Text.Trim() + "號";
                dict.Add("HouseNo", tbxNo1.Text.Trim());
            }
            else
            {
                dict.Add("HouseNo", "");
            }
            if (tbxNo2.Text.Trim() != "")//之幾
            {
                Address += "之" + tbxNo2.Text.Trim() + " ";
                dict.Add("HouseNoSub", tbxNo2.Text.Trim());
            }
            else
            {
                dict.Add("HouseNoSub", "");
            }
            // 2014/4/8 更正 <Floor1>-<Floor2>樓 --> <Floor1>樓之<Floor2>
            if (tbxFloor1.Text.Trim() != "" && tbxFloor2.Text.Trim() == "")//有無樓別
            {
                Address += tbxFloor1.Text.Trim() + "樓";
                dict.Add("Floor", tbxFloor1.Text.Trim());
                dict.Add("FloorSub", "");
            }
            else if (tbxFloor1.Text.Trim() == "" && tbxFloor2.Text.Trim() != "")
            {
                //Address += tbxFloor2.Text.Trim() + "樓";
                Address += "樓之" + tbxFloor2.Text.Trim() ;

                dict.Add("Floor", "");
                dict.Add("FloorSub", tbxFloor2.Text.Trim());
            }
            else if (tbxFloor1.Text.Trim() != "" && tbxFloor2.Text.Trim() != "")
            {
                //Address += tbxFloor1.Text.Trim() + "-" + tbxFloor2.Text.Trim() + "樓";
                Address += tbxFloor1.Text.Trim() + "樓之" + tbxFloor2.Text.Trim();

                dict.Add("Floor", tbxFloor1.Text.Trim());
                dict.Add("FloorSub", tbxFloor2.Text.Trim());
            }
            else
            {
                dict.Add("Floor", "");
                dict.Add("FloorSub", "");
            }
            if (tbxRoom.Text.Trim() != "")//有無室
            {
                //Address += tbxRoom.Text.Trim() + "室";
                Address += " " + tbxRoom.Text.Trim() + "室";
                dict.Add("Room", tbxRoom.Text.Trim());
            }
            else
            {
                dict.Add("Room", "");
            }
            if (tbxAttn.Text.Trim() != "")//有無Attn
            {
                dict.Add("Attn", tbxAttn.Text.Trim());
            }
            else
            {
                dict.Add("Attn", "");
            }
            dict.Add("Address", Address);
            dict.Add("OverseasCountry", "");
            dict.Add("OverseasAddress", "");
        }
        //國外
        if (this.cbxIsAbroad.Checked)
        {
            dict.Add("IsAbroad", "Y");
            //地址
            dict.Add("ZipCode", "");
            dict.Add("City", "");
            dict.Add("Area", "");
            string Address = "";
            if (tbxOverseasAddress.Text.Trim() != "")//有無其他
            {
                Address += tbxOverseasAddress.Text.Trim()+ " ";
                dict.Add("OverseasAddress", tbxOverseasAddress.Text.Trim() + " ");
            }
            else
            {
                dict.Add("OverseasAddress", "");
            }
            if (ddlOverseasCountry.SelectedIndex != 0)//有無國家/省城市/區
            {
                Address += ddlOverseasCountry.SelectedValue;
                dict.Add("OverseasCountry", ddlOverseasCountry.SelectedValue);
            }
            else
            {
                dict.Add("OverseasCountry", "");
            }
            dict.Add("Address", Address);
            dict.Add("Street", "");
            dict.Add("Section", "");
            dict.Add("Lane", "");
            dict.Add("Alley", "");
            dict.Add("HouseNo", "");
            dict.Add("HouseNoSub", "");
            dict.Add("Floor", "");
            dict.Add("FloorSub", "");
            dict.Add("Room", "");
            dict.Add("Attn", "");

        }
        //文宣品
        dict.Add("IsSendNewNum", tbxIsSendNewsNum.Text.Trim() != "" ? tbxIsSendNewsNum.Text.Trim() : "0");
        if (tbxIsSendNewsNum.Text.Trim() == "" || tbxIsSendNewsNum.Text.Trim() == "0")
        {
            dict.Add("IsSendNews","N");
            dict.Add("IsSendNewsNum", "0");
        }
        else
        {
            dict.Add("IsSendNews","Y");
            dict.Add("IsSendNewsNum", tbxIsSendNewsNum.Text.Trim());
        }
        if(cbxIsDVD.Checked)
        {
            dict.Add("IsDVD", "Y");
        }
        else
        {
            dict.Add("IsDVD", "N");
        }
        if (cbxIsSendEpaper.Checked)
        {
            dict.Add("IsSendEpaper", "Y");
        }
        else
        {
            dict.Add("IsSendEpaper", "N");
        }
        if (cbxIsGift.Checked)
        {
            dict.Add("IsGift", "Y");
        }
        else
        {
            dict.Add("IsGift", "N");
        }
        if (cbxIsBigAmtThank.Checked)
        {
            dict.Add("IsBigAmtThank", "Y");
        }
        else
        {
            dict.Add("IsBigAmtThank", "N");
        }
        //20150825增加欄位
        if (cbxIsPost.Checked)
        {
            dict.Add("IsPost", "Y");
        }
        else
        {
            dict.Add("IsPost", "N");
        }
        if (cbxIsContact.Checked)
        {
            dict.Add("IsContact", "N");
        }
        else
        {
            dict.Add("IsContact", "Y");
        }
        if (cbxIsErrAddress.Checked)
        {
            dict.Add("IsErrAddress", "Y");
        }
        else
        {
            dict.Add("IsErrAddress", "N");
        }
        dict.Add("IsMember", "N");
        dict.Add("Dept_Id", SessionInfo.DeptID);
        dict.Add("Invoice_Type", ddlInvoice_Type.SelectedItem.Text);
        //收據地址
        //國內
        if (this.cbxInvoice_IsLocal.Checked)
        {
            dict.Add("IsAbroad_Invoice", "N");
            dict.Add("Invoice_ZipCode", tbxInvoice_ZipCode.Text.Trim());
            if (ddlInvoice_City.SelectedIndex != 0)
            {
                dict.Add("Invoice_City", ddlInvoice_City.SelectedValue);
            }
            else
            {
                dict.Add("Invoice_City", "");
            }
            if (ddlInvoice_Area.SelectedIndex != 0)
            {
                dict.Add("Invoice_Area", ddlInvoice_Area.SelectedValue);
            }
            else
            {
                dict.Add("Invoice_Area", "");
            }

            string Address = "";
            if (tbxInvoice_Street.Text.Trim() != "")//有無街/大道
            {
                Address += tbxInvoice_Street.Text.Trim();
                dict.Add("Invoice_Street", tbxInvoice_Street.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_Street", "");
            }
            if (ddlInvoice_Section.SelectedIndex != 0)//有無段別
            {
                Address += ddlInvoice_Section.SelectedItem.Text + "段";
                dict.Add("Invoice_Section", ddlInvoice_Section.SelectedItem.Text);
            }
            else
            {
                dict.Add("Invoice_Section", "");
            }
            if (tbxInvoice_Lane.Text.Trim() != "")//有無巷
            {
                Address += tbxInvoice_Lane.Text.Trim() + "巷";
                dict.Add("Invoice_Lane", tbxInvoice_Lane.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_Lane", "");
            }
            if (tbxInvoice_Alley0.Text.Trim() != "")//有無弄衖
            {
                Address += tbxInvoice_Alley0.Text.Trim() + ddlInvoice_Alley.SelectedItem.Text;
                dict.Add("Invoice_Alley", tbxInvoice_Alley0.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_Alley", "");
            }
            //2015/08/12 更正 <No1>號之<No2>
            if (tbxInvoice_No1.Text.Trim() != "" )//有無號碼
            {
                Address += tbxInvoice_No1.Text.Trim() + "號";
                dict.Add("Invoice_HouseNo", tbxInvoice_No1.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_HouseNo", "");
            }
            if (tbxInvoice_No2.Text.Trim() != "")
            {
                Address += "之" + tbxInvoice_No2.Text.Trim() + " ";
                dict.Add("Invoice_HouseNoSub", tbxInvoice_No2.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_HouseNoSub", "");
            }
            // 2014/4/8 更正 <Floor1>-<Floor2>樓 --> <Floor1>樓之<Floor2>
            if (tbxInvoice_Floor1.Text.Trim() != "" && tbxInvoice_Floor2.Text.Trim() == "")//有無樓別
            {
                Address += tbxInvoice_Floor1.Text.Trim() + "樓";
                dict.Add("Invoice_Floor", tbxInvoice_Floor1.Text.Trim());
                dict.Add("Invoice_FloorSub", "");
            }
            else if (tbxInvoice_Floor1.Text.Trim() == "" && tbxInvoice_Floor2.Text.Trim() != "")
            {
                //Address += tbxInvoice_Floor2.Text.Trim() + "樓";
                Address +=  "樓之" + tbxInvoice_Floor2.Text.Trim();
                dict.Add("Invoice_Floor", "");
                dict.Add("Invoice_FloorSub", tbxInvoice_Floor2.Text.Trim());
            }
            else if (tbxInvoice_Floor1.Text.Trim() != "" && tbxInvoice_Floor2.Text.Trim() != "")
            {
                //Address += tbxInvoice_Floor1.Text.Trim() + "-" + tbxInvoice_Floor2.Text.Trim() + "樓";
                Address += tbxInvoice_Floor1.Text.Trim() + "樓之" + tbxInvoice_Floor2.Text.Trim();
                dict.Add("Invoice_Floor", tbxInvoice_Floor1.Text.Trim());
                dict.Add("Invoice_FloorSub", tbxInvoice_Floor2.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_Floor", "");
                dict.Add("Invoice_FloorSub", "");
            }
            if (tbxInvoice_Room.Text.Trim() != "")//有無室
            {
                //Address += tbxInvoice_Room.Text.Trim() + "室";
                Address += " " + tbxInvoice_Room.Text.Trim() + "室";
                dict.Add("Invoice_Room", tbxInvoice_Room.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_Room", "");
            }
            if (tbxInvoice_Attn.Text.Trim() != "")//有無Attn
            {
                dict.Add("Invoice_Attn", tbxInvoice_Attn.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_Attn", "");
            }
            dict.Add("Invoice_Address", Address);
            dict.Add("Invoice_OverseasCountry", "");
            dict.Add("Invoice_OverseasAddress", "");
        }
        //國外
        if (this.cbxInvoice_IsAbroad.Checked)
        {
            dict.Add("IsAbroad_Invoice", "Y");
            dict.Add("Invoice_ZipCode", "");
            dict.Add("Invoice_City", "");
            dict.Add("Invoice_Area", "");
            //地址
            string Address = "";
            if (tbxInvoice_OverseasAddress.Text.Trim() != "")//有無其他
            {
                Address += tbxInvoice_OverseasAddress.Text.Trim()+ " ";
                dict.Add("Invoice_OverseasAddress", tbxInvoice_OverseasAddress.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_OverseasAddress", "");
            }
            if (ddlInvoice_OverseasCountry.SelectedIndex != 0)//有無國家/省城市/區
            {
                Address += ddlInvoice_OverseasCountry.SelectedValue;
                dict.Add("Invoice_OverseasCountry", ddlInvoice_OverseasCountry.SelectedValue);
            }
            else
            {
                dict.Add("Invoice_OverseasCountry", "");
            }
            dict.Add("Invoice_Address", Address);
            dict.Add("Invoice_Street", "");
            dict.Add("Invoice_Section", "");
            dict.Add("Invoice_Lane", "");
            dict.Add("Invoice_Alley", "");
            dict.Add("Invoice_HouseNo", "");
            dict.Add("Invoice_HouseNoSub", "");
            dict.Add("Invoice_Floor", "");
            dict.Add("Invoice_FloorSub", "");
            dict.Add("Invoice_Room", "");
            dict.Add("Invoice_Attn", "");

        }
        //國內或國外都沒勾選
        if(cbxInvoice_IsLocal.Checked == false && cbxInvoice_IsAbroad.Checked == false)
        {
            dict.Add("IsAbroad_Invoice", "");
            dict.Add("Invoice_ZipCode", "");
            dict.Add("Invoice_Address", "");
            dict.Add("Invoice_City", "");
            dict.Add("Invoice_Area", "");
            dict.Add("Invoice_Street", "");
            dict.Add("Invoice_Section", "");
            dict.Add("Invoice_Lane", "");
            dict.Add("Invoice_Alley", "");
            dict.Add("Invoice_HouseNo", "");
            dict.Add("Invoice_HouseNoSub", "");
            dict.Add("Invoice_Floor", "");
            dict.Add("Invoice_FloorSub", "");
            dict.Add("Invoice_Room", "");
            dict.Add("Invoice_Attn", "");
            dict.Add("Invoice_OverseasCountry", "");
            dict.Add("Invoice_OverseasAddress", "");
        }
       
        dict.Add("Title2", ddlTitle2.SelectedItem.Text);
        // 2014/4/9 修改成與APS 2.0版一樣，收據抬頭無輸入資料以捐款人資料存入
        dict.Add("Invoice_Title", (tbxInvoice_Title.Text.Trim() != "" ? tbxInvoice_Title.Text.Trim() : txtDonor_Name.Text.Trim()));
        //dict.Add("Invoice_Title", tbxInvoice_Title.Text.Trim());
        dict.Add("Invoice_IDNo", tbxInvoice_IDNo.Text.Trim());
        //dict.Add("IsAnonymous", ddlIsAnonymous.SelectedItem.Text);
        if (cbxIsAnonymous.Checked)
        {
            dict.Add("IsAnonymous", "Y");
        }
        else
        {
            dict.Add("IsAnonymous", "N");
        }
        //if (ddlIsAnonymous.SelectedItem.Text == "不刊登")
        //{
        //    dict.Add("IsAnonymous", "Y");
        //}
        //else
        //{
        //    dict.Add("IsAnonymous", "N");
        //}
        dict.Add("Remark", tbxRemark.Text.Trim());
        dict.Add("ToGOODTV", tbxToGOODTV.Text.Trim());
        if (cbxIsFdc.Checked)
        {
            dict.Add("IsFdc", "Y");
        }
        else
        {
            dict.Add("IsFdc", "N");
        }
        dict.Add("Donate_No", "0");
        dict.Add("Donate_Total", "0");
        dict.Add("Donate_NoD", "0");
        dict.Add("Donate_TotalD", "0");
        dict.Add("Donate_NoC", "0");
        dict.Add("Donate_TotalC", "0");
        dict.Add("Donate_NoM", "0");
        dict.Add("Donate_TotalM", "0");
        dict.Add("Donate_NoS", "0");
        dict.Add("Donate_TotalS", "0");
        dict.Add("Donate_NoA", "0");
        dict.Add("Donate_TotalA", "0");
        dict.Add("Donate_NoND", "0");
        dict.Add("Donate_TotalND", "0");
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());

        // 2014/4/8 改成可回傳新增的捐款人編號，可在導向收據維護紀錄【新增】頁面時帶入參數，自動帶出捐款人資料
        //NpoDB.ExecuteSQLS(strSql, dict);
        Donor_Id = NpoDB.GetScalarS(strSql, dict);
        if (Donor_Id != "")
        {
            flag = true;

        }
        //ShowSysMsg("新增成功!");
        //AjaxShowMessage("捐款人資料已建立\n您可以新增捐款資料！");

    }

    //增加性別與稱謂連動  20140127
    protected void ddlSex_SelectedIndexChanged(object sender, EventArgs e)
    {

        // 2014/4/7 修改成值對值，不會因排序問題而對應錯誤。
        // 性別的資料欄位CaseID 對應 稱謂的資料欄位CaseID
        string caseSwitch = ddlSex.SelectedItem.Text;
        switch (caseSwitch)
        {
            case "男":
                caseSwitch = "先生";
                break;
            case "女":
                caseSwitch = "小姐/女士";
                break;
            case "歿":
                caseSwitch = "";
                // 2014/6/24 需求修改
                tbxIsSendNewsNum.Text = "";
                cbxIsDVD.Checked = false;
                cbxIsSendEpaper.Checked = false;
                cbxIsGift.Checked = false;
                cbxIsBigAmtThank.Checked = false;
                break;
            default:
                caseSwitch = "先生/女士";
                break;
        }
        for (int i = 0; i < ddlTitle.Items.Count; i++)
        {
            if (ddlTitle.Items[i].Text == caseSwitch)
                ddlTitle.Items[i].Selected = true;
            else
                ddlTitle.Items[i].Selected = false;
        }

    }



    //20140509 修改 by Ian_Kao 驗證放至btnEdit_Click中執行
    //20140425 新增 by Ian_Kao 另外新增驗證項來驗證是否身分別有填入
    //protected void CustomValidator1_ServerValidate(object source, ServerValidateEventArgs args)
    //{
    //    //var target = this.cblDonor_Type;
    //    //foreach (ListItem item in target.Items)
    //    //{
    //    //    if (item.Selected)
    //    //    {
    //    //        args.IsValid = true;
    //    //        return;
    //    //    }
    //    //}
    //    //args.IsValid = false;
    //}
}
