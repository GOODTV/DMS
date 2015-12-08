using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class MemberMgr_Member_Add : BasePage
{
    bool flag = false;
    // 2014/11/21 增加捐款人編號變數，可在導向到讀者管理【修改】頁面時帶入參數的變數
    string Donor_Id;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            LoadDropDownListData();
            lblCheckBoxList.Text = LoadCheckBoxListData();
            // 2014/8/13 修改若輸入「讀者姓名」查無此讀者資料，按「新增」按鈕後，應可直接帶入「讀者姓名」欄位。
            tbxDonor_Name.Text = Util.GetQueryString("Member_Name");
        }
        else//20140514 新增  儲存原本以勾選的身分別選項 以防遺失
        {
            lblCheckBoxList.Text = LoadCheckBoxListData(Util.GetControlValue("GN_DonorType"));
        }
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " ReadOnly();", true);
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //性別
        // 2014/4/11 加入排序
        //Util.FillDropDownList(ddlSex, Util.GetDataTable("CaseCode", "GroupName", "性別", "", ""), "CaseName", "CaseID", false);
        Util.FillDropDownList(ddlSex, Util.GetDataTable("CaseCode", "GroupName", "性別", "CaseID", ""), "CaseName", "CaseName", false);
        ddlSex.Items.Insert(0, new ListItem("", ""));
        ddlSex.SelectedIndex = 0;

        //稱謂
        // 2014/4/11 加入排序
        //Util.FillDropDownList(ddlTitle, Util.GetDataTable("CaseCode", "GroupName", "稱謂", "", ""), "CaseName", "CaseID", false);
        Util.FillDropDownList(ddlTitle, Util.GetDataTable("CaseCode", "GroupName", "稱謂", "CaseID", ""), "CaseName", "CaseName", false);
        ddlTitle.Items.Insert(0, new ListItem("", ""));
        ddlTitle.SelectedIndex = 3;

        //2014/5/14註解
        //身分別  2014/4/21新增
        //Util.FillCheckBoxList(cblDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "", ""), "CaseName", "CaseName", false);
        //cblDonor_Type.Items[0].Selected = false;

        //狀態
        Util.FillDropDownList(ddlMember_Status, Util.GetDataTable("CaseCode", "GroupName", "會員狀態", "", ""), "CaseName", "CaseName", false);
        ddlMember_Status.Items.Insert(0, new ListItem("", ""));
        ddlMember_Status.SelectedIndex = 2;

        /*20140701 移除欄位
        //教育程度
        Util.FillDropDownList(ddlEducation, Util.GetDataTable("CaseCode", "GroupName", "學歷", "", ""), "CaseName", "CaseName", false);
        ddlEducation.Items.Insert(0, new ListItem("", ""));
        ddlEducation.SelectedIndex = 0;
        //職業別
        Util.FillDropDownList(ddlOccupation, Util.GetDataTable("CaseCode", "GroupName", "職業", "", ""), "CaseName", "CaseName", false);
        ddlOccupation.Items.Insert(0, new ListItem("", ""));
        ddlOccupation.SelectedIndex = 0;
        //婚姻狀況
        Util.FillDropDownList(ddlMarriage, Util.GetDataTable("CaseCode", "GroupName", "婚姻", "", ""), "CaseName", "CaseName", false);
        ddlMarriage.Items.Insert(0, new ListItem("", ""));
        ddlMarriage.SelectedIndex = 0;
        //宗教信仰
        Util.FillDropDownList(ddlReligion, Util.GetDataTable("CaseCode", "GroupName", "宗教", "", ""), "CaseName", "CaseName", false);
        ddlReligion.Items.Insert(0, new ListItem("", ""));
        ddlReligion.SelectedIndex = 0;*/

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
        Util.FillDropDownList(ddlAlley, Util.GetDataTable("CaseCode", "GroupName", "弄衖別", "", ""), "CaseName", "CaseName", false);
        ddlAlley.Items.Insert(0, new ListItem("", ""));
        ddlAlley.SelectedIndex = 1;
    }
    //----------------------------------------------------------------------
    //20140509 新增 by Ian_Kao
    public string LoadCheckBoxListData()
    {
        string strSql = "select CaseName from CaseCode where CodeType = 'DonorType'";
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
        string strSql = "select CaseName from CaseCode where CodeType = 'DonorType'";
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
    //dropdownlist選城市自動帶入地區
    protected void ddlCity_SelectedIndexChanged(object sender, EventArgs e)
    {
        CityToArea(ddlArea, ddlCity);
        //20140813增加縣市與電話區碼連動
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
        tbxZipCode.Text = Util.GetCityCode(ddlArea.SelectedItem.Text, ddlCity.SelectedValue);
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
    }
    //----------------------------------------------------------------------
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        //string strMember_No = "";
        /*if (Util.GetControlValue("GN_DonorType") == "")
        {
            AjaxShowMessage("身分別欄位不得為空白");
            return;
        }*/
        //bool flag = false;
        try
        {
            //strMember_No = Donor_AddNew();
            Donor_AddNew();
            SetSysMsg("讀者資料已建立");
            //flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            //string strDonor_Id = GetMemberID(strMember_No);
            Response.Redirect(Util.RedirectByTime("Member_Edit.aspx", "Donor_Id=" + Donor_Id));
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("MemberQry.aspx"));
    }
    //----------------------------------------------------------------------
    public void Donor_AddNew()
    {
        string strSql = "insert into  Donor\n";
        strSql += "( Dept_Id, Introducer_Id, Introducer_Name, Donor_Name, Sex, title, Donor_Type, IDNo, Birthday,\n";
        strSql += " Cellular_Phone,Tel_Office_Loc, Tel_Office, Tel_Office_Ext, Fax_Loc, Fax, Email, Contactor, IsAbroad, ZipCode, City, Area, Address,\n";
        strSql += " Street, Section, Lane, Alley, HouseNo, HouseNoSub, Floor, FloorSub, Room, Attn, OverseasCountry, OverseasAddress, IsSendNews, IsSendNewsNum,\n";
        strSql += " IsSendEpaper, IsErrAddress, Remark, IsThanks, IsThanks_Add, Donate_No, Donate_Total, Donate_NoD, Donate_TotalD, Donate_NoC, Donate_TotalC,\n";
        strSql += " Donate_NoM, Donate_TotalM, Donate_NoS, Donate_TotalS, Donate_NoA, Donate_TotalA, Donate_NoND, Donate_TotalND, Create_Date, Create_DateTime,\n";
        strSql += " Create_User, Create_IP, IsMember, Member_Status, IsGroup, Group_No, Group_Licence, Group_CreateDate, Group_Person, Group_Person_JobTitle,\n";
        strSql += " Group_WebUrll, Group_Mission, Group_Service, Group_Other) values\n";
        strSql += "( @Dept_Id,@Introducer_Id,@Introducer_Name,@Donor_Name,@Sex,@title,@Donor_Type,@IDNo,@Birthday,\n";
        strSql += " @Cellular_Phone,@Tel_Office_Loc,@Tel_Office,@Tel_Office_Ext,@Fax_Loc,@Fax,@Email,@Contactor,@IsAbroad,@ZipCode,@City,@Area,@Address,\n";
        strSql += " @Street,@Section,@Lane,@Alley,@HouseNo,@HouseNoSub,@Floor,@FloorSub,@Room,@Attn,@OverseasCountry,@OverseasAddress,@IsSendNews,@IsSendNewsNum,\n";
        strSql += " @IsSendEpaper,@IsErrAddress,@Remark,@IsThanks,@IsThanks_Add,@Donate_No,@Donate_Total,@Donate_NoD,@Donate_TotalD,@Donate_NoC,@Donate_TotalC,\n";
        strSql += " @Donate_NoM,@Donate_TotalM,@Donate_NoS,@Donate_TotalS,@Donate_NoA,@Donate_TotalA,@Donate_NoND,@Donate_TotalND,@Create_Date,@Create_DateTime,\n";
        strSql += " @Create_User,@Create_IP,@IsMember,@Member_Status,@IsGroup,@Group_No,@Group_Licence,@Group_CreateDate,@Group_Person,@Group_Person_JobTitle,\n";
        strSql += " @Group_WebUrll,@Group_Mission,@Group_Service,@Group_Other) ";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Dept_Id", SessionInfo.DeptID);
        if(HFD_Donor_Id.Value!="")
        {
            dict.Add("Introducer_Id", HFD_Donor_Id.Value);
            dict.Add("Introducer_Name", tbxIntroducer_Name.Text.Trim());
        }
        else{
            dict.Add("Introducer_Id", "");
            dict.Add("Introducer_Name", "");
        }
        dict.Add("Donor_Name", tbxDonor_Name.Text.Trim());
        dict.Add("Sex", ddlSex.SelectedItem.Text);
        dict.Add("Title", ddlTitle.SelectedItem.Text);
        //dict.Add("Donor_Type", cblDonor_Type.SelectedValue);
        dict.Add("Donor_Type", Util.GetControlValue("GN_DonorType"));
        dict.Add("IDNo", tbxIDNo.Text.Trim());
        if (tbxBirthday.Text.Trim() != "")
        {
            dict.Add("Birthday", Util.DateTime2String(tbxBirthday.Text, DateType.yyyyMMdd, EmptyType.ReturnEmpty));
        }
        else
        {
            dict.Add("Birthday", null);
        }
        //20140701 調整欄位
        //dict.Add("Education", ddlEducation.SelectedItem.Text);
        //dict.Add("Occupation", ddlOccupation.SelectedItem.Text);
        //dict.Add("Marriage", ddlMarriage.SelectedItem.Text);
        //dict.Add("Religion", ddlReligion.SelectedItem.Text);
        //dict.Add("ReligionName", tbxReligionName.Text.Trim());
        dict.Add("Cellular_Phone", tbxCellular_Phone.Text.Trim());
        dict.Add("Tel_Office_Loc", tbxTel_Office_Loc.Text.Trim());
        dict.Add("Tel_Office", tbxTel_Office.Text.Trim());
        dict.Add("Tel_Office_Ext", tbxTel_Office_Ext.Text.Trim());
        dict.Add("Fax_Loc", tbxFax_Loc.Text.Trim());
        dict.Add("Fax", tbxFax.Text.Trim());

        dict.Add("Email", tbxEMail.Text.Trim());
        dict.Add("Contactor", tbxContactor.Text.Trim());
        //dict.Add("OrgName", tbxOrgName.Text.Trim());
        //dict.Add("JobTitle", tbxJobTitle.Text.Trim());
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
        //國內
        if (this.cbxIsLocal.Checked)
        {
            dict.Add("IsAbroad", "N");
            //地址
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
            if (tbxNo1.Text.Trim() != "" && tbxNo2.Text.Trim() == "")//有無號碼
            {
                Address += tbxNo1.Text.Trim() + "號";
                dict.Add("HouseNo", tbxNo1.Text.Trim());
                dict.Add("HouseNoSub", "");
            }
            else if (tbxNo1.Text.Trim() == "" && tbxNo2.Text.Trim() != "")
            {
                Address += tbxNo2.Text.Trim() + "號";
                dict.Add("HouseNo", "");
                dict.Add("HouseNoSub", tbxNo2.Text.Trim());
            }
            else if (tbxNo1.Text.Trim() != "" && tbxNo2.Text.Trim() != "")
            {
                Address += tbxNo1.Text.Trim() + "-" + tbxNo2.Text.Trim() + "號";
                dict.Add("HouseNo", tbxNo1.Text.Trim());
                dict.Add("HouseNoSub", tbxNo2.Text.Trim());
            }
            else
            {
                dict.Add("HouseNo", "");
                dict.Add("HouseNoSub", "");
            }
            if (tbxFloor1.Text.Trim() != "" && tbxFloor2.Text.Trim() == "")//有無樓別
            {
                Address += tbxFloor1.Text.Trim() + "樓";
                dict.Add("Floor", tbxFloor1.Text.Trim());
                dict.Add("FloorSub", "");
            }
            else if (tbxFloor1.Text.Trim() == "" && tbxFloor2.Text.Trim() != "")
            {
                Address += tbxFloor2.Text.Trim() + "樓";
                dict.Add("Floor", "");
                dict.Add("FloorSub", tbxFloor2.Text.Trim());
            }
            else if (tbxFloor1.Text.Trim() != "" && tbxFloor2.Text.Trim() != "")
            {
                Address += tbxFloor1.Text.Trim() + "-" + tbxFloor2.Text.Trim() + "樓";
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
                Address += tbxRoom.Text.Trim() + "室";
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
            string Address = "";
            if (tbxOverseasCountry.Text.Trim() != "")//有無國家/省城市/區
            {
                Address += tbxOverseasCountry.Text.Trim();
                dict.Add("OverseasCountry", tbxOverseasCountry.Text.Trim());
            }
            else
            {
                dict.Add("OverseasCountry", ""); 
            }
            if (tbxOverseasAddress.Text.Trim() != "")//有無其他
            {
                Address += tbxOverseasAddress.Text.Trim();
                dict.Add("OverseasAddress", tbxOverseasAddress.Text.Trim());
            }
            else
            {
                dict.Add("OverseasAddress", "");
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
        if (cbxIsSendEpaper.Checked)
        {
            dict.Add("IsSendEpaper", "Y");
        }
        else
        {
            dict.Add("IsSendEpaper", "N");
        }

        if (cbxIsErrAddress.Checked)
        {
            dict.Add("IsErrAddress", "Y");
        }
        else
        {
            dict.Add("IsErrAddress", "N");
        }
        dict.Add("Remark", tbxRemark.Text.Trim());

        dict.Add("IsThanks", "0");
        dict.Add("IsThanks_Add", "0");
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
        dict.Add("IsMember", "Y");
        //******設定會員編號******//2014/11/21取消會員編號 by 詩儀
//        string strSql2 = @"select isnull(MAX(Member_No),'') as Member_No from Donor
//                        where IsMember='Y'";
        //****執行查詢會員編號語法****//
        /*DataTable dt2 = NpoDB.QueryGetTable(strSql2);
        string Member_No = "";
        if (dt2.Rows.Count > 0)
        {
            string Member_No_value = dt2.Rows[0]["Member_No"].ToString();

            if (Member_No_value == "")
            {
                Member_No = "00001";
            }
            else
            {
                Member_No = (Convert.ToInt64(Member_No_value) + 1).ToString();
            }
        }*/
        //************************//
        //dict.Add("Member_No", Member_No);
        dict.Add("Member_Status", ddlMember_Status.SelectedItem.Text);
        dict.Add("IsGroup", "N");
        dict.Add("Group_No", "");
        dict.Add("Group_Licence", "");
        dict.Add("Group_CreateDate", null);
        dict.Add("Group_Person", "");
        dict.Add("Group_Person_JobTitle", "");
        dict.Add("Group_WebUrll", "");
        dict.Add("Group_Mission", "");
        dict.Add("Group_Service", "");
        dict.Add("Group_Other", "");
        // 2014/11/21 改成可回傳新增的捐款人編號，可在導向讀者管理【修改】頁面時帶入參數，自動帶出讀者資料
        //NpoDB.ExecuteSQLS(strSql, dict);
        //SetSysMsg("新增成功!");

        Donor_Id = NpoDB.GetScalarS(strSql, dict);
        if (Donor_Id != "")
        {
            flag = true;
        }
    }
    //20140410 修改by Ian_Kao 性別與稱謂連動  
    protected void ddlSex_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlSex.SelectedItem.Text == "女")
        {
            ddlTitle.Text = "小姐/女士";
        }
        else if (ddlSex.SelectedItem.Text == "男")
        {
            ddlTitle.Text = "先生";
        }
        else if (ddlSex.SelectedItem.Text == "歿")
        {
            ddlTitle.Text = "";
        }
        else
        {
            ddlTitle.Text = "先生/女士";
        }
    }
    private string GetMemberID(string strMember_No)
    {
        string strSql = " select Donor_Id from Donor ";
        strSql += " where Donor_Id = @Donor_Id and IsMember = @IsMember order by Create_DateTime desc";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", strMember_No);
        dict.Add("IsMember", 'Y');
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        DataRow dr = dt.Rows[0];
        return dr[0].ToString();
    }
}