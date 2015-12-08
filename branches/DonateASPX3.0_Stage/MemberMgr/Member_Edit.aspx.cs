using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class MemberMgr_Member_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Donor_Id");
            MemberMenu.Text = CaseUtil.MakeMemberMenu(HFD_Uid.Value, 1); //傳入目前選擇的 tab
            LoadDropDownListData();
            Form_DataBind();
        }
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " ReadOnly();", true);
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //性別
        Util.FillDropDownList(ddlSex, Util.GetDataTable("CaseCode", "GroupName", "性別", "", ""), "CaseName", "CaseName", false);
        ddlSex.Items.Insert(0, new ListItem("", ""));
        ddlSex.SelectedIndex = 0;

        //稱謂
        Util.FillDropDownList(ddlMember_Status, Util.GetDataTable("CaseCode", "GroupName", "會員狀態", "", ""), "CaseName", "CaseName", false);
        ddlMember_Status.Items.Insert(0, new ListItem("", ""));
        ddlMember_Status.SelectedIndex = 0;

        //20140514 註解
        //身分別  2014/4/21新增
        //Util.FillCheckBoxList(cblDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "", ""), "CaseName", "CaseName", false);
        //cblDonor_Type.Items[0].Selected = false;

        //狀態
        Util.FillDropDownList(ddlTitle, Util.GetDataTable("CaseCode", "GroupName", "稱謂", "", ""), "CaseName", "CaseName", false);
        ddlTitle.Items.Insert(0, new ListItem("", ""));
        ddlTitle.SelectedIndex = 0;

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

    }
    //----------------------------------------------------------------------
    public void CityToArea(DropDownList ddlTo, DropDownList ddlFrom)
    {
        Util.FillDropDownList(ddlTo, Util.GetDataTable("CodeCityNew", "ParentCityID", ddlFrom.SelectedValue, "", ""), "Name", "ZipCode", false);
        ddlTo.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlTo.SelectedIndex = 0;
    }
    //----------------------------------------------------------------------
    //帶入資料
    public void Form_DataBind()
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;

        //****變數設定****//
        uid = HFD_Uid.Value;

        //****設定查詢****//
        strSql = " select *  from DONOR where Donor_id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);

        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("MemberQry.aspx");

        DataRow dr = dt.Rows[0];
        //捐款人
        tbxDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
        //性別
        ddlSex.SelectedValue = dr["Sex"].ToString().Trim();
        //稱謂
        ddlTitle.SelectedValue = dr["Title"].ToString().Trim();
        //身分別
        //20140514  修改
        //cblDonor_Type.SelectedValue = dr["Donor_Type"].ToString().Trim();
        lblCheckBoxList.Text = LoadCheckBoxListData(dr["Donor_Type"].ToString());
        //狀態
        ddlMember_Status.SelectedValue = dr["Member_Status"].ToString().Trim();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //出生日期
        if (dr["Birthday"].ToString().Trim() != "")
        {
            tbxBirthday.Text = Util.DateTime2String(dr["Birthday"].ToString().Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty);
        }
        /*
        //教育程度
        ddlEducation.SelectedValue = dr["Education"].ToString().Trim();
        //職業別
        ddlOccupation.SelectedValue = dr["Occupation"].ToString().Trim();
        //婚姻狀況
        ddlMarriage.SelectedValue = dr["Marriage"].ToString().Trim();
        //宗教信仰
        ddlReligion.SelectedValue = dr["Religion"].ToString().Trim();
        //所屬教會
        tbxReligionName.Text = dr["ReligionName"].ToString().Trim();*/
        //手機
        tbxCellular_Phone.Text = dr["Cellular_Phone"].ToString().Trim();
        //電話
        tbxTel_Office_Loc.Text = dr["Tel_Office_Loc"].ToString().Trim();
        tbxTel_Office.Text = dr["Tel_Office"].ToString().Trim();
        tbxTel_Office_Ext.Text = dr["Tel_Office_Ext"].ToString().Trim();
        //傳真
        tbxFax_Loc.Text = dr["Fax_Loc"].ToString().Trim();
        tbxFax.Text = dr["Fax"].ToString().Trim();
        //Email
        tbxEMail.Text = dr["Email"].ToString();
        //聯絡人
        tbxContactor.Text = dr["Contactor"].ToString().Trim();
        /*//服務單位
        tbxOrgName.Text = dr["OrgName"].ToString().Trim();
        //職稱
        tbxJobTitle.Text = dr["JobTitle"].ToString().Trim();*/
        //台灣本島或海外地址
        if (dr["IsAbroad"].ToString().Trim() == "N")
        {
            cbxIsLocal.Checked = true;
            //區碼
            tbxZipCode.Text = dr["ZipCode"].ToString().Trim();
            //縣市
            if (dr["City"].ToString().Trim() != "")
            {
                ddlCity.Items.FindByText("縣 市").Selected = false;
                ddlCity.Items.FindByValue(dr["City"].ToString().Trim()).Selected = true;
            }
            // 鄉鎮市區
            CityToArea(ddlArea, ddlCity);
            //區域
            if (dr["Area"].ToString().Trim() != "")
            {
                ddlArea.Items.FindByText("鄉鎮市區").Selected = false;
                ddlArea.Items.FindByValue(dr["Area"].ToString().Trim()).Selected = true;
            }
            if (dr["Street"].ToString().Trim() != "")
            {
                tbxStreet.Text = dr["Street"].ToString().Trim();
            }
            if (dr["Section"].ToString().Trim() != "")
            {
                ddlSection.Items.FindByText("").Selected = false;
                ddlSection.Items.FindByText(dr["Section"].ToString().Trim()).Selected = true;
            }
            if (dr["Lane"].ToString().Trim() != "")
            {
                tbxLane.Text = dr["Lane"].ToString().Trim();
            }
            if (dr["Alley"].ToString().Trim() != "")
            {
                //20150518 通訊地址「弄」未正常顯示
                tbxAlley.Text = dr["Alley"].ToString().Trim();
                if (dr["Address"].ToString().IndexOf('弄') > -1)
                    ddlAlley.Text = "弄";
                else if (dr["Address"].ToString().IndexOf('衖') > -1)
                    ddlAlley.Text = "衖";
                else
                    ddlAlley.SelectedIndex = 0;
            }
            if (dr["HouseNo"].ToString().Trim() != "")
            {
                tbxNo1.Text = dr["HouseNo"].ToString().Trim();
            }
            if (dr["HouseNoSub"].ToString().Trim() != "")
            {
                tbxNo2.Text = dr["HouseNoSub"].ToString().Trim();
            }
            if (dr["Floor"].ToString().Trim() != "")
            {
                tbxFloor1.Text = dr["Floor"].ToString().Trim();
            }
            if (dr["FloorSub"].ToString().Trim() != "")
            {
                tbxFloor2.Text = dr["FloorSub"].ToString().Trim();
            }
            if (dr["Room"].ToString().Trim() != "")
            {
                tbxRoom.Text = dr["Room"].ToString().Trim();
            }
            if (dr["Attn"].ToString().Trim() != "")
            {
                tbxAttn.Text = dr["Attn"].ToString().Trim();
            }
        }
        if (dr["IsAbroad"].ToString().Trim() == "Y")
        {
            cbxIsAbroad.Checked = true;
            if (dr["OverseasCountry"].ToString().Trim() != "")
            {
                tbxOverseasCountry.Text = dr["OverseasCountry"].ToString().Trim();
            }
            if (dr["OverseasAddress"].ToString().Trim() != "")
            {
                tbxOverseasAddress.Text = dr["OverseasAddress"].ToString().Trim();
            }
        }
        //介紹人
        HFD_Donor_Id.Value = dr["Introducer_Id"].ToString().Trim();
        tbxIntroducer_Name.Text = dr["Introducer_Name"].ToString().Trim();
        //文宣品
        tbxIsSendNewsNum.Text = dr["IsSendNewsNum"].ToString().Trim();
        if (dr["IsSendEpaper"].ToString().Trim() == "Y")
        {
            cbxIsSendEpaper.Checked = true;
        }
        if (dr["IsErrAddress"].ToString().Trim() == "Y")
        {
            cbxIsErrAddress.Checked = true;
        }
        //捐款人備註
        tbxRemark.Text = dr["Remark"].ToString().Trim();
        //資料建檔日期 
        if (dr["Create_Date"].ToString() != "" && DateTime.Parse(dr["Create_Date"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxCreate_Date.Text = DateTime.Parse(dr["Create_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //資料建檔人員  
        if (dr["Create_User"].ToString() != "")
        {
            tbxCreate_User.Text = dr["Create_User"].ToString();
        }
        //最後異動日期  
        if (dr["LastUpdate_Date"].ToString() != "" && DateTime.Parse(dr["LastUpdate_Date"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxLastUpdate_Date.Text = DateTime.Parse(dr["LastUpdate_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //最後異動人員 
        if (dr["LastUpdate_User"].ToString() != "")
        {
            tbxLastUpdate_User.Text = dr["LastUpdate_User"].ToString();
        }
    }
    //------------------------------------------------------------------------
    //20140509 新增 by Ian_Kao
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
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        /*if (Util.GetControlValue("GN_DonorType") == "")
        {
            AjaxShowMessage("身分別欄位不得為空白");
            return;
        }*/
        bool flag = false;
        try
        {
            Donor_Edit();
            SetSysMsg("讀者資料修改資料成功");
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            Response.Redirect(Util.RedirectByTime("Member_Edit.aspx", "Donor_Id=" + HFD_Uid.Value));
        }
    }
    protected void btnDel_Click(object sender, EventArgs e)
    {
        string strSql = "";
        bool flag = false;
        try
        {
            //先確定無繳費紀錄存在
            strSql = " select * from Donate where Donor_Id ='" + HFD_Uid.Value +"'";
            DataTable dt = NpoDB.QueryGetTable(strSql);
            if (dt.Rows.Count > 0)
            {
                ShowSysMsg("此會員尚有繳費記錄存在，請先刪除繳費明細，才能刪除會員！");
                return;
            }
            else
            {
                //****變數宣告****//
                Dictionary<string, object> dict = new Dictionary<string, object>();

                //****設定SQL指令****//
                strSql = " update Donor set ";
                strSql += " DeleteDate = GetDate()";
                strSql += " where Donor_Id = @Donor_Id";

                dict.Add("Donor_Id", this.HFD_Uid.Value);

                //****執行語法****//
                NpoDB.ExecuteSQLS(strSql, dict);

                //20140416 紀錄Log檔 
                dict.Clear();
                dict.Add("Donor_Id", HFD_Uid.Value);
                string Donor_Name = NpoDB.GetScalarS("select Donor_Name from Donor where Donor_Id = @Donor_Id", dict);
                CaseUtil.InsertLogData(SessionInfo.UserID, "刪除資料", "刪除" + HFD_Uid.Value + Donor_Name);

                SetSysMsg("讀者資料刪除成功!");
                flag = true;
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            Response.Redirect(Util.RedirectByTime("MemberQry.aspx"));
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("MemberQry.aspx"));
    }
    public void Donor_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update Donor set ";
        strSql += " Introducer_Id = @Introducer_Id";
        strSql += ", Introducer_Name = @Introducer_Name";
        strSql += ", Donor_Name = @Donor_Name";
        strSql += ", Sex = @Sex";
        strSql += ", title = @title";
        strSql += ", Donor_Type = @Donor_Type";
        strSql += ", IDNo = @IDNo";
        strSql += ", Birthday = @Birthday";
        /*strSql += ", Education = @Education";
        strSql += ", Occupation = @Occupation";
        strSql += ", Marriage = @Marriage";
        strSql += ", Religion = @Religion";
        strSql += ", ReligionName = @ReligionName";*/
        strSql += ", Cellular_Phone = @Cellular_Phone";
        strSql += ", Tel_Office_Loc = @Tel_Office_Loc";
        strSql += ", Tel_Office = @Tel_Office";
        strSql += ", Tel_Office_Ext = @Tel_Office_Ext";
        strSql += ", Fax_Loc = @Fax_Loc";
        strSql += ", Fax = @Fax";
        strSql += ", Email = @Email";
        strSql += ", Contactor = @Contactor";
        //strSql += ", OrgName = @OrgName";
        //strSql += ", JobTitle = @JobTitle";
        strSql += ", IsAbroad=@IsAbroad";
        strSql += ", ZipCode= @ZipCode";
        strSql += ", City= @City";
        strSql += ", Area= @Area";
        strSql += ", Address= @Address";
        strSql += ", Street= @Street";
        strSql += ", Section= @Section";
        strSql += ", Lane= @Lane";
        strSql += ", Alley= @Alley";
        strSql += ", HouseNo= @HouseNo";
        strSql += ", HouseNoSub= @HouseNoSub";
        strSql += ", Floor= @Floor";
        strSql += ", FloorSub= @FloorSub";
        strSql += ", Room= @Room";
        strSql += ", Attn= @Attn";
        strSql += ", OverseasCountry= @OverseasCountry";
        strSql += ", OverseasAddress= @OverseasAddress";
        strSql += ", IsSendNews= @IsSendNews";
        strSql += ", IsSendNewsNum= @IsSendNewsNum";
        strSql += ", IsSendEpaper= @IsSendEpaper";
        strSql += ", IsErrAddress= @IsErrAddress";
        strSql += ", Remark = @Remark ";
        strSql += ", Member_Status = @Member_Status";
        strSql += ", LastUpdate_Dept_Id = @LastUpdate_Dept_Id";
        strSql += ", LastUpdate_Date = @LastUpdate_Date";
        strSql += ", LastUpdate_DateTime = @LastUpdate_DateTime";
        strSql += ", LastUpdate_User = @LastUpdate_User";
        strSql += ", LastUpdate_IP = @LastUpdate_IP";
        strSql += " where Donor_Id = @Donor_Id";

        if (HFD_Donor_Id.Value != "")
        {
            dict.Add("Introducer_Id", HFD_Donor_Id.Value);
            dict.Add("Introducer_Name", tbxIntroducer_Name.Text.Trim());
        }
        else
        {
            dict.Add("Introducer_Id", "");
            dict.Add("Introducer_Name", "");
        }
        dict.Add("Donor_Name", tbxDonor_Name.Text.Trim());
        dict.Add("Sex", ddlSex.SelectedItem.Text);
        dict.Add("Title", ddlTitle.SelectedItem.Text);
        //dict.Add("Donor_Type", cblDonor_Type.SelectedValue);
        //20140514 修改
        dict.Add("Donor_Type", Util.GetControlValue("GN_DonorType"));
        dict.Add("IDNo", tbxIDNo.Text.Trim());
        if (tbxBirthday.Text.Trim() != "")
        {
            dict.Add("Birthday", Util.DateTime2String(tbxBirthday.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty));
        }
        else
        {
            dict.Add("Birthday", null);
        }
        /*dict.Add("Education", ddlEducation.SelectedItem.Text);
        dict.Add("Occupation", ddlOccupation.SelectedItem.Text);
        dict.Add("Marriage", ddlMarriage.SelectedItem.Text);
        dict.Add("Religion", ddlReligion.SelectedItem.Text);
        dict.Add("ReligionName", tbxReligionName.Text.Trim());*/
        dict.Add("Cellular_Phone", tbxCellular_Phone.Text.Trim());
        dict.Add("Tel_Office_Loc", tbxTel_Office_Loc.Text.Trim());
        dict.Add("Tel_Office", tbxTel_Office.Text.Trim());
        dict.Add("Tel_Office_Ext", tbxTel_Office_Ext.Text.Trim());
        dict.Add("Fax_Loc", tbxFax_Loc.Text.Trim());
        dict.Add("Fax", tbxFax.Text.Trim());
        dict.Add("Email", tbxEMail.Text.Trim());
        dict.Add("Contactor", tbxContactor.Text.Trim());
        /*dict.Add("OrgName", tbxOrgName.Text.Trim());
        dict.Add("JobTitle", tbxJobTitle.Text.Trim());*/
        
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
            dict.Add("ZipCode", "");
            dict.Add("City", "");
            dict.Add("Area", "");
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
            dict.Add("IsSendNews", "N");
            dict.Add("IsSendNewsNum", "0");
        }
        else
        {
            dict.Add("IsSendNews", "Y");
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

        dict.Add("Member_Status", ddlMember_Status.SelectedItem.Text);
        dict.Add("LastUpdate_Dept_Id", SessionInfo.DeptID);
        dict.Add("LastUpdate_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("LastUpdate_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        dict.Add("LastUpdate_User", SessionInfo.UserName);
        dict.Add("LastUpdate_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        dict.Add("Donor_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    //20140410 修改by Ian_Kao 性別與稱謂連動  
    protected void ddlSex_SelectedIndexChanged(object sender, EventArgs e)
    {
        // 性別的資料欄位CaseID 對應 稱謂的資料欄位CaseID
        string caseSwitch = ddlSex.SelectedItem.Text;
        switch (caseSwitch)
        {
            case "男":
                ddlTitle.SelectedItem.Text = "先生";
                break;
            case "女":
                ddlTitle.SelectedItem.Text = "小姐/女士";
                break;
            case "歿":
                ddlTitle.SelectedItem.Text = "";
                break;
            default:
                ddlTitle.SelectedItem.Text = "先生/女士";
                break;
        }
    }
}