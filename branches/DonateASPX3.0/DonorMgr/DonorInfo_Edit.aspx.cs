using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;
using System.Web.UI.HtmlControls;

public partial class DonorMgr_DonorInfo_Edit : BasePage
{

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            //權控處理
            AuthrityControl();

            Session["cType"] = "";
            HFD_Uid.Value = Util.GetQueryString("Donor_Id");
            MemberMenu.Text = CaseUtil.MakeMenu(HFD_Uid.Value, 1); //傳入目前選擇的 tab
            //產出下拉式選單
            LoadDropDownListData();
            //載入捐款人資料
            Form_DataBind();
        }
    }

    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_Update", btnEdit);
        Authrity.CheckButtonRight("_Delete", btnDel);
        //Authrity.CheckButtonRight("_Update", btnGroupEdit);
        //MemberMenu.Visible = Authrity.RightCheck("_Update");
        Authrity.CheckButtonRight("_AddNew", btnAddDonateData);
    }

    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //性別
        // 2014/4/7 加入排序
        //Util.FillDropDownList(ddlSex, Util.GetDataTable("CaseCode", "GroupName", "性別", "", ""), "CaseName", "CaseName", false);
        Util.FillDropDownList(ddlSex, Util.GetDataTable("CaseCode", "GroupName", "性別", "CaseID", ""), "CaseName", "CaseName", false);
        ddlSex.Items.Insert(0, new ListItem("", ""));
        ddlSex.SelectedIndex = 0;

        //稱謂
        // 2014/4/7 加入排序
        //Util.FillDropDownList(ddlTitle, Util.GetDataTable("CaseCode", "GroupName", "稱謂", "", ""), "CaseName", "CaseName", false);
        Util.FillDropDownList(ddlTitle, Util.GetDataTable("CaseCode", "GroupName", "稱謂", "CaseID", ""), "CaseName", "CaseName", false);
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
        Util.FillDropDownList(ddlSection, Util.GetDataTable("CaseCode", "GroupName", "段別", "", ""), "CaseName", "CaseID", false);
        ddlSection.Items.Insert(0, new ListItem("", ""));
        ddlSection.SelectedIndex = 0;

        //弄衖
        Util.FillDropDownList(ddlAlley, Util.GetDataTable("CaseCode", "GroupName", "弄衖別", "", ""), "CaseName", "CaseID", false);
        ddlAlley.Items.Insert(0, new ListItem("", ""));
        ddlAlley.SelectedIndex = 1;

        //收據開立
        Util.FillDropDownList(ddlInvoice_Type, Util.GetDataTable("CaseCode", "GroupName", "收據開立", "", ""), "CaseName", "CaseID", false);
        ddlInvoice_Type.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Type.SelectedIndex = 3;

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
        Util.FillDropDownList(ddlInvoice_Section, Util.GetDataTable("CaseCode", "GroupName", "段別", "", ""), "CaseName", "CaseID", false);
        ddlInvoice_Section.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Section.SelectedIndex = 0;

        //弄衖(收據地址)
        Util.FillDropDownList(ddlInvoice_Alley, Util.GetDataTable("CaseCode", "GroupName", "弄衖別", "", ""), "CaseName", "CaseID", false);
        ddlInvoice_Alley.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Alley.SelectedIndex = 1;

        //徵信錄原則
        //Util.FillDropDownList(ddlIsAnonymous, Util.GetDataTable("CaseCode", "GroupName", "徵信錄", "", ""), "CaseName", "CaseID", false);
        //ddlIsAnonymous.SelectedIndex = 0;

        //國家 20150428新增
        Util.FillDropDownList(ddlOverseasCountry, Util.GetDataTable("CaseCode", "GroupName", "國家", "CaseID", ""), "CaseName", "CaseName", false);
        ddlOverseasCountry.Items.Insert(0, new ListItem("", ""));
        ddlOverseasCountry.SelectedIndex = 0;

        Util.FillDropDownList(ddlInvoice_OverseasCountry, Util.GetDataTable("CaseCode", "GroupName", "國家", "CaseID", ""), "CaseName", "CaseName", false);
        ddlInvoice_OverseasCountry.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_OverseasCountry.SelectedIndex = 0;
    }

    //----------------------------------------------------------------------
    public void CityToArea(DropDownList ddlTo, DropDownList ddlFrom)
    {
        Util.FillDropDownList(ddlTo, Util.GetDataTable("CodeCity", "ParentCityID", ddlFrom.SelectedValue, "", ""), "Name", "ZipCode", false);
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
        strSql = " select *  from DONOR where Donor_id='" + uid +"'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("DonorInfo.aspx");

        DataRow dr = dt.Rows[0];
        //捐款人編號
        lblDonor_Id.Text = dr["Donor_Id"].ToString().Trim();
        //捐款人
        txtDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
        //性別
        ddlSex.SelectedValue = dr["Sex"].ToString().Trim();
        //稱謂
        ddlTitle.SelectedValue = dr["Title"].ToString().Trim();
        //身分別
        //20140509  修改 by Ian_Kao
        //cblDonor_Type.SelectedValue = dr["Donor_Type"].ToString().Trim();
        lblCheckBoxList.Text = LoadCheckBoxListData(dr["Donor_Type"].ToString());
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //出生日期
        if (dr["Birthday"].ToString().Trim() != "")
        {
            txtBirthday.Text = DateTime.Parse(dr["Birthday"].ToString().Trim()).ToShortDateString().ToString();
        }
        //手機
        tbxCellular_Phone.Text = dr["Cellular_Phone"].ToString().Trim();
        //電話(區碼)
        tbxTel_Office_Loc.Text = dr["Tel_Office_Loc"].ToString().Trim();
        //電話
        tbxTel_Office.Text = dr["Tel_Office"].ToString().Trim();
        //電話(分機)
        tbxTel_Office_Ext.Text = dr["Tel_Office_Ext"].ToString().Trim();
        //傳真(區碼)
        tbxFax_Loc.Text = dr["Fax_Loc"].ToString().Trim();
        //傳真
        tbxFax.Text = dr["Fax"].ToString().Trim();
        //Email
        tbxEMail.Text = dr["Email"].ToString();
        //聯絡人
        tbxContactor.Text = dr["Contactor"].ToString().Trim();
        //職稱
        //tbxJobTitle.Text = dr["JobTitle"].ToString().Trim();
        //台灣本島或海外地址
        if (dr["IsAbroad"].ToString().Trim() == "N")
        {
            cbxIsLocal.Checked = true;
        }
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
            //Modify by GoodTV Tanya:通訊地址「弄」未正常顯示
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
        if (dr["IsAbroad"].ToString().Trim() == "Y")
        {
            cbxIsAbroad.Checked = true;
        }
        if (dr["OverseasCountry"].ToString().Trim() != "")
        {
            //tbxOverseasCountry.Text = dr["OverseasCountry"].ToString().Trim();
            //20150428新增
            ddlOverseasCountry.Items.FindByText("").Selected = false;
            ddlOverseasCountry.Items.FindByText(dr["OverseasCountry"].ToString().Trim()).Selected = true;
        }
        if (dr["OverseasAddress"].ToString().Trim() != "")
        {
            tbxOverseasAddress.Text = dr["OverseasAddress"].ToString().Trim();
        }
        //20150617新增OTT地址
        if (dr["OTT_Address"].ToString().Trim() != "")
        {
            tbxOTT_Address.Text = dr["OTT_Address"].ToString().Trim();
        }
        //文宣品
        tbxIsSendNewsNum.Text = dr["IsSendNewsNum"].ToString().Trim();
        if (dr["IsDVD"].ToString().Trim() == "Y")
        {
            cbxIsDVD.Checked = true;
        }
        if (dr["IsSendEpaper"].ToString().Trim() == "Y")
        {
            cbxIsSendEpaper.Checked = true;
        }
        if (dr["IsGift"].ToString().Trim() == "Y")
        {
            cbxIsGift.Checked = true;
        }
        if (dr["IsBigAmtThank"].ToString().Trim() == "Y")
        {
            cbxIsBigAmtThank.Checked = true;
        }
        if (dr["IsPost"].ToString().Trim() == "Y")
        {
            cbxIsPost.Checked = true;
        }
        if (dr["IsContact"].ToString().Trim() == "N")
        {
            cbxIsContact.Checked = true;
        }
        if (dr["IsErrAddress"].ToString().Trim() == "Y")
        {
            cbxIsErrAddress.Checked = true;
        }

        //收據開立
        if (dr["Invoice_Type"].ToString().Trim() != "")
        {
            if (dr["Invoice_Type"].ToString().Trim() == "不寄" || dr["Invoice_Type"].ToString().Trim() == "單次收據" || dr["Invoice_Type"].ToString().Trim() == "年度證明")
            {
                ddlInvoice_Type.Items.FindByText("年度證明").Selected = false;
                ddlInvoice_Type.Items.FindByText(dr["Invoice_Type"].ToString().Trim()).Selected = true;
            }
            else
            {
                ddlInvoice_Type.SelectedIndex = 0;
            }
        }
        //收據稱謂
        if (dr["Title2"].ToString().Trim() != "")
        {
            ddlTitle2.Items.FindByText("").Selected = false;
            ddlTitle2.Items.FindByText(dr["Title2"].ToString().Trim()).Selected = true;
        }
        //收據台灣本島或海外地址
        if (dr["IsAbroad_Invoice"].ToString().Trim() == "N")
        {
            cbxInvoice_IsLocal.Checked = true;
        }
        //區碼
        tbxInvoice_ZipCode.Text = dr["Invoice_ZipCode"].ToString().Trim();
        //縣市
        if (dr["Invoice_City"].ToString().Trim().Trim() != "")
        {
            ddlInvoice_City.Items.FindByText("縣 市").Selected = false;
            ddlInvoice_City.Items.FindByValue(dr["Invoice_City"].ToString().Trim()).Selected = true;
        }
        // 鄉鎮市區
        CityToArea(ddlInvoice_Area, ddlInvoice_City);
        //區域  
        if (dr["Invoice_Area"].ToString().Trim() != "")
        {
            ddlInvoice_Area.Items.FindByText("鄉鎮市區").Selected = false;
            ddlInvoice_Area.Items.FindByValue(dr["Invoice_Area"].ToString().Trim()).Selected = true;
        }
        if (dr["Invoice_Street"].ToString().Trim() != "")
        {
            tbxInvoice_Street.Text = dr["Invoice_Street"].ToString().Trim();
        }
        if (dr["Invoice_Section"].ToString().Trim() != "")
        {
            ddlInvoice_Section.Items.FindByText("").Selected = false;
            ddlInvoice_Section.Items.FindByText(dr["Invoice_Section"].ToString().Trim()).Selected = true;
        }
        if (dr["Invoice_Lane"].ToString().Trim() != "")
        {
            tbxInvoice_Lane.Text = dr["Invoice_Lane"].ToString().Trim();
        }
        if (dr["Invoice_Alley"].ToString().Trim() != "")
        {
            //20140808 Modify by 詩儀:收據地址「弄」未正常顯示
            tbxInvoice_Alley0.Text = dr["Invoice_Alley"].ToString().Trim();
            if (dr["Invoice_Address"].ToString().IndexOf('弄') > -1)
                ddlInvoice_Alley.Text = "弄";
            else if (dr["Invoice_Address"].ToString().IndexOf('衖') > -1)
                ddlInvoice_Alley.Text = "衖";
            else
                ddlInvoice_Alley.SelectedIndex = 0;
        }
        if (dr["Invoice_HouseNo"].ToString().Trim() != "")
        {
            tbxInvoice_No1.Text = dr["Invoice_HouseNo"].ToString().Trim();
        }
        if (dr["Invoice_HouseNoSub"].ToString().Trim() != "")
        {
            tbxInvoice_No2.Text = dr["Invoice_HouseNoSub"].ToString().Trim();
        }
        if (dr["Invoice_Floor"].ToString().Trim() != "")
        {
            tbxInvoice_Floor1.Text = dr["Invoice_Floor"].ToString().Trim();
        }
        if (dr["Invoice_FloorSub"].ToString().Trim() != "")
        {
            tbxInvoice_Floor2.Text = dr["Invoice_FloorSub"].ToString().Trim();
        }
        if (dr["Invoice_Room"].ToString().Trim() != "")
        {
            tbxInvoice_Room.Text = dr["Invoice_Room"].ToString().Trim();
        }
        if (dr["Invoice_Attn"].ToString().Trim() != "")
        {
            tbxInvoice_Attn.Text = dr["Invoice_Attn"].ToString().Trim();
        }
        if (dr["IsAbroad_Invoice"].ToString().Trim() == "Y")
        {
            cbxInvoice_IsAbroad.Checked = true;
        }
        if (dr["Invoice_OverseasCountry"].ToString().Trim() != "")
        {
            //tbxInvoice_OverseasCountry.Text = dr["Invoice_OverseasCountry"].ToString().Trim();
            //20150428新增
            ddlInvoice_OverseasCountry.Items.FindByText("").Selected = false;
            ddlInvoice_OverseasCountry.Items.FindByText(dr["Invoice_OverseasCountry"].ToString().Trim()).Selected = true;
        }
        if (dr["Invoice_OverseasAddress"].ToString().Trim() != "")
        {
            tbxInvoice_OverseasAddress.Text = dr["Invoice_OverseasAddress"].ToString().Trim();
        }
        //收據抬頭
        tbxInvoice_Title.Text = dr["Invoice_Title"].ToString().Trim();
        hfInvoiceTitle.Value = tbxInvoice_Title.Text;
        //收據身分證/統編
        tbxInvoice_IDNo.Text = dr["Invoice_IDNo"].ToString().Trim();
        //徵信錄原則
        if (dr["IsAnonymous"].ToString().Trim() == "Y")
        {
            cbxIsAnonymous.Checked = true;
        }
        //if (dr["IsAnonymous"].ToString().Trim() != "")
        //{
        //    //if (dr["Report_Type"].ToString().Trim() == "匿名" || dr["Report_Type"].ToString().Trim() == "不寄" || dr["Report_Type"].ToString().Trim() == "刊登其他名稱")
        //    //{
        //        ddlIsAnonymous.Items.FindByText("刊登").Selected = false;
        //        if (dr["IsAnonymous"].ToString().Trim() == "Y")
        //        {
        //            ddlIsAnonymous.Items.FindByText("不刊登").Selected = true;
        //        }
        //        else
        //        {
        //            ddlIsAnonymous.Items.FindByText("刊登").Selected = true;
        //        }
                
        //    //}
        //    //else
        //    //{
        //    //    ddlReport_Type.Items.Add(dr["Report_Type"].ToString());
        //    //}
        //}

        //上傳至國稅局
        if (dr["IsFdc"].ToString().Trim() == "Y")
        {
            cbxIsFdc.Checked=true;
        }
        //捐款人備註
        tbxRemark.Text = dr["Remark"].ToString().Trim();
        //想對GOOD TV說的話
        tbxToGOODTV.Text = dr["ToGOODTV"].ToString().Trim();
        //首次捐款日期  
        if (dr["Begin_DonateDate"].ToString() != "" && DateTime.Parse(dr["Begin_DonateDate"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxBegin_DonateDate.Text = DateTime.Parse(dr["Begin_DonateDate"].ToString()).ToString("yyyy/MM/dd");
        }
        //最近捐款日期  
        if (dr["Last_DonateDate"].ToString() != "" && DateTime.Parse(dr["Last_DonateDate"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxLast_DonateDate.Text = DateTime.Parse(dr["Last_DonateDate"].ToString()).ToString("yyyy/MM/dd");
        }
        //累計捐款次數  
        if (dr["Donate_No"].ToString() != "")
        {
            tbxDonate_No.Text = dr["Donate_No"].ToString();
        }
        else
        {
            tbxDonate_No.Text = "0";
        }
        //累計捐款金額  
        if (dr["Donate_Total"].ToString() != "")
        {
            tbxDonate_Total.Text = (Convert.ToInt64(dr["Donate_Total"])).ToString();
        }
        else
        {
            tbxDonate_Total.Text = "0";
        }
        //首次捐物日期  
        if (dr["Begin_DonateDateC"].ToString() != "" && DateTime.Parse(dr["Begin_DonateDateC"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxBegin_DonateDateC.Text = DateTime.Parse(dr["Begin_DonateDateC"].ToString()).ToString("yyyy/MM/dd");  
        }
        //最近捐物日期
        if (dr["Last_DonateDateC"].ToString() != "" && DateTime.Parse(dr["Last_DonateDateC"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxLast_DonateDateC.Text = DateTime.Parse(dr["Last_DonateDateC"].ToString()).ToString("yyyy/MM/dd");  
        }
        //累計捐物次數  
        if (dr["Donate_NoC"].ToString() != "")
        {
            tbxDonate_NoC.Text = dr["Donate_NoC"].ToString();
        }
        else
        {
            tbxDonate_NoC.Text = "0";
        }
        //累計折合現金
        if (dr["Donate_TotalC"].ToString() != "")
        {
            tbxDonate_TotalC.Text = (Convert.ToInt64(dr["Donate_TotalC"])).ToString();
        }
        else
        {
            tbxDonate_TotalC.Text = "0";
        }
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

        // 2014/7/14 增加所屬捐款人群組的滑鼠Tip呈現與表單內呈現
        //20140521 取得捐款人知所屬群組
        //lblDonorGroup.Text = CaseUtil.GetDonorGroup(uid);
        From_DonorGroup(uid);
        
    }

    // 2014/7/14 增加所屬捐款人群組的滑鼠Tip呈現與表單內呈現
    private void From_DonorGroup(string strUid)
    {
        string strSql = @"
                    select G.GroupClassName,I.GroupItemName,I.[Supplement],M.GroupItemUid
                    from GroupMapping M
                    join GroupItem I
                    on M.GroupItemUid=I.Uid
                    and I.DeleteDate is null
                    join GroupClass G
                    on I.GroupClassUid = g.[Uid]
                    and G.DeleteDate is null
                    where M.DonorId = @DonorID
                    and M.DeleteDate is null
                    order by G.GroupClassName,I.GroupItemName
                        ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("DonorID", strUid);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {

            DataRow dr;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dr = dt.Rows[i];
                string strItemUid = dr["GroupItemUid"].ToString().Trim();
                string strMember = "";
                string strSql2 = @"
					select Donor_Id,Donor_Name,Sex,Title,Donor_Type,Cellular_Phone,Tel_Office,Tel_Office_Loc,Tel_Office_Ext,Email
					,(Case When D.IsAbroad = 'N' then B.Name + C.Name + D.[Address] Else D.[Address] End) as [Address] 
					,Invoice_Type,Remark
                    from GroupMapping GM1 
                    join Donor D on GM1.DonorId = D.Donor_Id and D.DeleteDate is null
			        left Join CODECITY As B on D.City = B.ZipCode 
				    left Join CODECITY As C on D.Area = C.ZipCode
	    			where GM1.DeleteDate is null and GM1.GroupItemUid = @ItemUid
                    order by Donor_Id
                        ";

                Dictionary<string, object> dict2 = new Dictionary<string, object>();
                dict2.Add("ItemUid", strItemUid);
                DataTable dt2 = NpoDB.GetDataTableS(strSql2, dict2);
                if (dt2.Rows.Count > 0)
                {

                    DataRow dr2;
                    for (int j = 0; j < dt2.Rows.Count; j++)
                    {
                        dr2 = dt2.Rows[j];
                        string strMemberTitle = "";
                        string strTel = dr2["Tel_Office_Loc"].ToString().Trim() == "" ? "" : "(" + dr2["Tel_Office_Loc"].ToString().Trim() + ")";
                        strTel += dr2["Tel_Office"].ToString().Trim();
                        strTel += strTel + dr2["Tel_Office_Ext"].ToString().Trim() == "" ? "" : "#" + dr2["Tel_Office_Ext"].ToString().Trim();
                        strMemberTitle += "性別：" + dr2["Sex"].ToString().Trim() + "\n";
                        strMemberTitle += "稱謂：" + dr2["Title"].ToString().Trim() + "\n";
                        strMemberTitle += "身分別：" + dr2["Donor_Type"].ToString().Trim() + "\n";
                        strMemberTitle += "手機：" + dr2["Cellular_Phone"].ToString().Trim() + "\n";
                        strMemberTitle += "電話(日)：" + strTel + "\n";
                        strMemberTitle += "E-Mail：" + dr2["Email"].ToString().Trim() + "\n";
                        strMemberTitle += "通訊地址：" + dr2["Address"].ToString().Trim() + "\n";
                        strMemberTitle += "收據開立：" + dr2["Invoice_Type"].ToString().Trim() + "\n";
                        strMemberTitle += "捐款人備註：\n" + dr2["Remark"].ToString().Trim() + "\n";
                        if (strUid != dr2["Donor_Id"].ToString().Trim())
                        {
                            strMember += (strMember == "" ? "<a href='" + Util.RedirectByTime("DonorInfo_Edit.aspx", "Donor_Id=") +
                                dr2["Donor_Id"].ToString().Trim() + "' title='" + strMemberTitle + "'>" + dr2["Donor_Id"].ToString().Trim() + "(" + dr2["Donor_Name"].ToString().Trim() + ")"
                                + "</a>" : ", " + "<a href='" + Util.RedirectByTime("DonorInfo_Edit.aspx", "Donor_Id=") +
                                dr2["Donor_Id"].ToString().Trim() + "' title='" + strMemberTitle + "'>" + dr2["Donor_Id"].ToString().Trim() + "(" + dr2["Donor_Name"].ToString().Trim() + ")" + "</a>");
                        }
                        else
                        {
                            strMember += (strMember == "" ? "<font color='blue'>" + dr2["Donor_Id"].ToString().Trim() + "(" + dr2["Donor_Name"].ToString().Trim() + ")"
                                + "</font>" : ", " + "<font color='blue'>" + dr2["Donor_Id"].ToString().Trim() + "(" + dr2["Donor_Name"].ToString().Trim() + ")" + "</font>");
                        }


                    }

                }


                DonorGroup.Text += @"<tr style='background-color: #FFFFCC'><td>" + dr["GroupClassName"].ToString().Trim() + "</td><td><A href='" + Util.RedirectByTime("../DonorGroupMgr/GroupItemMember.aspx", "GroupItemUid=") + strItemUid + "'>" + dr["GroupItemName"].ToString().Trim() + "</A></td><td>" + dr["Supplement"].ToString().Trim() + "&nbsp;</td><td>" + strMember + "&nbsp;</td><td><img style='CURSOR: pointer' src='../images/delete2.gif' border=0 alt='刪除群組代表' onclick='deleteGroupItem(this," + strItemUid + ");' /></td></tr>";
            }

        }

    }

    //------------------------------------------------------------------------
    //20140509 新增 by Ian_Kao
    public string LoadCheckBoxListData(string DonorTypeValues)
    {
        string strSql = "select CaseName from CaseCode where CodeType = 'DonorType' order by CaseID";
        DataTable dt = NpoDB.GetDataTableS(strSql, null);

        List<ControlData> list = new List<ControlData>();
        list.Clear();
        foreach (DataRow dr in dt.Rows)
        {
            string CaseName = dr["CaseName"].ToString().Trim();
            bool ShowBR = false;
            list.Add(new ControlData("Checkbox", "GN_DonorType", "CB_DonorType", CaseName, CaseName, ShowBR, DonorTypeValues));
        }
        return HtmlUtil.RenderControl(list);
    }
    //----------------------------------------------------------------------
    //dropdownlist選縣市自動帶入地區
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
        if (ddlArea.Items.FindByText("鄉鎮市區").Selected == true)
            tbxZipCode.Text = "";
        else
            tbxZipCode.Text = Util.GetCityCode(ddlArea.SelectedItem.Text, ddlCity.SelectedValue).Substring(0, 3);
    }
    //----------------------------------------------------------------------
    //dropdownlist選城市自動帶入地區
    protected void ddlInvoice_City_SelectedIndexChanged(object sender, EventArgs e)
    {
        CityToArea(ddlInvoice_Area, ddlInvoice_City);
    }
    //----------------------------------------------------------------------
    //dropdownlist選地區自動填入郵遞區號
    protected void ddlInvoice_Area_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlInvoice_Area.Items.FindByText("鄉鎮市區").Selected == true)
            tbxInvoice_ZipCode.Text = "";
        else
            tbxInvoice_ZipCode.Text = Util.GetCityCode(ddlInvoice_Area.SelectedItem.Text, ddlInvoice_City.SelectedValue).Substring(0, 3);
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

                //海外地址全部清空  20150428取消清空
                cbxInvoice_IsAbroad.Checked = false;
                //tbxInvoice_OverseasCountry.Text = "";
                //tbxInvoice_OverseasAddress.Text = "";
            }
            if (cbxIsAbroad.Checked)
            {
                cbxInvoice_IsAbroad.Checked = true;
                //tbxInvoice_OverseasCountry.Text = tbxOverseasCountry.Text;
                ddlInvoice_OverseasCountry.SelectedValue = ddlOverseasCountry.SelectedValue;
                tbxInvoice_OverseasAddress.Text = tbxOverseasAddress.Text;

                //台灣本島全部清空  20150428取消清空
                cbxInvoice_IsLocal.Checked = false;
                //tbxInvoice_ZipCode.Text = "";
                //ddlInvoice_City.SelectedIndex = 0;
                //if (ddlInvoice_City.Items.FindByText("縣 市") == null)
                //{
                //    ddlInvoice_City.Items.Insert(0, new ListItem("縣 市", "縣 市"));
                //}
                //ddlInvoice_City.SelectedIndex = 0;
                //ddlInvoice_Area.SelectedIndex = 0;
                //tbxInvoice_Street.Text = "";
                //ddlInvoice_Section.SelectedIndex = 0;
                //tbxInvoice_Lane.Text = "";
                //tbxInvoice_Alley0.Text = "";
                //ddlInvoice_Alley.SelectedIndex = 0;
                //tbxInvoice_No1.Text = "";
                //tbxInvoice_No2.Text = "";
                //tbxInvoice_Floor1.Text = "";
                //tbxInvoice_Floor2.Text = "";
                //tbxInvoice_Room.Text = "";
                //tbxInvoice_Attn.Text = "";
            }

            // 2014/4/9 修改收據稱謂欄位可帶入捐款人稱謂欄位值
            ddlTitle2.SelectedItem.Text = ddlTitle.SelectedItem.Text;
            //2014/5/21 修改可帶入收據抬頭
            tbxInvoice_Title.Text = txtDonor_Name.Text;
        }
    }
    //----------------------------------------------------------------------
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        //20140509 修改 by Ian_Kao
        ////20140425 新增 by Ian_Kao 另外新增驗證項來驗證是否身分別有填入
        //if (!IsValid)
        //{
        //    AjaxShowMessage("身分別欄位不得為空白");
        //    return;
        //}
        if (Util.GetControlValue("GN_DonorType") == "")
        {
            AjaxShowMessage("身分別欄位不得為空白");
            return;
        }
        bool flag = false;
        try
        {
            Donor_Edit();
            // 2014/4/9 修改顯示訊息執行順序
            //SetSysMsg("修改資料成功");
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }

        if (flag == true)
        {
            // 2014/5/21 重新導入此頁面,身分別同步顯示至最新狀態
            if (hfInvoiceTitle.Value == tbxInvoice_Title.Text)
            {
                //SetSysMsg("修改資料成功！");
                ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('修改資料成功！');location.href=('DonorInfo_Edit.aspx?Donor_Id=" + HFD_Uid.Value + "');</script>");
            }
            else
            {
                //SetSysMsg("已建檔之收據資料應自行手動修改。\\n\\n修改資料成功！");
                ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('已建檔之收據資料應自行手動修改。\\n\\n修改資料成功！');location.href=('DonorInfo_Edit.aspx?Donor_Id=" + HFD_Uid.Value + "');</script>");
            }
            //Response.Redirect(Util.RedirectByTime("DonorInfo_Edit.aspx", "Donor_Id=" + HFD_Uid.Value));
            // 2014/4/9 修改用另外的顯示訊息函式，但不作用，回頭。
            //AjaxShowMessage("修改資料成功!");
            // 2014/4/7 修改不導向，保留原修改頁面
            //Response.Redirect(Util.RedirectByTime("DonorInfo.aspx"));
        }

    }
    protected void btnDel_Click(object sender, EventArgs e)
    {
        bool flag = false;
        string strSql = "";

        DataTable dt = null;
        Dictionary<string, object> dict_checkD = new Dictionary<string, object>();
        strSql = "select * from DONATE where Donor_Id=@Donor_Id and IsNull(Issue_Type,'') <> 'D' ";

        dict_checkD.Add("Donor_Id", this.HFD_Uid.Value);
        dt = NpoDB.GetDataTableS(strSql, dict_checkD);
        if (dt.Rows.Count != 0)
        {
            ClientScript.RegisterStartupScript(this.GetType(),"s", "<script>alert('此人尚有捐款記錄，請做捐款資料合併後，再刪除！！');</script>");
            return;
        }
        Dictionary<string, object> dict_checkP = new Dictionary<string, object>();
        strSql = "select * from Pledge where Donor_Id=@Donor_Id ";

        dict_checkP.Add("Donor_Id", this.HFD_Uid.Value);
        dt = NpoDB.GetDataTableS(strSql, dict_checkP);
        if (dt.Rows.Count != 0)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('此人尚有轉帳授權書記錄，請做轉帳授權書資料合併後，再刪除！！');</script>");
            return;
        }
        //20150623 增加公關贈品防呆
        Dictionary<string, object> dict_checkG = new Dictionary<string, object>();
        strSql = "select * from GiftData where Donor_Id=@Donor_Id ";

        dict_checkG.Add("Donor_Id", this.HFD_Uid.Value);
        dt = NpoDB.GetDataTableS(strSql, dict_checkG);
        if (dt.Rows.Count != 0)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('此人尚有公關贈品記錄，請做公關贈品資料合併後，再刪除！！');</script>");
            return;
        }
        try
        {
            //****變數宣告****//
            Dictionary<string, object> dict = new Dictionary<string, object>();

            //****設定SQL指令****//
            strSql = " update DONOR set ";
            strSql += " DeleteDate = GetDate()";
            strSql += ", LastUpdate_Dept_Id = @LastUpdate_Dept_Id";
            strSql += ", LastUpdate_Date = @LastUpdate_Date";
            strSql += ", LastUpdate_DateTime = @LastUpdate_DateTime";
            strSql += ", LastUpdate_User = @LastUpdate_User";
            strSql += ", LastUpdate_IP = @LastUpdate_IP";
            strSql += " where Donor_Id = @Donor_Id";

            dict.Add("LastUpdate_Dept_Id", SessionInfo.DeptID);
            dict.Add("LastUpdate_Date", DateTime.Now.ToString("yyyy-MM-dd"));
            dict.Add("LastUpdate_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
            dict.Add("LastUpdate_User", SessionInfo.UserName);
            dict.Add("LastUpdate_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
            dict.Add("Donor_Id", this.HFD_Uid.Value);

            //****執行語法****//
            NpoDB.ExecuteSQLS(strSql, dict);

            //20140416 紀錄Log檔 
            dict.Clear();
            dict.Add("Donor_Id", HFD_Uid.Value);
            string Donor_Name = NpoDB.GetScalarS("select Donor_Name from Donor where Donor_Id = @Donor_Id", dict);
            CaseUtil.InsertLogData(SessionInfo.UserID, "刪除資料", "刪除" + HFD_Uid.Value + Donor_Name);

            SetSysMsg("刪除資料成功!");
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            Response.Redirect(Util.RedirectByTime("DonorInfo.aspx"));
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("DonorInfo.aspx"));
    }
    protected void btnAddDonateData_Click(object sender, EventArgs e)
    {
        Session["cType"] = "DonorInfo_Edit";
        Response.Redirect(Util.RedirectByTime("../DonateMgr/Donate_Add.aspx", "Donor_Id=" + HFD_Uid.Value));
    }
    public void Donor_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update DONOR set ";
        strSql += "  Donor_Name = @Donor_Name";
        strSql += ", Sex = @Sex";
        strSql += ", title = @title";
        strSql += ", Donor_Type = @Donor_Type";
        strSql += ", IDNo = @IDNo";
        strSql += ", Birthday = @Birthday";
        strSql += ", Cellular_Phone = @Cellular_Phone";
        strSql += ", Tel_Office_Loc = @Tel_Office_Loc";
        strSql += ", Tel_Office = @Tel_Office";
        strSql += ", Tel_Office_Ext = @Tel_Office_Ext";
        strSql += ", Fax_Loc = @Fax_Loc";
        strSql += ", Fax = @Fax";
        strSql += ", Email = @Email";
        strSql += ", Contactor = @Contactor";
        //strSql += ", JobTitle = @JobTitle";
        strSql += ", IsAbroad=@IsAbroad";
        strSql += ", ZipCode= @ZipCode";
        strSql += ", City= @City";
        strSql += ", Area= @Area";
        strSql += ", Address= @Address";
        strSql += ", IsSendNews= @IsSendNews";
        strSql += ", IsSendNewsNum= @IsSendNewsNum";
        strSql += ", IsDVD= @IsDVD";
        strSql += ", IsSendEpaper= @IsSendEpaper";
        strSql += ", IsGift= @IsGift";
        strSql += ", IsBigAmtThank= @IsBigAmtThank";
        strSql += ", IsPost= @IsPost";
        strSql += ", IsContact= @IsContact";
        strSql += ", IsErrAddress= @IsErrAddress";
        strSql += ", Invoice_Type= @Invoice_Type";
        strSql += ", Title2= @Title2";
        strSql += ", Invoice_Title= @Invoice_Title";
        strSql += ", Invoice_IDNo= @Invoice_IDNo";
        strSql += ", IsAbroad_Invoice= @IsAbroad_Invoice";
        strSql += ", IsAnonymous= @IsAnonymous";
        strSql += ", Remark = @Remark ";
        strSql += ", ToGOODTV = @ToGOODTV ";
        strSql += ", IsFdc= @IsFdc";
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
        strSql += ", Invoice_Street= @Invoice_Street";
        strSql += ", Invoice_Section= @Invoice_Section";
        strSql += ", Invoice_Lane= @Invoice_Lane";
        strSql += ", Invoice_Alley= @Invoice_Alley";
        strSql += ", Invoice_HouseNo= @Invoice_HouseNo";
        strSql += ", Invoice_HouseNoSub= @Invoice_HouseNoSub";
        strSql += ", Invoice_Floor= @Invoice_Floor";
        strSql += ", Invoice_FloorSub= @Invoice_FloorSub";
        strSql += ", Invoice_Room= @Invoice_Room";
        strSql += ", Invoice_Attn= @Invoice_Attn";
        strSql += ", Invoice_OverseasCountry= @Invoice_OverseasCountry";
        strSql += ", Invoice_OverseasAddress= @Invoice_OverseasAddress";
        strSql += ", Invoice_ZipCode= @Invoice_ZipCode";
        strSql += ", Invoice_City= @Invoice_City";
        strSql += ", Invoice_Area= @Invoice_Area";
        strSql += ", Invoice_Address= @Invoice_Address";
        strSql += ", LastUpdate_Dept_Id = @LastUpdate_Dept_Id";
        strSql += ", LastUpdate_Date = @LastUpdate_Date";
        strSql += ", LastUpdate_DateTime = @LastUpdate_DateTime";
        strSql += ", LastUpdate_User = @LastUpdate_User";
        strSql += ", LastUpdate_IP = @LastUpdate_IP";
        strSql += ", IsMember = @IsMember";
        strSql += ", OTT_Address = @OTT_Address";
        strSql += " where Donor_Id = @Donor_Id";
        dict.Add("Donor_Name", txtDonor_Name.Text.Trim());
        dict.Add("Sex", ddlSex.SelectedItem.Text);
        dict.Add("Title", ddlTitle.SelectedItem.Text);
        //20140509 修改 by Ian_Kao
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
        //國內通訊地址
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
        string Address1 = "";
        if (tbxStreet.Text.Trim() != "")
        {
            Address1 += tbxStreet.Text.Trim();
            dict.Add("Street", tbxStreet.Text.Trim());
        }
        else
        {
            dict.Add("Street", "");
        }
        if (ddlSection.SelectedIndex != 0)//有無段別
        {
            Address1 += ddlSection.SelectedItem.Text + "段";
            dict.Add("Section", ddlSection.SelectedItem.Text);
        }
        else
        {
            dict.Add("Section", "");
        }
        if (tbxLane.Text.Trim() != "")//有無巷
        {
            Address1 += tbxLane.Text.Trim() + "巷";
            dict.Add("Lane", tbxLane.Text.Trim());
        }
        else
        {
            dict.Add("Lane", "");
        }
        if (tbxAlley.Text.Trim() != "")//有無弄衖
        {
            Address1 += tbxAlley.Text.Trim() + ddlAlley.SelectedItem.Text;
            dict.Add("Alley", tbxAlley.Text.Trim());
        }
        else
        {
            dict.Add("Alley", "");
        }
        //2015/08/12 更正 <No1>號之<No2>
        if (tbxNo1.Text.Trim() != "")//有無號碼
        {
            Address1 += tbxNo1.Text.Trim() + "號";
            dict.Add("HouseNo", tbxNo1.Text.Trim());
        }
        else
        {
            dict.Add("HouseNo", "");
        }
        if (tbxNo2.Text.Trim() != "")//之幾
        {
            Address1 += "之" + tbxNo2.Text.Trim() + " ";
            dict.Add("HouseNoSub", tbxNo2.Text.Trim());
        }
        else
        {
            dict.Add("HouseNoSub", "");
        }
        // 2014/4/8 更正 <Floor1>-<Floor2>樓 --> <Floor1>樓之<Floor2>
        if (tbxFloor1.Text.Trim() != "" && tbxFloor2.Text.Trim() == "")//有無樓別
        {
            Address1 += tbxFloor1.Text.Trim() + "樓";
            dict.Add("Floor", tbxFloor1.Text.Trim());
            dict.Add("FloorSub", "");
        }
        else if (tbxFloor1.Text.Trim() == "" && tbxFloor2.Text.Trim() != "")
        {
            //Address += tbxFloor2.Text.Trim() + "樓";
            Address1 += "樓之" + tbxFloor2.Text.Trim();
            dict.Add("Floor", "");
            dict.Add("FloorSub", tbxFloor2.Text.Trim());
        }
        else if (tbxFloor1.Text.Trim() != "" && tbxFloor2.Text.Trim() != "")
        {
            //Address += tbxFloor1.Text.Trim() + "-" + tbxFloor2.Text.Trim() + "樓";
            Address1 += tbxFloor1.Text.Trim() + "樓之" + tbxFloor2.Text.Trim();
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
            Address1 += " " + tbxRoom.Text.Trim() + "室";
            dict.Add("Room", tbxRoom.Text.Trim());
        }
        else
        {
            dict.Add("Room", "");
        }
        if (tbxAttn.Text.Trim() != "")
        {
            dict.Add("Attn", tbxAttn.Text.Trim());
        }
        else
        {
            dict.Add("Attn", "");
        }
        //國內
        if (this.cbxIsLocal.Checked)
        {
            dict.Add("IsAbroad", "N");
            dict.Add("Address", Address1);
        }
        //國外通訊地址
        string Address2 = "";
        if (tbxOverseasAddress.Text.Trim() != "")//有無其他
        {
            Address2 += tbxOverseasAddress.Text.Trim() + " ";
            dict.Add("OverseasAddress", tbxOverseasAddress.Text.Trim());
        }
        else
        {
            dict.Add("OverseasAddress", "");
        }
        if (ddlOverseasCountry.SelectedIndex != 0)//有無國家/省城市/區
        {
            Address2 += ddlOverseasCountry.SelectedValue;
            dict.Add("OverseasCountry", ddlOverseasCountry.SelectedValue);
        }
        else
        {
            dict.Add("OverseasCountry", "");
        }
        //國外
        if (this.cbxIsAbroad.Checked)
        {
            dict.Add("IsAbroad", "Y");
            dict.Add("Address", Address2);
        }
        //文宣品
        dict.Add("IsSendNewsNum", tbxIsSendNewsNum.Text.Trim());
        if (tbxIsSendNewsNum.Text.Trim() == "" || tbxIsSendNewsNum.Text.Trim() == "0")
        {
            dict.Add("IsSendNews", "N");
        }
        else
        {
            dict.Add("IsSendNews", "Y");
        }
        if (this.cbxIsDVD.Checked)
        {
            dict.Add("IsDVD", "Y");
        }
        else
        {
            dict.Add("IsDVD", "N");
        }
        if (this.cbxIsSendEpaper.Checked)
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
        if (this.cbxIsContact.Checked)
        {
            dict.Add("IsContact", "N");
        }
        else
        {
            dict.Add("IsContact", "Y");
        }
        if (this.cbxIsErrAddress.Checked)
        {
            dict.Add("IsErrAddress", "Y");
        }
        else
        {
            dict.Add("IsErrAddress", "N");
        }
        dict.Add("Invoice_Type", ddlInvoice_Type.SelectedItem.Text);
        //國內或國外都沒勾選
        if (cbxInvoice_IsLocal.Checked == false && cbxInvoice_IsAbroad.Checked == false)
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
        }else
        {
            //國內收據地址
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

            string Address3 = "";
            if (tbxInvoice_Street.Text.Trim() != "")//有無街/大道
            {
                Address3 += tbxInvoice_Street.Text.Trim();
                dict.Add("Invoice_Street", tbxInvoice_Street.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_Street", "");
            }
            if (ddlInvoice_Section.SelectedIndex != 0)//有無段別
            {
                Address3 += ddlInvoice_Section.SelectedItem.Text + "段";
                dict.Add("Invoice_Section", ddlInvoice_Section.SelectedItem.Text);
            }
            else
            {
                dict.Add("Invoice_Section", "");
            }
            if (tbxInvoice_Lane.Text.Trim() != "")//有無巷
            {
                Address3 += tbxInvoice_Lane.Text.Trim() + "巷";
                dict.Add("Invoice_Lane", tbxInvoice_Lane.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_Lane", "");
            }
            if (tbxInvoice_Alley0.Text.Trim() != "")//有無弄衖
            {
                Address3 += tbxInvoice_Alley0.Text.Trim() + ddlInvoice_Alley.SelectedItem.Text;
                dict.Add("Invoice_Alley", tbxInvoice_Alley0.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_Alley", "");
            }
            //2015/08/12 更正 <No1>號之<No2>
            if (tbxInvoice_No1.Text.Trim() != "")//有無號碼
            {
                Address3 += tbxInvoice_No1.Text.Trim() + "號";
                dict.Add("Invoice_HouseNo", tbxInvoice_No1.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_HouseNo", "");
            }
            if (tbxInvoice_No2.Text.Trim() != "")
            {
                Address3 += "之" + tbxInvoice_No2.Text.Trim() + " ";
                dict.Add("Invoice_HouseNoSub", tbxInvoice_No2.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_HouseNoSub", "");
            }
            // 2014/4/8 更正 <Floor1>-<Floor2>樓 --> <Floor1>樓之<Floor2>
            if (tbxInvoice_Floor1.Text.Trim() != "" && tbxInvoice_Floor2.Text.Trim() == "")//有無樓別
            {
                Address3 += tbxInvoice_Floor1.Text.Trim() + "樓";
                dict.Add("Invoice_Floor", tbxInvoice_Floor1.Text.Trim());
                dict.Add("Invoice_FloorSub", "");
            }
            else if (tbxInvoice_Floor1.Text.Trim() == "" && tbxInvoice_Floor2.Text.Trim() != "")
            {
                //Address += tbxInvoice_Floor2.Text.Trim() + "樓";
                Address3 += "樓之" + tbxInvoice_Floor2.Text.Trim();
                dict.Add("Invoice_Floor", "");
                dict.Add("Invoice_FloorSub", tbxInvoice_Floor2.Text.Trim());
            }
            else if (tbxInvoice_Floor1.Text.Trim() != "" && tbxInvoice_Floor2.Text.Trim() != "")
            {
                //Address += tbxInvoice_Floor1.Text.Trim() + "-" + tbxInvoice_Floor2.Text.Trim() + "樓";
                Address3 += tbxInvoice_Floor1.Text.Trim() + "樓之" + tbxInvoice_Floor2.Text.Trim();
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
                Address3 += " " + tbxInvoice_Room.Text.Trim() + "室";
                dict.Add("Invoice_Room", tbxInvoice_Room.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_Room", "");
            }
            if (tbxInvoice_Attn.Text.Trim() != "")
            {
                dict.Add("Invoice_Attn", tbxInvoice_Attn.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_Attn", "");
            }
            //國內
            if (this.cbxInvoice_IsLocal.Checked)
            {
                dict.Add("IsAbroad_Invoice", "N");
                dict.Add("Invoice_Address", Address3);
            }
            //國外收據地址
            string Address4 = "";
            if (tbxInvoice_OverseasAddress.Text.Trim() != "")//有無其他
            {
                Address4 += tbxInvoice_OverseasAddress.Text.Trim() + " ";
                dict.Add("Invoice_OverseasAddress", tbxInvoice_OverseasAddress.Text.Trim());
            }
            else
            {
                dict.Add("Invoice_OverseasAddress", "");
            }
            if (ddlInvoice_OverseasCountry.SelectedIndex != 0)//有無國家/省城市/區
            {
                Address4 += ddlInvoice_OverseasCountry.SelectedValue;
                dict.Add("Invoice_OverseasCountry", ddlInvoice_OverseasCountry.SelectedValue);
            }
            else
            {
                dict.Add("Invoice_OverseasCountry", "");
            }
            //國外
            if (this.cbxInvoice_IsAbroad.Checked)
            {
                dict.Add("IsAbroad_Invoice", "Y");
                dict.Add("Invoice_Address", Address4);
            }
        }
        dict.Add("Title2", ddlTitle2.SelectedItem.Text);
        dict.Add("Invoice_Title", tbxInvoice_Title.Text.Trim());
        dict.Add("Invoice_IDNo", tbxInvoice_IDNo.Text.Trim());
        //dict.Add("IsAnonymous", ddlIsAnonymous.SelectedItem.Text);
        if (this.cbxIsAnonymous.Checked)
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
        dict.Add("LastUpdate_Dept_Id", SessionInfo.DeptID);
        dict.Add("LastUpdate_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("LastUpdate_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        dict.Add("LastUpdate_User", SessionInfo.UserName);
        dict.Add("LastUpdate_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        dict.Add("IsMember", "N");
        dict.Add("OTT_Address", tbxOTT_Address.Text);
        dict.Add("Donor_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    //20140127增加性別與稱謂連動
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

        /* 會有排序問題而對應錯誤。
        if (ddlSex.SelectedIndex == 0)
        {
            ddlTitle.SelectedIndex = 4;
        }
        if (ddlSex.SelectedIndex == 1)
        {
            ddlTitle.SelectedIndex = 2;
        }
        if (ddlSex.SelectedIndex == 2)
        {
            ddlTitle.SelectedIndex = 3;
        }
        */

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