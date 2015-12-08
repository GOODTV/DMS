using System;
using System.Collections.Generic;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonorGroupMgr_GroupItemMember : BasePage
{
    GroupAuthrity ga = null;

    protected void Page_Load(object sender, EventArgs e)
    {

        //ScriptManager1.RegisterAsyncPostBackControl(btnQuery);
        //ScriptManager1.RegisterAsyncPostBackControl(btnDonorQuery);

        if (!IsPostBack)
        {

            Session["ProgID"] = "GroupItem";
            //權控處理
            AuthrityControl();
           
            LoadDropDownListData();

            if (String.IsNullOrEmpty(Request.Params["GroupItemUid"]))
            {
                lblGroupItemGridList.Text = "<div align='center'>** 請先輸入查詢條件 **</div>";
                //btnExcel.Enabled = false;
            }
            else
            {
                hfGroupItemUid.Value = Request.Params["GroupItemUid"].ToString();
                //txtGroupItemUid.Text = Request.Params["GroupItemUid"].ToString();
                LoadFormData();
                //txtGroupItemUid.Text = "";
            }

            //Session["strSql"] = "";
            DonorLoadDropDownListData();
            lblDonorGridList.Text = "<div align='center'>** 請先輸入查詢條件 **</div>";
            PanelAbroad.Visible = false;

            if (Request.Cookies["GroupItemHelp"] != null)
            {
                hfHelp.Value = Request.Cookies["GroupItemHelp"].Value.ToString();
            }

            Session["Donor_Id"] = "";
            Session["Donor_Id_Old"] = "";
            Session["Tel_Office"] = "";
            Session["ZipCode"] = "";
            Session["Street"] = "";
            Session["Lane"] = "";
            Session["Alley"] = "";
            Session["HouseNo"] = "";
            Session["HouseNoSub"] = "";
            Session["Floor"] = "";
            Session["FloorSub"] = "";
            Session["Room"] = "";
            Session["IsSendNewsNum"] = "";
            Session["ClickQuery"] = "";
            //20140425 新增 by Ian_Kao
            Session["Donor_Name"] = "";
            Session["Invoice_Title"] = "";
            Session["Email"] = "";

        }
    }

    //----------------------------------------------------------------------
    public void AuthrityControl()
    {
        ga = Authrity.GetGroupRight();
        if (Authrity.CheckPageRight(ga.Focus) == false)
        {
            return;
        }
        btnQuery.Visible = ga.Query;
        btnDonorQuery.Visible = ga.Query;
    }

    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        string strSql = @"
                          select uid, GroupClassName
                          from GroupClass gc
                          where DeleteDate is null
                         ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        Util.FillDropDownList(ddlGroupClass, dt, "GroupClassName", "uid", true);
    }

    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        // 增加查詢條件顯示
        string strSQLWhere = "";

        string strSql = "";
        strSql = @"
                   select ROW_NUMBER() OVER(ORDER BY gi.uid) as [序號] ,gi.uid,
                   gc.GroupClassName as 群組類別,
                   gi.GroupItemName as 群組代表,
                   gi.Supplement as 備註
                   from GroupItem gi
                   left join GroupClass gc on gi.GroupClassUid=gc.uid
                   where gi.DeleteDate is null
                  ";

        if (ddlGroupClass.SelectedValue != "")
        {
            strSql += "and gi.GroupClassUID=@GroupClassUID ";
            strSQLWhere += "群組類別：" + ddlGroupClass.SelectedItem.ToString();
        }
        if (txtGroupItemName.Text != "")
        {
            strSql += "and gi.GroupItemName like @GroupItemName ";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "群組代表：" + txtGroupItemName.Text.Trim();
        }
        if (hfGroupItemUid.Value != "")
        {
            strSql += "and gi.uid = @GroupItemUid ";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "群組代表編號：" + hfGroupItemUid.Value;
        }
        if (txtSupplement.Text != "")
        {
            strSql += "and gi.Supplement like @Supplement ";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "備註：" + txtSupplement.Text.Trim();
        }
        if (txtDonorID.Text.Trim() != "")
        {
            strSql += "and gi.uid in (SELECT GroupItemUid FROM GroupMapping where DeleteDate is null and DonorID = @DonorID)  ";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "捐款人編號：" + txtDonorID.Text.Trim();
        }

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("GroupClassUID", ddlGroupClass.SelectedValue);
        dict.Add("GroupItemName", "%" + txtGroupItemName.Text.Trim() + "%");
        dict.Add("Supplement", "%" + txtSupplement.Text.Trim() + "%");
        dict.Add("DonorID", txtDonorID.Text.Trim());
        dict.Add("GroupItemUid", hfGroupItemUid.Value);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        /*
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("uid");
        npoGridView.DisableColumn.Add("uid");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
        npoGridView.EditLink = Util.RedirectByTime("GroupItem_Edit.aspx", "GroupItemUID=");
        lblGridList.Text = npoGridView.Render();
        */

        lblGroupItemGridList.Text = "";

        int intRowsCount = dt.Rows.Count;

        if (intRowsCount > 0 && intRowsCount <= 100)
        {

            string strBody = "<table class='table_h' width='100%'>";
            strBody += @"<TR><TH noWrap><SPAN>序號</SPAN></TH>
                            <TH noWrap><SPAN>群組類別</SPAN></TH>
                            <TH noWrap><SPAN>群組代表</SPAN></TH>
                            <TH noWrap><SPAN>備註</SPAN></TH>
                            <TH style='white-space: normal;'><SPAN>成員列表 [捐款人編號(捐款人名稱)]</SPAN></TH>
                        </TR>";

            foreach (DataRow dr in dt.Rows)
            {

                string strItemUid = dr["uid"].ToString().Trim();

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
                    strMember = "";
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
                        //strMember += (strMember == "" ? "" : ", ") + "<SPAN onclick=\"window.event.cancelBubble=true;window.open('" +
                        strMember += "<SPAN onclick=\"window.event.cancelBubble=true;window.open('" +
                            Util.RedirectByTime("/DonorMgr/DonorInfo_Edit.aspx", "Donor_Id=" + dr2["Donor_Id"].ToString().Trim()) +
                            "','_self','')\" title='" + strMemberTitle + "'><font color='blue'>" + dr2["Donor_Id"].ToString().Trim() + "</font>(" + dr2["Donor_Name"].ToString().Trim() + ")</SPAN>";


                    }

                }

                //strBody += "<TR style=\"cursor: pointer; cursor: hand;\" onclick =\"window.open('" + Util.RedirectByTime("GroupItem_Edit.aspx", "GroupItemUID=" + strItemUid) + "','_self','')\">";
                strBody += "<TR id='" + strItemUid + "'>";
                strBody += "<TD noWrap align='right'><SPAN>" + dr["序號"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["群組類別"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["群組代表"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["備註"].ToString() + "</SPAN></TD>";
                strBody += "<TD>" + strMember + "&nbsp;</TD></TR>";
            }

            strBody += "</table>";
            lblGroupItemGridList.Text = "<font color='blue'>【查詢條件】</font>" + (strSQLWhere == "" ? "全部" : strSQLWhere) + " <font color='blue'>【查詢筆數】</font>" + intRowsCount + "筆<div id='divGroupItem' style='overflow: auto;'>" + strBody + "</div>";
            Session["GridList"] = lblGroupItemGridList.Text;

        }
        else if (intRowsCount > 100)
        {
            lblGroupItemGridList.Text = "<div align='center' style='color: red;'>** 查詢出有" + intRowsCount + "筆，已超過100筆，請增加查詢條件來縮小資料的範圍，以利作業順暢。 **</div>";
        }
        else
        {
            lblGroupItemGridList.Text = "<div align='center'>** 沒有符合條件的資料 **</div>";
        }

    }

    //----------------------------------------------------------------------
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("GroupItem_Edit.aspx?Mode=ADD&", ""));
    }

    //----------------------------------------------------------------------
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        //UpdatePanel2.Update();
        LoadFormData();
        //setGroupItemHelp(hfHelp.Value);
        //btnExcel.Enabled = true;
    }

    //----------------------------------------------------------------------
    protected void btnExcel_Click(object sender, EventArgs e)
    {
        Response.Redirect("GroupItemQry_Excel.aspx");
    }

    //---------------------------------------------------------------------------
    public void DonorLoadDropDownListData()
    {

        Util.FillDropDownList(ddlCity, Util.GetDataTable("CodeCity", "ParentCityID", "0", "", ""), "Name", "ZipCode", false);
        ddlCity.Items.Insert(0, new ListItem("縣 市", "縣 市"));
        ddlCity.SelectedIndex = 0;

        ddlArea.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea.SelectedIndex = 0;

        Util.FillDropDownList(ddlSection, Util.GetDataTable("CaseCode", "GroupName", "段別", "", ""), "CaseName", "CaseID", false);
        ddlSection.Items.Insert(0, new ListItem("請選擇", "請選擇"));
        ddlSection.SelectedIndex = 0;

        Util.FillDropDownList(ddlAlley, Util.GetDataTable("CaseCode", "GroupName", "弄衖別", "", ""), "CaseName", "CaseID", false);
        ddlAlley.Items.Insert(0, new ListItem("請選擇", "請選擇"));
        ddlAlley.SelectedIndex = 0;

    }
 
    //----------------------------------------------------------------------
    public void btnDonorQuery_Click(object sender, EventArgs e)
    {
        //UpdatePanel3.Update();
        // SQL-> AND串條件
        Query();
        setGroupItemHelp(hfHelp.Value);
        Session["ClickQuery"] = "Query";
    }

    //---------------------------------------------------------------------------
    public void Query()
    {
        string strSql;
        DataTable dt;
        strSql = " select Distinct Donor_Id , Donor_Id as 編號, Donor_Name as 捐款人, (Case When IsMember='Y' Then 'V' Else '' End) as 讀者, Donor_Type as 身分別, Tel_Office as 連絡電話, Cellular_Phone as 手機號碼, ";
        strSql += " (Case When IsAbroad = 'Y' Then DONOR.OverseasAddress + DONOR.OverseasCountry Else Case When ISNULL(DONOR.City,'')='' Then Address Else Case When ISNULL(A.mValue,'')<>ISNULL(B.mValue,'') Then ISNULL(DONOR.ZipCode,'')+A.mValue+ISNULL(B.mValue,'')+Address End End End) as [通訊地址] ";
        strSql += " From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode ";
        strSql += " where DeleteDate is null ";

        Dictionary<string, object> dict = new Dictionary<string, object>();

        // 增加查詢條件顯示
        string strSQLWhere = "";

        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140425 修改 by Ian_Kao
            //strSql += "and (Donor_Name Like N'%" + tbxDonor_Name.Text.Trim() + "%' Or NickName  Like N'%" + tbxDonor_Name.Text.Trim() + "%' Or Contactor  Like N'%" + tbxDonor_Name.Text.Trim() + "%' Or Invoice_Title Like N'%" + tbxDonor_Name.Text.Trim() + "%')";
            strSql += "and (Donor_Name Like @Donor_Name Or Contactor  Like @Donor_Name Or Invoice_Title Like @Donor_Name)";
            dict.Add("Donor_Name", "%" + tbxDonor_Name.Text.Trim() + "%");
            Session["Donor_Name"] = "%" + tbxDonor_Name.Text.Trim() + "%";
            strSQLWhere += "捐款人：" + tbxDonor_Name.Text.Trim();
        }
        if (tbxDonor_Id.Text.Trim() != "")
        {
            strSql += "and Donor_Id=@Donor_Id ";
            dict.Add("Donor_Id", tbxDonor_Id.Text.Trim());
            Session["Donor_Id"] = tbxDonor_Id.Text.Trim();
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "捐款人編號：" + tbxDonor_Id.Text.Trim();
        }
        //20140127新增
        if (tbxInvoice_Title.Text.Trim() != "")
        {
            //20140425 修改 by Ian_Kao
            //strSql += "and Invoice_Title LIKE N'%" + tbxInvoice_Title.Text.Trim() + "%' ";
            strSql += "and Invoice_Title LIKE @Invoice_Title ";
            dict.Add("Invoice_Title", "%" + tbxInvoice_Title.Text.Trim() + "%");
            Session["Invoice_Title"] = "%" + tbxInvoice_Title.Text.Trim() + "%";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "收據抬頭：" + tbxInvoice_Title.Text.Trim();
        }
        if (tbxTel_Office.Text.Trim() != "")
        {
            //strSql += "and Tel_Office=@Tel_Office ";
            //dict.Add("Tel_Office", tbxTel_Office.Text.Trim());
            //20140807 修正可查詢電話、手機
            strSql += "and (Tel_Office LIKE '%" + tbxTel_Office.Text.Trim() + "%' or Tel_Home LIKE '%" + tbxTel_Office.Text.Trim() + "%' or Cellular_Phone LIKE '%" + tbxTel_Office.Text.Trim() + "%') ";
            Session["Tel_Office"] = tbxTel_Office.Text.Trim();
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "連絡電話：" + tbxTel_Office.Text.Trim();
        }
        string strWhere = "";
        if (cbxIsAbroad.Checked)
        {
            strSql += "and (IsAbroad = 'Y' Or IsAbroad_Invoice = 'Y') ";
            if (tbxIsAbroad_Address.Text.Trim() != "")
            {
                //20140425 修改 by Ian_Kao
                //strSql += "and 通訊地址 LIKE N'%" + tbxIsAbroad_Address.Text.Trim() + "%' ";
                strSql += "and (OverseasAddress LIKE N'%" + tbxIsAbroad_Address.Text.Trim() + "%' or OverseasCountry LIKE N'%" + tbxIsAbroad_Address.Text.Trim() + "%' )";
                strWhere += (strWhere == "" ? tbxIsAbroad_Address.Text.Trim() : "," + tbxIsAbroad_Address.Text.Trim());
            }
        }
        if (cbxIsAbroad.Checked == false)
        {
            if (tbxZipCode.Text.Trim() != "")
            {
                strSql += "and B.ZipCode=@ZipCode ";
                dict.Add("ZipCode", tbxZipCode.Text.Trim());
                Session["ZipCode"] = tbxZipCode.Text.Trim();
                strWhere += (strWhere == "" ? tbxZipCode.Text.Trim() : "," + tbxZipCode.Text.Trim());
            }
            if (ddlCity.SelectedIndex != 0)
            {
                strSql += "and City='" + ddlCity.SelectedValue + "' ";
                strWhere += (strWhere == "" ? ddlCity.SelectedValue : "," + ddlCity.SelectedValue);
            }
            if (ddlArea.SelectedIndex != 0)
            {
                strSql += "and Area='" + ddlArea.SelectedValue + "' ";
            }
            if (tbxStreet.Text.Trim() != "")
            {
                strSql += "and Street=@Street ";
                dict.Add("Street", tbxStreet.Text.Trim());
                Session["Street"] = tbxStreet.Text.Trim();
                strWhere += (strWhere == "" ? tbxStreet.Text.Trim() : "," + tbxStreet.Text.Trim());
            }
            if (ddlSection.SelectedIndex != 0)
            {
                strSql += "and Section='" + ddlSection.SelectedItem.Text + "' ";
                strWhere += (strWhere == "" ? ddlSection.SelectedItem.Text : "," + ddlSection.SelectedItem.Text);
            }
            if (tbxLane.Text.Trim() != "")
            {
                strSql += "and Lane=@Lane ";
                dict.Add("Lane", tbxLane.Text.Trim());
                Session["Lane"] = tbxLane.Text.Trim();
                strWhere += (strWhere == "" ? tbxLane.Text.Trim() : "," + tbxLane.Text.Trim());
            }
            if (tbxAlley.Text.Trim() != "")
            {
                strSql += "and Alley=@Alley ";
                dict.Add("Alley", tbxAlley.Text.Trim());
                Session["Alley"] = tbxAlley.Text.Trim();
                strWhere += (strWhere == "" ? tbxAlley.Text.Trim() : "," + tbxAlley.Text.Trim());
            }
            if (tbxNo1.Text.Trim() != "")
            {
                strSql += "and HouseNo=@HouseNo ";
                dict.Add("HouseNo", tbxNo1.Text.Trim());
                Session["HouseNo"] = tbxNo1.Text.Trim();
                strWhere += (strWhere == "" ? tbxNo1.Text.Trim() : "," + tbxNo1.Text.Trim());
            }
            if (tbxNo2.Text.Trim() != "")
            {
                strSql += "and HouseNoSub=@HouseNoSub ";
                dict.Add("HouseNoSub", tbxNo2.Text.Trim());
                Session["HouseNoSub"] = tbxNo2.Text.Trim();
                strWhere += (strWhere == "" ? tbxNo2.Text.Trim() : "," + tbxNo2.Text.Trim());
            }
            if (tbxFloor1.Text.Trim() != "")
            {
                strSql += "and Floor=@Floor ";
                dict.Add("Floor", tbxFloor1.Text.Trim());
                Session["Floor"] = tbxFloor1.Text.Trim();
                strWhere += (strWhere == "" ? tbxFloor1.Text.Trim() : "," + tbxFloor1.Text.Trim());
            }
            if (tbxFloor2.Text.Trim() != "")
            {
                strSql += "and FloorSub=@FloorSub ";
                dict.Add("FloorSub", tbxFloor2.Text.Trim());
                Session["FloorSub"] = tbxFloor2.Text.Trim();
                strWhere += (strWhere == "" ? tbxFloor2.Text.Trim() : "," + tbxFloor2.Text.Trim());
            }
            if (tbxRoom.Text.Trim() != "")
            {
                strSql += "and Room=@Room ";
                dict.Add("Room", tbxRoom.Text.Trim());
                Session["Room"] = tbxRoom.Text.Trim();
                strWhere += (strWhere == "" ? tbxRoom.Text.Trim() : "," + tbxRoom.Text.Trim());
            }
        }
        strSql += "order by Donor_Id desc";
        dt = NpoDB.GetDataTableS(strSql, dict);

        if (!String.IsNullOrEmpty(strWhere))
        {
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "地址";   //：" + strWhere;
        }

        int intRowsCount = dt.Rows.Count;
        if (intRowsCount == 0)
        {
            lblDonorGridList.Text = "<div align='center'>** 沒有符合條件的資料 **</div>";
        }
        else if (intRowsCount > 100)
        {
            lblDonorGridList.Text = "<div align='center' style='color: red;'>** 查詢出有" + intRowsCount + "筆，已超過100筆，請增加查詢條件來縮小資料的範圍，以利作業順暢。 **</div>";
        }
        else
        {
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.Keys.Add("Donor_Id");
            npoGridView.DisableColumn.Add("Donor_Id");
            npoGridView.ShowPage = false;
            //npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            //npoGridView.EditLink = Util.RedirectByTime("../DonorMgr/DonorInfo_Edit.aspx", "Donor_Id=");
            //lblDonorGridList.Text = npoGridView.Render();
            lblDonorGridList.Text = "<font color='blue'>【查詢條件】</font>" + (strSQLWhere == "" ? "全部" : strSQLWhere) + " <font color='blue'>【查詢筆數】</font>" + intRowsCount + "筆<div id='divDonor' style='overflow: auto;'>" + npoGridView.Render() + "</div>";

            //Session["strSql"] = strSql;
        }

    }

    //地址海外選項異動
    protected void cbxIsAbroad_CheckedChanged(object sender, EventArgs e)
    {
        if (cbxIsAbroad.Checked)
        {
            PanelAbroad.Visible = true;
            PanelLocal.Visible = false;
        }
        else
        {
            PanelAbroad.Visible = false;
            PanelLocal.Visible = true;
        }
    }

    //dropdownlist選城市自動帶入地區
    protected void ddlCity_SelectedIndexChanged(object sender, EventArgs e)
    {
        Util.FillDropDownList(ddlArea, Util.GetDataTable("CodeCity", "ParentCityID", ddlCity.SelectedValue, "", ""), "Name", "ZipCode", false);
        ddlArea.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea.SelectedIndex = 0;
        tbxZipCode.Text = "";
    }

    //dropdownlist選地區自動填入郵遞區號
    protected void ddlArea_SelectedIndexChanged(object sender, EventArgs e)
    {
        tbxZipCode.Text = Util.GetCityCode(ddlArea.SelectedItem.Text, ddlCity.SelectedValue).Substring(0, 3);
    }

    protected void setGroupItemHelp(string HelpValue)
    {

        System.Web.HttpCookie GroupItemHelp = new System.Web.HttpCookie("GroupItemHelp");
        // Set the cookie value.
        GroupItemHelp.Value = HelpValue;
        // Set the cookie expiration date.
        GroupItemHelp.Expires = DateTime.Now.AddDays(7);
        // Add the cookie.
        Response.Cookies.Add(GroupItemHelp);

    }

}
