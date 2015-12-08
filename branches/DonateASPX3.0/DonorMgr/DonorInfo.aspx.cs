using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
/*
 * 群組名稱設定
 * 資料庫: DONOR
 * 相關程式:
 * List頁面: DONORInfo.aspx
 * 新增頁面: DONORInfo_Add.aspx 
 * 修改頁面: DONORInfo_Edit.aspx
 */
public partial class DonorMgr_DonatorInfo : BasePage
{
    #region NpoGridView 處理換頁相關程式碼
    Button btnNextPage, btnPreviousPage, btnGoPage;
    HiddenField HFD_CurrentPage, HFD_CurrentQuerye;

    override protected void OnInit(EventArgs e)
    {
        CreatePageControl();
        base.OnInit(e);
    }
    private void CreatePageControl()
    {
        // Create dynamic controls here.
        btnNextPage = new Button();
        btnNextPage.ID = "btnNextPage";
        Form1.Controls.Add(btnNextPage);
        btnNextPage.Click += new System.EventHandler(btnNextPage_Click);

        btnPreviousPage = new Button();
        btnPreviousPage.ID = "btnPreviousPage";
        Form1.Controls.Add(btnPreviousPage);
        btnPreviousPage.Click += new System.EventHandler(btnPreviousPage_Click);

        btnGoPage = new Button();
        btnGoPage.ID = "btnGoPage";
        Form1.Controls.Add(btnGoPage);
        btnGoPage.Click += new System.EventHandler(btnGoPage_Click);

        HFD_CurrentPage = new HiddenField();
        HFD_CurrentPage.Value = "1";
        HFD_CurrentPage.ID = "HFD_CurrentPage";
        Form1.Controls.Add(HFD_CurrentPage);

        HFD_CurrentQuerye = new HiddenField();
        HFD_CurrentQuerye.Value = "Query";
        HFD_CurrentQuerye.ID = "HFHFD_CurrentQuerye";
        Form1.Controls.Add(HFD_CurrentQuerye);
    }
    protected void btnPreviousPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.MinusStringNumber(HFD_CurrentPage.Value);
        if (Session["ClickQuery"].ToString() == "Query")
        {
            Query();
        }
        else
        {
            FuzzyQuery();
        }
    }
    protected void btnNextPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.AddStringNumber(HFD_CurrentPage.Value);
        if (Session["ClickQuery"].ToString() == "Query")
        {
            Query();
        }
        else
        {
            FuzzyQuery();
        }
    }
    protected void btnGoPage_Click(object sender, EventArgs e)
    {
        if (Session["ClickQuery"].ToString() == "Query")
        {
            Query();
        }
        else
        {
            FuzzyQuery();
        }
    }
    #endregion NpoGridView 處理換頁相關程式碼
    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "DonorInfo";
        //權控處理
        AuthrityControl();

        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            Session["strSql"] = "";
            LoadDropDownListData();
            lblGridList.Text = "** 請先輸入查詢條件 **";
            PanelAbroad.Visible = false ;

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
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_AddNew", btnAdd);
        Authrity.CheckButtonRight("_Query", btnQuery);
        Authrity.CheckButtonRight("_Query", btnFuzzyQuery);
        Authrity.CheckButtonRight("_Print", btnPrint);
        Authrity.CheckButtonRight("_Print", btnToxls);
        Authrity.CheckButtonRight("_Print", btnToPhone);
        Authrity.CheckButtonRight("_Print", brnToEMail);
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        Util.FillDropDownList(ddlDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "CaseID", ""), "CaseName", "CaseID", false);
        ddlDonor_Type.Items.Insert(0, new ListItem("請選擇", "請選擇"));
        ddlDonor_Type.SelectedIndex = 0;

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

        //經手人
        Util.FillDropDownList(ddlCreate_User, Util.GetDataTable("AdminUser", "1", "1", "", ""), "UserName", "UserName", false);
        ddlCreate_User.Items.Insert(0, new ListItem("", ""));
        ddlCreate_User.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    //dropdownlist選城市自動帶入地區
    protected void ddlCity_SelectedIndexChanged(object sender, EventArgs e)
    {
        Util.FillDropDownList(ddlArea, Util.GetDataTable("CodeCity", "ParentCityID", ddlCity.SelectedValue, "", ""), "Name", "ZipCode", false);
        ddlArea.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea.SelectedIndex = 0;
        tbxZipCode.Text = "";
    }
    //---------------------------------------------------------------------------
    //dropdownlist選地區自動填入郵遞區號
    protected void ddlArea_SelectedIndexChanged(object sender, EventArgs e)
    {
        tbxZipCode.Text = Util.GetCityCode(ddlArea.SelectedItem.Text, ddlCity.SelectedValue).Substring(0, 3);
    }
    //---------------------------------------------------------------------------
    //新增
    public void btnAdd_Click(object sender, EventArgs e)
    {
        // 2014/5/22 修改若輸入「捐款人」查無此捐款人資料，按「新增」按鈕後，應可直接帶入「捐款人」和「收據抬頭」欄位。
        //Response.Redirect(Util.RedirectByTime("DonorInfo_Add.aspx"));
        Response.Redirect(Util.RedirectByTime("DonorInfo_Add.aspx", "donor_name=" + tbxDonor_Name.Text));
    }
    //查詢
    public void btnQuery_Click(object sender, EventArgs e)
    {
        // SQL-> AND串條件
        Query();
        Session["ClickQuery"] = "Query";
    }
    //雷同資料查詢
    protected void btnFuzzyQuery_Click(object sender, EventArgs e)
    {
        // SQL-> Or串條件
        FuzzyQuery();
        Session["ClickQuery"] = "FuzzyQuery";
    }
    public void btnToxls_Click(object sender, EventArgs e)
    {
        Query();
        Response.Redirect("DonorInfo_Print_Excel.aspx");
    }
    public void btnToPhone_Click(object sender, EventArgs e)
    {
        Query();
        Response.Redirect("DonorInfo_Print_Phone.aspx");
    }
    public void btnToEMail_Click(object sender, EventArgs e)
    {
        Query();
        Response.Redirect("DonorInfo_Print_Email.aspx");
    }
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
    //---------------------------------------------------------------------------
    public void Query()
    {
        string strSql;
        DataTable dt;
        strSql = " select Distinct Donor_Id , Donor_Id as 編號, Donor_Name as 捐款人, IDNo as [身分證/統編], (Case When IsMember='Y' Then 'V' Else '' End) as 讀者, Donor_Type as 身分別, Tel_Office as 連絡電話, Cellular_Phone as 手機號碼,\n";
        strSql += " (Case When IsAbroad = 'Y' Then DONOR.OverseasAddress + DONOR.OverseasCountry Else Case When ISNULL(DONOR.City,'')='' Then Address Else Case When ISNULL(A.mValue,'')<>ISNULL(B.mValue,'') Then ISNULL(DONOR.ZipCode,'')+A.mValue+ISNULL(B.mValue,'')+Address End End End) as [通訊地址],\n";
        strSql += " CONVERT(VARCHAR(10) ,Begin_DonateDate, 111) as 首捐日,  CONVERT(VARCHAR(10) ,Last_DonateDate,111) as 末捐日,\n";
        strSql += " Donate_No as 捐款次數, REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Total),1),'.00','') as 累計捐款金額, Create_User as 建檔人員, Donor_Id_Old as 舊編號\n";
        strSql += " From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode\n";
        strSql += " where DeleteDate is null ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        
        if (tbxDonor_Id.Text.Trim() != "")
        {
            strSql += "and Donor_Id=@Donor_Id ";
            dict.Add("Donor_Id", tbxDonor_Id.Text.Trim());
            Session["Donor_Id"] = tbxDonor_Id.Text.Trim();
        }
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140425 修改 by Ian_Kao
            //strSql += "and (Donor_Name Like N'%" + tbxDonor_Name.Text.Trim() + "%' Or NickName  Like N'%" + tbxDonor_Name.Text.Trim() + "%' Or Contactor  Like N'%" + tbxDonor_Name.Text.Trim() + "%' Or Invoice_Title Like N'%" + tbxDonor_Name.Text.Trim() + "%')";
            strSql += "and (Donor_Name Like @Donor_Name Or Contactor  Like @Donor_Name Or Invoice_Title Like @Donor_Name)";
            dict.Add("Donor_Name", "%" + tbxDonor_Name.Text.Trim() + "%");
            Session["Donor_Name"] = "%" + tbxDonor_Name.Text.Trim() + "%";
        }
        //20140623 修改成轉換成數值查詢 by Samson
        if (tbxDonor_Id_Old.Text.Trim() != "")
        {
            strSql += "and cast(Donor_Id_Old as numeric)=cast(@Donor_Id_Old as numeric) ";
            dict.Add("Donor_Id_Old", tbxDonor_Id_Old.Text.Trim());
            Session["Donor_Id_Old"] = tbxDonor_Id_Old.Text.Trim();
        }
        if (tbxTel_Office.Text.Trim() != "")
        {
            //strSql += "and Tel_Office=@Tel_Office ";
            //dict.Add("Tel_Office", tbxTel_Office.Text.Trim());
            //20140807 修正可查詢電話、手機
            strSql += "and (Tel_Office LIKE '%" + tbxTel_Office.Text.Trim() + "%' or Tel_Home LIKE '%" + tbxTel_Office.Text.Trim() + "%' or Cellular_Phone LIKE '%" + tbxTel_Office.Text.Trim() + "%') ";
            Session["Tel_Office"] = tbxTel_Office.Text.Trim();
        }
        if (ddlDonor_Type.SelectedIndex != 0)
        {
            strSql += "and Donor_Type like '%" + ddlDonor_Type.SelectedItem.Text + "%' ";
        }
        //20140127新增
        if (tbxInvoice_Title.Text.Trim() != "")
        {
            //20140425 修改 by Ian_Kao
            //strSql += "and Invoice_Title LIKE N'%" + tbxInvoice_Title.Text.Trim() + "%' ";
            strSql += "and Invoice_Title LIKE @Invoice_Title ";
            dict.Add("Invoice_Title", "%" + tbxInvoice_Title.Text.Trim() + "%");
            Session["Invoice_Title"] = "%" + tbxInvoice_Title.Text.Trim() + "%";
        }
        if (tbxIDNo.Text.Trim() != "")
        {
            strSql += "and IDNo LIKE N'%" + tbxIDNo.Text.Trim() + "%' ";
        }
        if (ddlCreate_User.SelectedIndex != 0)
        {
            strSql += " and DONOR.Create_User = '" + ddlCreate_User.SelectedItem.Text + "'";
        }
        //20140522 新增 by詩儀
        //20140623 修改成模糊查詢 by Samson
        if (tbxEMail.Text.Trim() != "")
        {
            strSql += "and Email LIKE @Email ";
            dict.Add("Email", "%" + tbxEMail.Text.Trim() + "%");
            Session["Email"] = "%" + tbxEMail.Text.Trim() + "%";
        }
        if (tbxIsSendNewsNum.Text.Trim() != "")
        {
            strSql += "and IsSendNewsNum=@IsSendNewsNum ";
            dict.Add("IsSendNewsNum", tbxIsSendNewsNum.Text.Trim());
            Session["IsSendNewsNum"] = tbxIsSendNewsNum.Text.Trim();
        }
        if (cbxIsDVD.Checked)
        {
            strSql += "and IsDVD='Y' ";
        }
        if (cbxIsSendEpaper.Checked)
        {
            strSql += "and IsSendEpaper='Y' ";
        }
        if (cbxIsAbroad.Checked)
        {
            strSql += "and (IsAbroad = 'Y' Or IsAbroad_Invoice = 'Y') ";
            if (tbxIsAbroad_Address.Text.Trim() != "")
            {
                //20140425 修改 by Ian_Kao
                //strSql += "and 通訊地址 LIKE N'%" + tbxIsAbroad_Address.Text.Trim() + "%' ";
                strSql += "and (OverseasAddress LIKE N'%" + tbxIsAbroad_Address.Text.Trim() + "%' or OverseasCountry LIKE N'%" + tbxIsAbroad_Address.Text.Trim() + "%' )";
            }
        }
        if (cbxIsAbroad.Checked == false)
        {
            if (tbxZipCode.Text.Trim() != "")
            {
                strSql += "and B.ZipCode=@ZipCode ";
                dict.Add("ZipCode", tbxZipCode.Text.Trim());
                Session["ZipCode"] = tbxZipCode.Text.Trim();
            }
            if (ddlCity.SelectedIndex != 0)
            {
                strSql += "and City='" + ddlCity.SelectedValue + "' ";
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
            }
            if (ddlSection.SelectedIndex != 0)
            {
                strSql += "and Section='" + ddlSection.SelectedItem.Text + "' ";
            }
            if (tbxLane.Text.Trim() != "")
            {
                strSql += "and Lane=@Lane ";
                dict.Add("Lane", tbxLane.Text.Trim());
                Session["Lane"] = tbxLane.Text.Trim();
            }
            if (tbxAlley.Text.Trim() != "")
            {
                strSql += "and Alley=@Alley ";
                dict.Add("Alley", tbxAlley.Text.Trim());
                Session["Alley"] = tbxAlley.Text.Trim();
            }
            if (tbxNo1.Text.Trim() != "")
            {
                strSql += "and HouseNo=@HouseNo ";
                dict.Add("HouseNo", tbxNo1.Text.Trim());
                Session["HouseNo"] = tbxNo1.Text.Trim();
            }
            if (tbxNo2.Text.Trim() != "")
            {
                strSql += "and HouseNoSub=@HouseNoSub ";
                dict.Add("HouseNoSub", tbxNo2.Text.Trim());
                Session["HouseNoSub"] = tbxNo2.Text.Trim();
            }
            if (tbxFloor1.Text.Trim() != "")
            {
                strSql += "and Floor=@Floor ";
                dict.Add("Floor", tbxFloor1.Text.Trim());
                Session["Floor"] = tbxFloor1.Text.Trim();
            }
            if (tbxFloor2.Text.Trim() != "")
            {
                strSql += "and FloorSub=@FloorSub ";
                dict.Add("FloorSub", tbxFloor2.Text.Trim());
                Session["FloorSub"] = tbxFloor2.Text.Trim();
            }
            if (tbxRoom.Text.Trim() != "")
            {
                strSql += "and Room=@Room ";
                dict.Add("Room", tbxRoom.Text.Trim());
                Session["Room"] = tbxRoom.Text.Trim();
            }
        }
        strSql += "order by Donor_Id desc";
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.Keys.Add("Donor_Id");
            npoGridView.DisableColumn.Add("Donor_Id");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            npoGridView.EditLink = Util.RedirectByTime("DonorInfo_Edit.aspx", "Donor_Id=");
            lblGridList.Text = npoGridView.Render();

            Session["strSql"] = strSql;
        }
        
    }
    //---------------------------------------------------------------------------
    public void FuzzyQuery()
    {
        string strSql;
        DataTable dt;
        strSql = " select Distinct Donor_Id , Donor_Id as 編號, (Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End) as 捐款人, IDNo as [身分證/統編], (Case When IsMember='Y' Then 'V' Else '' End) as 讀者, Donor_Type as 身分別, Tel_Office as 連絡電話, Cellular_Phone as 手機號碼,\n";
        strSql += " (Case When IsAbroad = 'Y' Then DONOR.OverseasAddress + DONOR.OverseasCountry Else Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+Address End End End) as [通訊地址],\n";
        strSql += " CONVERT(VARCHAR(10) ,Begin_DonateDate, 111) as 首捐日,  CONVERT(VARCHAR(10) ,Last_DonateDate,111) as 末捐日,\n";
        strSql += " Donate_No as 捐款次數, REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Total),1),'.00','') as 累計捐款金額, Create_User as 建檔人員, Donor_Id_Old as 舊編號\n";
        strSql += " From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode\n";
        strSql += " where DeleteDate is null ";
        int cnt = 0;
        Dictionary<string, object> dict = new Dictionary<string, object>();
        
        if (tbxDonor_Id.Text.Trim() != "")
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            strSql += "Donor_Id=@Donor_Id ";
            dict.Add("Donor_Id", tbxDonor_Id.Text.Trim());
            Session["Donor_Id"] = tbxDonor_Id.Text.Trim();
        }

        if (tbxDonor_Name.Text.Trim() != "")
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            else
            {
                strSql += " Or ";
            }
            //20140425 修改 by Ian_Kao
            //strSql += "(Donor_Name Like N'%" + tbxDonor_Name.Text.Trim() +"%' Or NickName  Like N'%" + tbxDonor_Name.Text.Trim() +"%' Or Contactor  Like N'%" + tbxDonor_Name.Text.Trim() +"%' Or Invoice_Title Like N'%" + tbxDonor_Name.Text.Trim() +"%')";
            strSql += "(Donor_Name Like @Donor_Name Or NickName  Like @Donor_Name Or Contactor  Like @Donor_Name Or Invoice_Title Like @Donor_Name) ";
            dict.Add("Donor_Name", "%" + tbxDonor_Name.Text.Trim() + "%");
            Session["Donor_Name"] = "%" + tbxDonor_Name.Text.Trim() + "%";
        }

        if (tbxDonor_Id_Old.Text.Trim() != "")
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            else
            {
                strSql += " Or ";
            }
            strSql += "Donor_Id_Old=@Donor_Id_Old ";
            dict.Add("Donor_Id_Old", tbxDonor_Id_Old.Text.Trim());
            Session["Donor_Id_Old"] = tbxDonor_Id_Old.Text.Trim();
        }
        
        if (tbxTel_Office.Text.Trim() != "")
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            else
            {
                strSql += " Or ";
            }
            strSql += "Tel_Office=@Tel_Office ";
            dict.Add("Tel_Office", tbxTel_Office.Text.Trim());
            Session["Tel_Office"] = tbxTel_Office.Text.Trim();
        }
        
        if (ddlDonor_Type.SelectedIndex != 0)
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            else
            {
                strSql += " Or ";
            }
            strSql += "Donor_Type like '%" + ddlDonor_Type.SelectedItem.Text + "%' ";
        }
        
        if (tbxIsSendNewsNum.Text.Trim() != "")
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            else
            {
                strSql += " Or ";
            }
            strSql += "IsSendNewsNum=@IsSendNewsNum ";
            dict.Add("IsSendNewsNum", tbxIsSendNewsNum.Text.Trim());
            Session["IsSendNewsNum"] = tbxIsSendNewsNum.Text.Trim();
        }

        //20140127新增
        if (tbxInvoice_Title.Text.Trim() != "")
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            else
            {
                strSql += " Or ";
            }
            
            //20140425 修改 by Ian_Kao
            //strSql += "Invoice_Title LIKE N'%" + tbxInvoice_Title.Text.Trim() + "%' ";
            strSql += "Invoice_Title LIKE @Invoice_Title ";
            dict.Add("Invoice_Title", "%" + tbxInvoice_Title.Text.Trim() + "%");
            Session["Invoice_Title"] = "%" + tbxInvoice_Title.Text.Trim() + "%";
        }
        if (tbxIDNo.Text.Trim() != "")
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            else
            {
                strSql += " Or ";
            }
            strSql += "IDNo LIKE N'%" + tbxIDNo.Text.Trim() + "%' ";
        }
        if (ddlCreate_User.SelectedIndex != 0)
        {
            strSql += " and DONOR.Create_User = '" + ddlCreate_User.SelectedItem.Text + "'";
        }
        //20140522 新增 by詩儀
        if (tbxEMail.Text.Trim() != "")
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            else
            {
                strSql += " Or ";
            }
            strSql += "Email LIKE @Email ";
            dict.Add("Email", "%" + tbxEMail.Text.Trim() + "%");
            Session["Email"] = "%" + tbxEMail.Text.Trim() + "%";
        }
        if (cbxIsDVD.Checked)
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            else
            {
                strSql += " Or ";
            }
            strSql += "IsDVD='Y' ";
        }
        
        if (cbxIsSendEpaper.Checked)
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            else
            {
                strSql += " Or ";
            }
            strSql += "IsSendEpaper='Y' ";
        }
        
        if (cbxIsAbroad.Checked)
        {
            if (cnt == 0)
            {
                strSql += " And (";
                cnt++;
            }
            else
            {
                strSql += " Or ";
            }
            strSql += "(IsAbroad = 'Y' Or IsAbroad_Invoice = 'Y') ";
            if (tbxIsAbroad_Address.Text.Trim() != "")
            {
                //20140425 修改 by Ian_Kao
                //strSql += "Or Address LIKE N'%" + tbxIsAbroad_Address.Text.Trim() + "%' ";
                strSql += "Or (OverseasAddress LIKE N'%" + tbxIsAbroad_Address.Text.Trim() + "%' or OverseasCountry LIKE N'%" + tbxIsAbroad_Address.Text.Trim() + "%' )";
            }
        }
        if (cbxIsAbroad.Checked == false)
        {
            
            if (tbxZipCode.Text.Trim() != "")
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "B.ZipCode=@ZipCode ";
                dict.Add("ZipCode", tbxZipCode.Text.Trim());
                Session["ZipCode"] = tbxZipCode.Text.Trim();
            }
            
            if (ddlCity.SelectedIndex != 0)
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "City='" + ddlCity.SelectedValue + "' ";
            }
            
            if (ddlArea.SelectedIndex != 0)
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "Area='" + ddlArea.SelectedValue + "' ";
            }
            
            if (tbxStreet.Text.Trim() != "")
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "Street=@Street ";
                dict.Add("Street", tbxStreet.Text.Trim());
                Session["Street"] = tbxStreet.Text.Trim();
            }
            
            if (ddlSection.SelectedIndex != 0)
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "Section='" + ddlSection.SelectedItem.Text + "' ";
            }
            
            if (tbxLane.Text.Trim() != "")
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "Lane=@Lane ";
                dict.Add("Lane", tbxLane.Text.Trim());
                Session["Lane"] = tbxLane.Text.Trim();
            }
            
            if (tbxAlley.Text.Trim() != "")
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "Alley=@Alley ";
                dict.Add("Alley", tbxAlley.Text.Trim());
                Session["Alley"] = tbxAlley.Text.Trim();
            }
            
            if (tbxNo1.Text.Trim() != "")
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "HouseNo=@HouseNo ";
                dict.Add("HouseNo", tbxNo1.Text.Trim());
                Session["HouseNo"] = tbxNo1.Text.Trim();
            }
            
            if (tbxNo2.Text.Trim() != "")
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "HouseNoSub=@HouseNoSub ";
                dict.Add("HouseNoSub", tbxNo2.Text.Trim());
                Session["HouseNoSub"] = tbxNo2.Text.Trim();
            }
            
            if (tbxFloor1.Text.Trim() != "")
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "Floor=@Floor ";
                dict.Add("Floor", tbxFloor1.Text.Trim());
                Session["Floor"] = tbxFloor1.Text.Trim();
            }
            
            if (tbxFloor2.Text.Trim() != "")
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "FloorSub=@FloorSub ";
                dict.Add("FloorSub", tbxFloor2.Text.Trim());
                Session["FloorSub"] = tbxFloor2.Text.Trim();
            }
            
            if (tbxRoom.Text.Trim() != "")
            {
                if (cnt == 0)
                {
                    strSql += " And (";
                    cnt++;
                }
                else
                {
                    strSql += " Or ";
                }
                strSql += "Room=@Room ";
                dict.Add("Room", tbxRoom.Text.Trim());
                Session["Room"] = tbxRoom.Text.Trim();
            }
        }
        if (cnt != 0)
        {
            strSql += " )";
        }
        strSql += "order by Donor_Id desc";
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.Keys.Add("Donor_Id");
            npoGridView.DisableColumn.Add("Donor_Id");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            npoGridView.EditLink = Util.RedirectByTime("DonorInfo_Edit.aspx", "Donor_Id=");
            lblGridList.Text = npoGridView.Render();

            Session["strSql"] = strSql;
        }
    }
}