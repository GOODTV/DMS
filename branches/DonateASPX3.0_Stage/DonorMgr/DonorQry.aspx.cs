using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class DonorMgr_DonorQry : BasePage
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
        LoadFormData();
    }
    protected void btnNextPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.AddStringNumber(HFD_CurrentPage.Value);
        LoadFormData();
    }
    protected void btnGoPage_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    #endregion NpoGridView 處理換頁相關程式碼
    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "DonorQry";
        //權控處理
        AuthrityControl();
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            LoadDropDownListData();
            LoadCheckBoxListData();
        }
    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_Query", btnQuery);
        Authrity.CheckButtonRight("_Print", btnPrint);
        Authrity.CheckButtonRight("_Print", btnToxls);
    }
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.SelectedIndex = 1;
  
        //生日月份
        Util.FillDropDownList(ddlBirthMonth, Util.GetDataTable("CaseCode", "GroupName", "生日月份", "", ""), "CaseName", "CaseID", false);
        ddlBirthMonth.Items.Insert(0, new ListItem("請選擇", "請選擇"));
        ddlBirthMonth.SelectedIndex = 0;

        //通訊地址-縣市
        Util.FillDropDownList(ddlCity, Util.GetDataTable("CodeCity", "ParentCityID", "0", "", ""), "Name", "ZipCode", false);
        ddlCity.Items.Insert(0, new ListItem("縣　市", "縣　市"));
        ddlCity.SelectedIndex = 0;

        //收據地址-縣市
        Util.FillDropDownList(ddlCity2, Util.GetDataTable("CodeCity", "ParentCityID", "0", "", ""), "Name", "ZipCode", false);
        ddlCity2.Items.Insert(0, new ListItem("縣　市", "縣　市"));
        ddlCity2.SelectedIndex = 0;

        //募款活動
        //Util.FillDropDownList(ddlActName, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        //ddlActName.Items.Insert(0, new ListItem("", ""));
        //ddlActName.SelectedIndex = 0;

        //多久未捐款
        Util.FillDropDownList(ddlHowLong, Util.GetDataTable("CaseCode", "GroupName", "多久未捐款", "CaseID", ""), "CaseName", "CaseID", false);
        ddlHowLong.Items.Insert(0, new ListItem("請選擇", "請選擇"));
        ddlHowLong.SelectedIndex = 0;
        
        //捐款人群組
        string strSql = @"
                          select uid, GroupClassName
                          from GroupClass gc
                          where DeleteDate is null
                         ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        Util.FillDropDownList(ddlGroupClass, dt, "GroupClassName", "uid", true);
    }
    public void LoadCheckBoxListData()
    {
        //身份別
        Util.FillCheckBoxList(cblDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "CaseID", ""), "CaseName", "CaseID", false);
        cblDonor_Type.Items[0].Selected = false;
        //文宣品
        cblPropaganda.Items.Add("紙本月刊");
        cblPropaganda.Items.Add("DVD");
        cblPropaganda.Items.Add("電子文宣");
        cblPropaganda.Items.Add("大額謝卡");
        cblPropaganda.Items.Add("郵簡");
    }
    //---------------------------------------------------------------------------
    //dropdownlist選縣市自動帶入鄉鎮市區
    protected void ddlCity_SelectedIndexChanged(object sender, EventArgs e)
    {
        Util.FillDropDownList(ddlArea, Util.GetDataTable("CodeCity", "ParentCityID", ddlCity.SelectedValue, "", ""), "Name", "ZipCode", false);
        ddlArea.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea.SelectedIndex = 0;
    }
    protected void ddlCity2_SelectedIndexChanged(object sender, EventArgs e)
    {
        Util.FillDropDownList(ddlArea2, Util.GetDataTable("CodeCity", "ParentCityID", ddlCity2.SelectedValue, "", ""), "Name", "ZipCode", false);
        ddlArea2.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea2.SelectedIndex = 0;
    }
    //---------------------------------------------------------------------------
    //查詢
    public void LoadFormData()
    { 
        string strSql;
        DataTable dt;
        strSql = @"select
                        Distinct DONOR.Donor_Id as 捐款人編號, 
	                    Donor_Name as 捐款人,Title as 稱謂,
	                    Sex as 性別,
	                    Donor_Type as 身分別,
	                    Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
				                            (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
				                            Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
		                                Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
				                            Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End as 電話日,
	                    Cellular_Phone as 手機,
	                    Email as 電子信箱,
                        Case When DONOR.IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'')  
	                              else (Case when ISNULL(Attn,'') <> '' then DONOR.ZipCode + IsNull(B.Name,'') + IsNull(C.Name,'') + Address+'('+ISNULL(Attn,'')+')' 
				                        Else DONOR.ZipCode + IsNull(B.Name,'') + IsNull(C.Name,'') + Address End) END as 通訊地址,
	                    CONVERT(varchar(12) , [Last_DonateDate], 111 ) as 最近捐款日,
	                    --Donate_No as 累積次數,
	                    --REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Total),1),'.00','') as 累積金額
                        REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,do.Sum_Amt),1),'.00','') AS '累積金額', 
						do.Times AS '累積次數',IsSendNewsNum AS '月刊本數'
	                From DONOR 
                    INNER JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
							            FROM DONATE 
                                        WHERE ISNULL(Issue_Type,'') != 'D' ";

        if (txtDonateDateS.Text.Trim() != "")
        {
            //strSql += @" and Last_DonateDate >='" + txtDonateDateS.Text.Trim() + "'";
            //20150828欄位搜尋改成末捐日期改成捐款日期
            strSql += @" and Donate_Date >='" + txtDonateDateS.Text.Trim() + "'";
        }
        if (txtDonateDateE.Text.Trim() != "")
        {
            strSql += @" and Donate_Date <='" + txtDonateDateE.Text.Trim() + "'";
        }
        //捐款單筆金額 20150826新增
        if (txtDonate_AmtS.Text.Trim() != "")
        {
            strSql += " and Donate_Amt > '" + txtDonate_AmtS.Text.Trim() + "'";
        }
        if (txtDonate_AmtE.Text.Trim() != "")
        {
            strSql += " and Donate_Amt < '" + txtDonate_AmtS.Text.Trim() + "'";
        }
        if (tbxDonate_Total_Amt.Text.Trim() != "")
        {
            strSql += " GROUP BY Donor_Id Having SUM(Donate_Amt) > " + tbxDonate_Total_Amt.Text.Trim() + " ) do ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id ) do ";
        }
        strSql += @"On Donor.Donor_Id=do.Donor_Id 
                    WhereClause2
                    Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode
                    Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode 
                    where DONOR.DeleteDate is null  ";

        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and Dept_Id='" + ddlDept.SelectedValue + "' ";
        }
        //身分別-----------------------------------------------------------------------------------//
        bool first = true; int cnt = 0;
        for (int i = 0; i < cblDonor_Type.Items.Count; i++)
        {
            if (cblDonor_Type.Items[i].Selected && first)
            {
                strSql += " and (Donor_Type like '%" + cblDonor_Type.Items[i].ToString() + "%' "; first = false; cnt++;
            }
            else if (cblDonor_Type.Items[i].Selected)
            {
                strSql += " or Donor_Type like '%" + cblDonor_Type.Items[i].ToString() + "%' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //----------------------------------------------------------------------------------------//
        if (ddlBirthMonth.SelectedIndex != 0)
        {
            strSql += " and cast(MONTH([Birthday])as nvarchar)+'月'='" + ddlBirthMonth.SelectedItem.Text + "'";
        }
        if (ddlCity.SelectedIndex != 0)
        {
            strSql += " and City= '" + ddlCity.SelectedValue + "'";
        }
        if (ddlArea.SelectedIndex != 0 && ddlArea.Items.Count != 0)
        {
            strSql += " and Area like '%" + ddlArea.SelectedValue + "'";
        }
        if (ddlCity2.SelectedIndex != 0)
        {
            strSql += " and Invoice_City= '" + ddlCity2.SelectedValue + "'";
        }
        if (ddlArea2.SelectedIndex != 0 && ddlArea2.Items.Count != 0)
        {
            strSql += " and Invoice_Area like '%" + ddlArea2.SelectedValue + "'";
        }
        //if (ddlActName.SelectedIndex != 0)
        //{
        //    strSql = strSql.Replace("WhereClause1", " Join (Select Donor_Id,Act_Id From Donate Where Dept_Id = '" + ddlDept.SelectedValue + "' And Issue_Type<>'D') As A On Donor.Donor_Id=A.Donor_Id join Act T on A.Act_Id=T.Act_Id ");
        //    strSql += "and T.Act_ShortName='" + ddlActName.SelectedItem.Text + "'";
        //}
        //else
        //{
        //    strSql = strSql.Replace("WhereClause1", "");
        //}
        //if (txtDonor_Name.Text.Trim() != "")
        //{
        //    //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
        //    //strSql += " and Donor_Name like '%" + txtDonor_Name.Text.Trim() + "%'"; 
        //    strSql += " and Donor_Name like '%" + txtDonor_Name.Text.Trim().Replace("'", "''") + "%'";
        //}
        
        if (txtDonate_TotalS.Text.Trim() != "")
        {
            strSql += @" and cast(Donate_Total as int) >='" + txtDonate_TotalS.Text.Trim() + "'";
        }
        if (txtDonate_TotalE.Text.Trim() != "")
        {
            strSql += @" and cast(Donate_Total as int) <='" + txtDonate_TotalE.Text.Trim() + "'";
        }
        if (ddlHowLong.SelectedIndex != 0)
        {
            strSql += @" and --case when DATEDIFF(m,[Last_DonateDate],GETDATE())<=5 then '3個月' else(
							    --case when DATEDIFF(m,[Last_DonateDate],GETDATE())<9 and DATEDIFF(m,[Last_DonateDate],GETDATE())>=6 then '6個月' else(
								    case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<3 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=1 then '一年以上三年以下' else(
									    case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<5 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=3 then '三年以上五年以下' else(
										    case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<7 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=5 then '五年以上七年以下' else(
                                                case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<10 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=7 then '七年以上十年以下' else( 
											        case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<15 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=10 then '十年以上十五年以下' else( 
											            case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>15 then '十五年以上' end
                                                    )end	
                                                )end	
										    )end
									    )end
								    --)end
							    --)end
	                        )end ='" + ddlHowLong.SelectedItem.Text + "'";
        }
        if (txtDonor_IdS.Text.Trim() != "")
        {
            strSql += @" and DONOR.Donor_Id >='" + txtDonor_IdS.Text.Trim() + "'";
        }
        if (txtDonor_IdE.Text.Trim() != "")
        {
            strSql += @" and DONOR.Donor_Id <='" + txtDonor_IdE.Text.Trim() + "'";
        }
        if (txtDonate_NoS.Text.Trim() != "")
        {
            strSql += @" and Donate_No >='" + txtDonate_NoS.Text.Trim() + "'";
        }
        if (txtDonate_NoE.Text.Trim() != "")
        {
            strSql += @" and Donate_No <='" + txtDonate_NoE.Text.Trim() + "'";
        }
        //末捐日期 20150916新增
        if (txtLastDonateDateS.Text.Trim() != "")
        {
            strSql += @" and Last_DonateDate >='" + txtLastDonateDateS.Text.Trim() + "'";
        }
        if (txtLastDonateDateE.Text.Trim() != "")
        {
            strSql += @" and Last_DonateDate <='" + txtLastDonateDateE.Text.Trim() + "'";
        }
        //文宣品-----------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblPropaganda.Items.Count; i++)
        {
            if (cblPropaganda.Items[i].Selected && first)
            {
                strSql += " and ("; first = false;
                switch (i)
                {
                    case 0: strSql += " IsSendNews='Y'"; break;
                    case 1: strSql += " IsDVD='Y' "; break;
                    case 2: strSql += " IsSendEpaper='Y' "; break;
                    case 3: strSql += " IsBigAmtThank='Y' "; break;
                    case 4: strSql += " IsPost='Y' "; break;
                }
                cnt++;
            }
            else if (cblPropaganda.Items[i].Selected)
            {
                strSql += " and";
                switch (i)
                {
                    case 0: strSql += " IsSendNews='Y' "; break;
                    case 1: strSql += " IsDVD='Y' "; break;
                    case 2: strSql += " IsSendEpaper='Y' "; break;
                    case 3: strSql += " IsBigAmtThank='Y' "; break;
                    case 4: strSql += " IsPost='Y' "; break;
                }
                cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //----------------------------------------------------------------------------------------//
        //20150826增加不含錯址、歿、不主動聯絡
        if (cbxIsErrAddress.Checked)
        {
            strSql += " and (IsErrAddress='N' or IsErrAddress is NULL)";
        }
        if (cbxNoDie.Checked)
        {
            strSql += " and IsNull(Sex,'') <> '歿'";
        }
        if (cbxIsContact.Checked)
        {
            strSql += " and IsContact='N'";
        }
        //群組類別 20150826新增
        if (ddlGroupClass.SelectedIndex != 0)
        {
            strSql = strSql.Replace("WhereClause2", @" Join (Select DonorId,GroupItemUid From GroupMapping Where DeleteDate is NULL) As GM On Donor.Donor_Id=GM.DonorId 
                                                       join GroupItem GI on GM.GroupItemUid=GI.Uid  
                                                       join GroupClass GC on GI.GroupClassUid=GC.Uid ");
            strSql += "and GC.GroupClassName='" + ddlGroupClass.SelectedItem.Text + "'";
            if (txtGroupItemName.Text.Trim() != "")
            {
                strSql += "and GI.GroupItemName='" + txtGroupItemName.Text.Trim() + "'";
            }
        }
        else
        {
            strSql = strSql.Replace("WhereClause2", "");
        }
        //區域 20150828新增
        if (rblAddress.SelectedValue == "1")
        {
            strSql += " and  IsAbroad = 'N'";
        }
        else if (rblAddress.SelectedValue == "2")
        {
            strSql += " and  IsAbroad = 'Y'";
        }
        else
        {
            
        }
        // 2014/4/8 增加捐款人編號的排序
        //排序-----------------------------------------------------------------------------------//
        strSql += " order by DONOR.Donor_Id";

        Session["strSql"] = strSql;
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
            lblAmt.Text = "0";
            return;
        }
        //Grid initial
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        // 2014/4/8 顯示捐款人編號
        //npoGridView.DisableColumn.Add("捐款人編號");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
        npoGridView.Keys.Add("捐款人編號");
        npoGridView.EditLink = Util.RedirectByTime("DonorInfo_Edit.aspx", "Donor_Id=");
        lblGridList.Text = npoGridView.Render();
        lblAmt.Text = Donate_Amt();
    }
    private string Donate_Amt()
    {
        string strSql = @"select CONVERT(VARCHAR,count(do.Donor_Id)),REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(do.Sum_Amt)),1),'.00','') as Donate_Amt 
                            From DONOR 
                            INNER JOIN (SELECT Donor_Id ,SUM(Donate_Amt) AS Sum_Amt ,COUNT(Donor_Id) AS Times  
							                    FROM DONATE 
                                                WHERE ISNULL(Issue_Type,'') != 'D'";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        if (txtDonateDateS.Text.Trim() != "")
        {
            //strSql += @" and Last_DonateDate >='" + txtDonateDateS.Text.Trim() + "'";
            //20150828欄位搜尋改成末捐日期改成捐款日期
            strSql += @" and Donate_Date >='" + txtDonateDateS.Text.Trim() + "'";
        }
        if (txtDonateDateE.Text.Trim() != "")
        {
            strSql += @" and Donate_Date <='" + txtDonateDateE.Text.Trim() + "'";
        }
        //捐款單筆金額 20150826新增
        if (txtDonate_AmtS.Text.Trim() != "")
        {
            strSql += " and Donate_Amt > '" + txtDonate_AmtS.Text.Trim() + "'";
        }
        if (txtDonate_AmtE.Text.Trim() != "")
        {
            strSql += " and Donate_Amt < '" + txtDonate_AmtS.Text.Trim() + "'";
        }
        if (tbxDonate_Total_Amt.Text.Trim() != "")
        {
            strSql += " GROUP BY Donor_Id Having SUM(Donate_Amt) > " + tbxDonate_Total_Amt.Text.Trim() + " ) do ";
        }
        else
        {
            strSql += " GROUP BY Donor_Id ) do ";
        }
        strSql += @"On Donor.Donor_Id=do.Donor_Id 
                    WhereClause2
                    Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode
                    Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode 
                    where DONOR.DeleteDate is null  ";

        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and Dept_Id='" + ddlDept.SelectedValue + "' ";
        }
        //身分別-----------------------------------------------------------------------------------//
        bool first = true; int cnt = 0;
        for (int i = 0; i < cblDonor_Type.Items.Count; i++)
        {
            if (cblDonor_Type.Items[i].Selected && first)
            {
                strSql += " and (Donor_Type like '%" + cblDonor_Type.Items[i].ToString() + "%' "; first = false; cnt++;
            }
            else if (cblDonor_Type.Items[i].Selected)
            {
                strSql += " or Donor_Type like '%" + cblDonor_Type.Items[i].ToString() + "%' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //----------------------------------------------------------------------------------------//
        if (ddlBirthMonth.SelectedIndex != 0)
        {
            strSql += " and cast(MONTH([Birthday])as nvarchar)+'月'='" + ddlBirthMonth.SelectedItem.Text + "'";
        }
        if (ddlCity.SelectedIndex != 0)
        {
            strSql += " and City= '" + ddlCity.SelectedValue + "'";
        }
        if (ddlArea.SelectedIndex != 0 && ddlArea.Items.Count != 0)
        {
            strSql += " and Area like '%" + ddlArea.SelectedValue + "'";
        }
        if (ddlCity2.SelectedIndex != 0)
        {
            strSql += " and Invoice_City= '" + ddlCity2.SelectedValue + "'";
        }
        if (ddlArea2.SelectedIndex != 0 && ddlArea2.Items.Count != 0)
        {
            strSql += " and Invoice_Area like '%" + ddlArea2.SelectedValue + "'";
        }
        //if (ddlActName.SelectedIndex != 0)
        //{
        //    strSql = strSql.Replace("WhereClause1", " Join (Select Donor_Id,Act_Id From Donate Where Dept_Id = '" + ddlDept.SelectedValue + "' And Issue_Type<>'D') As A On Donor.Donor_Id=A.Donor_Id join Act T on A.Act_Id=T.Act_Id ");
        //    strSql += "and T.Act_ShortName='" + ddlActName.SelectedItem.Text + "'";
        //}
        //else
        //{
        //    strSql = strSql.Replace("WhereClause1", "");
        //}
        //if (txtDonor_Name.Text.Trim() != "")
        //{
        //    //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
        //    //strSql += " and Donor_Name like '%" + txtDonor_Name.Text.Trim() + "%'"; 
        //    strSql += " and Donor_Name like '%" + txtDonor_Name.Text.Trim().Replace("'", "''") + "%'";
        //}

        if (txtDonate_TotalS.Text.Trim() != "")
        {
            strSql += @" and cast(Donate_Total as int) >='" + txtDonate_TotalS.Text.Trim() + "'";
        }
        if (txtDonate_TotalE.Text.Trim() != "")
        {
            strSql += @" and cast(Donate_Total as int) <='" + txtDonate_TotalE.Text.Trim() + "'";
        }
        if (ddlHowLong.SelectedIndex != 0)
        {
            strSql += @" and --case when DATEDIFF(m,[Last_DonateDate],GETDATE())<=5 then '3個月' else(
							    --case when DATEDIFF(m,[Last_DonateDate],GETDATE())<9 and DATEDIFF(m,[Last_DonateDate],GETDATE())>=6 then '6個月' else(
								    case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<3 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=1 then '一年' else(
									    case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<5 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=3 then '三年' else(
										    case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<7 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=5 then '五年' else(
                                                case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<10 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=7 then '七年' else( 
											        case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<15 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=10 then '十年' else( 
											            case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>15 then '十五年以上' end
                                                    )end	
                                                )end	
										    )end
									    )end
								    --)end
							    --)end
	                        )end ='" + ddlHowLong.SelectedItem.Text + "'";
        }
        if (txtDonor_IdS.Text.Trim() != "")
        {
            strSql += @" and DONOR.Donor_Id >='" + txtDonor_IdS.Text.Trim() + "'";
        }
        if (txtDonor_IdE.Text.Trim() != "")
        {
            strSql += @" and DONOR.Donor_Id <='" + txtDonor_IdE.Text.Trim() + "'";
        }
        if (txtDonate_NoS.Text.Trim() != "")
        {
            strSql += @" and Donate_No >='" + txtDonate_NoS.Text.Trim() + "'";
        }
        if (txtDonate_NoE.Text.Trim() != "")
        {
            strSql += @" and Donate_No <='" + txtDonate_NoE.Text.Trim() + "'";
        }

        //文宣品-----------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblPropaganda.Items.Count; i++)
        {
            if (cblPropaganda.Items[i].Selected && first)
            {
                strSql += " and ("; first = false;
                switch (i)
                {
                    case 0: strSql += " (IsSendNewsNum<>'' or IsSendNewsNum<>'0')"; break;
                    case 1: strSql += " IsDVD='Y' "; break;
                    case 2: strSql += " IsSendEpaper='Y' "; break;
                    case 3: strSql += " IsBigAmtThank='Y' "; break;
                    case 4: strSql += " IsPost='Y' "; break;
                }
                cnt++;
            }
            else if (cblPropaganda.Items[i].Selected)
            {
                strSql += " and";
                switch (i)
                {
                    case 0: strSql += " (IsSendNewsNum<>'' or IsSendNewsNum<>'0') "; break;
                    case 1: strSql += " IsDVD='Y' "; break;
                    case 2: strSql += " IsSendEpaper='Y' "; break;
                    case 3: strSql += " IsBigAmtThank='Y' "; break;
                    case 4: strSql += " IsPost='Y' "; break;
                }
                cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //----------------------------------------------------------------------------------------//
        //20150826增加不含錯址、歿、不主動聯絡
        if (cbxIsErrAddress.Checked)
        {
            strSql += " and (IsErrAddress='N' or IsErrAddress is NULL)";
        }
        if (cbxNoDie.Checked)
        {
            strSql += " and IsNull(Sex,'') <> '歿'";
        }
        if (cbxIsContact.Checked)
        {
            strSql += " and IsContact='N'";
        }
        //群組類別 20150826新增
        if (ddlGroupClass.SelectedIndex != 0)
        {
            strSql = strSql.Replace("WhereClause2", @" Join (Select DonorId,GroupItemUid From GroupMapping Where DeleteDate is NULL) As GM On Donor.Donor_Id=GM.DonorId 
                                                       join GroupItem GI on GM.GroupItemUid=GI.Uid  
                                                       join GroupClass GC on GI.GroupClassUid=GC.Uid ");
            strSql += "and GC.GroupClassName='" + ddlGroupClass.SelectedItem.Text + "'";
            if (txtGroupItemName.Text.Trim() != "")
            {
                strSql += "and GI.GroupItemName='" + txtGroupItemName.Text.Trim() + "'";
            }
        }
        else
        {
            strSql = strSql.Replace("WhereClause2", "");
        }
        //區域 20150828新增
        if (rblAddress.SelectedValue == "1")
        {
            strSql += " and  IsAbroad = 'N'";
        }
        else if (rblAddress.SelectedValue == "2")
        {
            strSql += " and  IsAbroad = 'Y'";
        }
        else
        {
            
        }

        DataTable dt;
        //20140425 Modify by GoodTV Tanya
        //dt = NpoDB.QueryGetScalar(strSql, dict);
        dt = NpoDB.GetDataTableS(strSql, dict);
        DataRow dr = dt.Rows[0];
        string Donate_Amt = dr["Donate_Amt"].ToString();
        return Donate_Amt;
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        HFD_Query_Flag.Value = "Y";
        LoadFormData();
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Response.Redirect("DonorQry_Print_Excel.aspx");
    }
}