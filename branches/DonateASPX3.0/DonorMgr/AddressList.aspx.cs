using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonorMgr_AddressList : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "AddressList";
        //權控處理
        AuthrityControl();
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
        Authrity.CheckButtonRight("_Print", btnAddress);
        Authrity.CheckButtonRight("_Print", btnPost);
    }
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.Items.Insert(0, new ListItem("", ""));
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
        Util.FillDropDownList(ddlActName, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        ddlActName.Items.Insert(0, new ListItem("", ""));
        ddlActName.SelectedIndex = 0;

        //多久未捐款
        Util.FillDropDownList(ddlHowLong, Util.GetDataTable("CaseCode", "GroupName", "多久未捐款", "", ""), "CaseName", "CaseID", false);
        ddlHowLong.Items.Insert(0, new ListItem("", ""));
        ddlHowLong.SelectedIndex = 0;

        //20140425 修改 by Ian_Kao
        //生日卡
        Util.FillDropDownList(ddlBirthMonth2, Util.GetDataTable("CaseCode", "GroupName", "生日月份", "", ""), "CaseName", "CaseID", false);
        ddlBirthMonth2.Items.Insert(0, new ListItem("請選擇", "請選擇"));
        ddlBirthMonth2.SelectedValue = DateTime.Now.AddMonths(1).Month.ToString();

        //名條格式
        Util.FillDropDownList(ddlFormat, Util.GetDataTable("CaseCode", "GroupName", "郵遞標籤", "", ""), "CaseName", "CaseID", false);
        ddlFormat.Items.Insert(0, new ListItem("", ""));
        ddlFormat.SelectedIndex = 8;
    }
    public void LoadCheckBoxListData()
    {
        //身份別
        Util.FillCheckBoxList(cblDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "CaseID", ""), "CaseName", "CaseID", false);
        cblDonor_Type.Items[0].Selected = false;
        //捐款方式
        Util.FillCheckBoxList(cblDonate_Payment, Util.GetDataTable("CaseCode", "GroupName", "捐款方式", "ABS(CaseID)", ""), "CaseName", "CaseID", false);
        cblDonate_Payment.Items[0].Selected = false;
        //捐款用途
        Util.FillCheckBoxList(cblDonate_Purpose, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseID", false);
        cblDonate_Purpose.Items[0].Selected = false;
        //捐款類別
        cblDonate_Type.Items.Add("單次捐款");
        cblDonate_Type.Items.Add("定期捐款");
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
    protected void btnAddress_Click(object sender, EventArgs e)
    {
        //20140425 修改 by Ian_Kao 驗證是否有選擇月份
        if (rbPrint_Type_2.Checked && ddlBirthMonth2.SelectedIndex == 0)
        {
            ShowSysMsg("請選擇生日卡月份");
            return;
        }
        Sql();
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Address_OnClick(" + ddlFormat.SelectedIndex + ");", true);
    }
    protected void btnPost_Click(object sender, EventArgs e)
    {
        //20140425 修改 by Ian_Kao 驗證是否有選擇月份
        if (rbPrint_Type_2.Checked && ddlBirthMonth2.SelectedIndex == 0)
        {
            ShowSysMsg("請選擇生日卡月份");
            return;
        }
        Sql();
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Post_OnClick();", true);
    }
    private void Sql()
    {
        string strSql;
        //20140425 修改 by Ian_Kao
        string strSql2 = "";//儲存篩選條件的地方，因為人數過多所以先在裡層篩選Donor人數
        //DataTable dt;
        strSql = @"select
                        Distinct d.Donor_Id as [編號], 
	                    Donor_Name as [捐款人],
	                    Title as [稱謂],
                        WhereClause1
                    WhereClause2
                    WhereClause3
                    Left Join CODECITY As A On d.City=A.mCode Left Join CODECITY As B On d.Area=B.mCode
	                join Donate Do on d.Donor_id = Do.Donor_Id
                    WhereClause4
                    where DeleteDate is null  ";

        if (ddlDept.SelectedIndex != 0)
        {
            strSql = strSql.Replace("WhereClause3", " join Dept dn on d.Dept_Id=dn.DeptId");
            strSql += " and dn.DeptId='" + ddlDept.SelectedValue + "' ";
        }
        else
        {
            strSql = strSql.Replace("WhereClause3", "");
        }
        //----------------------------------------------------------------------------------------//
        if (txtDonor_Name.Text.Trim() != "")
        {
            //20140425 修改 by Ian_Kao
            //strSql += "and Donor_Name like '%" + txtDonor_Name.Text.Trim() + "%' ";
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql2 += "and Donor_Name like '%" + txtDonor_Name.Text.Trim() + "%' ";
            strSql2 += "and Donor_Name like '%" + txtDonor_Name.Text.Trim().Replace("'", "''") + "%' ";
        }
        //身分別-----------------------------------------------------------------------------------//
        bool first = true; int cnt = 0;
        for (int i = 0; i < cblDonor_Type.Items.Count; i++)
        {
            if (cblDonor_Type.Items[i].Selected && first)
            {
                //20140425 修改 by Ian_Kao
                //strSql += "and (Donor_Type='" + cblDonor_Type.Items[i].ToString() + "' "; first = false; cnt++;
                strSql2 += "and (Donor_Type like '%" + cblDonor_Type.Items[i].ToString() + "%' "; first = false; cnt++;
            }
            else if (cblDonor_Type.Items[i].Selected)
            {
                //20140425 修改 by Ian_Kao
                //strSql += "or Donor_Type='" + cblDonor_Type.Items[i].ToString() + "' "; cnt++;
                strSql2 += "or Donor_Type like '%" + cblDonor_Type.Items[i].ToString() + "%' "; cnt++;
            }
        }
        if (cnt != 0)
        { strSql2 += ")\n"; };
        //----------------------------------------------------------------------------------------//
        if (ddlBirthMonth.SelectedIndex != 0)
        {
            //20140425 修改 by Ian_Kao
            //strSql += "and cast(MONTH([Birthday])as nvarchar)+'月'='" + ddlBirthMonth.SelectedItem.Text + "'";
            strSql2 += "and cast(MONTH([Birthday])as nvarchar)+'月'='" + ddlBirthMonth.SelectedItem.Text + "'";
        }
        if (ddlCity.SelectedIndex != 0)
        {
            //20140425 修改 by Ian_Kao
            //strSql += "and City= '" + ddlCity.SelectedValue + "'";
            strSql2 += "and City= '" + ddlCity.SelectedValue + "'";
        }
        if (ddlArea.SelectedIndex != 0 && ddlArea.Items.Count != 0)
        {
            //20140425 修改 by Ian_Kao
            //strSql += "and Area like '%" + ddlArea.SelectedValue + "'";
            strSql2 += "and Area like '%" + ddlArea.SelectedValue + "'";
        }
        if (ddlCity2.SelectedIndex != 0)
        {
            //20140425 修改 by Ian_Kao
            //strSql += "and invoice_City= '" + ddlCity2.SelectedValue + "'";
            strSql2 += "and invoice_City= '" + ddlCity2.SelectedValue + "'";
        }
        if (ddlArea2.SelectedIndex != 0 && ddlArea2.Items.Count != 0)
        {
            //20140425 修改 by Ian_Kao
            //strSql += "and invoice_Area like '%" + ddlArea2.SelectedValue + "'";
            strSql2 += "and invoice_Area like '%" + ddlArea2.SelectedValue + "'";
        }
        if (ddlActName.SelectedIndex != 0)
        {
            strSql = strSql.Replace("WhereClause4", " join Act T on Do.Act_Id=T.Act_Id ");
            strSql += "and T.Act_ShortName='" + ddlActName.SelectedItem.Text + "' ";
        }
        else
        {
            strSql = strSql.Replace("WhereClause4", "");
        }
        if (txtDonor_Name.Text.Trim() != "")
        {
            //20140425 修改 by Ian_Kao
            //strSql += "and Donor_Name like '%" + txtDonor_Name.Text.Trim() + "%' ";
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql2 += "and Donor_Name like '%" + txtDonor_Name.Text.Trim() + "%' ";
            strSql2 += "and Donor_Name like '%" + txtDonor_Name.Text.Trim().Replace("'", "''") + "%' ";
        }
        if (txtLast_DonateDateS.Text.Trim() != "")
        {
            //20140425 修改 by Ian_Kao
            //strSql += @"and Last_DonateDate >='" + txtLast_DonateDateS.Text.Trim() + "'";
            strSql2 += @"and Last_DonateDate >='" + txtLast_DonateDateS.Text.Trim() + "'";
        }
        if (txtLast_DonateDateE.Text.Trim() != "")
        {
            //20140425 修改 by Ian_Kao
            //strSql += @"and Last_DonateDate <='" + txtLast_DonateDateE.Text.Trim() + "'";
            strSql2 += @"and Last_DonateDate <='" + txtLast_DonateDateE.Text.Trim() + "'";
        }
        if (txtDonate_TotalS.Text.Trim() != "")
        {
            //20140425 修改 by Ian_Kao
            //strSql += @"and cast(Donate_Total as int) >='" + txtDonate_TotalS.Text.Trim() + "'";
            strSql2 += @"and cast(Donate_Total as int) >='" + txtDonate_TotalS.Text.Trim() + "'";
        }
        if (txtDonate_TotalE.Text.Trim() != "")
        {
            //20140425 修改 by Ian_Kao
            //strSql += @"and cast(Donate_Total as int) <='" + txtDonate_TotalE.Text.Trim() + "'";
            strSql2 += @"and cast(Donate_Total as int) <='" + txtDonate_TotalE.Text.Trim() + "'";
        }
        if (ddlHowLong.SelectedIndex != 0)
        {
            //20140425 修改 by Ian_Kao
//            strSql += @"and case when DATEDIFF(m,[Last_DonateDate],GETDATE())<=5 then '3個月' else(
//							    case when DATEDIFF(m,[Last_DonateDate],GETDATE())<9 and DATEDIFF(m,[Last_DonateDate],GETDATE())>=6 then '6個月' else(
//								    case when DATEDIFF(m,[Last_DonateDate],GETDATE())<12 and DATEDIFF(m,[Last_DonateDate],GETDATE())>=9 then '9個月' else(
//									    case when (DATEDIFF(m,[Last_DonateDate],GETDATE())/12)=1 then '一年' else(
//										    case when (DATEDIFF(m,[Last_DonateDate],GETDATE())/12)=2 then '二年' else(
//											    case when (DATEDIFF(m,[Last_DonateDate],GETDATE())/12)>3 then '三年以上' end							
//										    )end
//									    )end
//								    )end
//							    )end
//	                        )end ='" + ddlHowLong.SelectedItem.Text + "'";
            strSql2 += @"and case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<3 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=1 then '一年以上三年以下' else(
								case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<5 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=3 then '三年以上五年以下' else(
									case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<7 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=5 then '五年以上七年以下' else(
                                        case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<10 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=7 then '七年以上十年以下' else( 
											case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<15 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=10 then '十年以上十五年以下' else( 
											    case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>15 then '十五年以上' end
                                            )end	
                                        )end	
									)end
								)end
	                        )end ='" + ddlHowLong.SelectedItem.Text + "'";
        }
        if (txtDonor_IdS.Text.Trim() != "")
        {
            strSql += @"and d.Donor_Id >='" + txtDonor_IdS.Text.Trim() + "'";
        }
        if (txtDonor_IdE.Text.Trim() != "")
        {
            strSql += @"and d.Donor_Id <='" + txtDonor_IdE.Text.Trim() + "'";
        }
        if (txtDonate_NoS.Text.Trim() != "")
        {
            strSql += @"and Donate_No >='" + txtDonate_NoS.Text.Trim() + "'";
        }
        if (txtDonate_NoE.Text.Trim() != "")
        {
            strSql += @"and Donate_No <='" + txtDonate_NoE.Text.Trim() + "'";
        }
        if (tbxInvoice_NoS.Text.Trim() != "")
        {
            strSql += @"and do.Invoice_No >='" + tbxInvoice_NoS.Text.Trim() + "'";
        }
        if (tbxInvoice_NoE.Text.Trim() != "")
        {
            strSql += @"and do.Invoice_No <='" + tbxInvoice_NoE.Text.Trim() + "'";
        }
        if (cblDonate_Type.Items[0].Selected && cblDonate_Type.Items[1].Selected)
        {
            strSql += @"and (do.Donate_Type ='單次捐款' or do.Donate_Type ='長期捐款')";
        }
        if (cblDonate_Type.Items[0].Selected && cblDonate_Type.Items[1].Selected == false)
        {
            strSql += @"and do.Donate_Type ='單次捐款'";
        }
        if (cblDonate_Type.Items[0].Selected == false && cblDonate_Type.Items[1].Selected)
        {
            strSql += @"and do.Donate_Type ='長期捐款'";
        }
        //捐款方式-----------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblDonate_Payment.Items.Count; i++)
        {
            if (cblDonate_Payment.Items[i].Selected && first)
            {
                if (i != cblDonate_Payment.Items.Count - 1)
                {
                    strSql += " and (do.Donate_Payment='" + cblDonate_Payment.Items[i].ToString() + "' "; first = false; cnt++;
                }
                else
                {
                    strSql += @" and (do.Donate_Payment<>'現金' and do.Donate_Payment<>'劃撥' and do.Donate_Payment<>'網路信用卡' and do.Donate_Payment<>'支票' and 
                                 do.Donate_Payment<>'ACH轉帳授權書' and do.Donate_Payment<>'匯款' and do.Donate_Payment<>'ATM' and  
                                 do.Donate_Payment<>'信用卡授權書(聯信)' and do.Donate_Payment<>'信用卡授權書(一般)' and do.Donate_Payment<>'郵局帳戶轉帳授權書' "; first = false; cnt++;
                }
            }
            else if (cblDonate_Payment.Items[i].Selected)
            {
                if (i != cblDonate_Payment.Items.Count - 1)
                {
                    strSql += " or do.Donate_Payment='" + cblDonate_Payment.Items[i].ToString() + "' "; cnt++;
                }
                else
                {
                    strSql += @" or (do.Donate_Payment<>'現金' and do.Donate_Payment<>'劃撥' and do.Donate_Payment<>'網路信用卡' and do.Donate_Payment<>'支票' and 
                                 do.Donate_Payment<>'ACH轉帳授權書' and do.Donate_Payment<>'匯款' and do.Donate_Payment<>'ATM' and 
                                 do.Donate_Payment<>'信用卡授權書(聯信)' and do.Donate_Payment<>'信用卡授權書(一般)' and do.Donate_Payment<>'郵局帳戶轉帳授權書' ) "; first = false; cnt++;
                }
            }
        }
        if (cnt != 0)
        { strSql += ")\n"; }
        //捐款用途-----------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblDonate_Purpose.Items.Count; i++)
        {
            if (cblDonate_Purpose.Items[i].Selected && first)
            {
                first = false;
                strSql += "and (";
                strSql += " do.Donate_Purpose like '%" + cblDonate_Purpose.Items[i].Text + "%'";
                cnt++;
            }
            else if (cblDonate_Purpose.Items[i].Selected)
            {
                strSql += "or";
                strSql += " do.Donate_Purpose like '%" + cblDonate_Purpose.Items[i].Text + "%'";
                cnt++;
            }
        }
        if (cnt != 0)
        { strSql += ")\n"; }
        //名條用途
        if (rbPrint_Type_1.Checked)
        {
            //20140425 修改 by Ian_Kao
            //strSql += " and (IsSendNewsNum<>'' or IsSendNewsNum<>'0')";
            strSql2 += " and (IsSendNewsNum<>'' or IsSendNewsNum<>'0')";
        }
        //20140425 修改 by Ian_Kao
        if (rbPrint_Type_2.Checked)
        {
            if (ddlBirthMonth2.SelectedIndex != 0)
            {
                strSql2 += "and cast(MONTH([Birthday])as nvarchar)+'月'='" + ddlBirthMonth2.SelectedItem.Text + "'";
            }
            //strSql += " and month(Birthday) = '" + tbxBirth_Month.Text.Trim() + "'";
        }
        
        if (rbPrint_Type_3.Checked)
        {
            //20140425 修改 by Ian_Kao
            //strSql += " and IsSendEpaper='Y' ";
            strSql2 += " and IsSendEpaper='Y' ";
        }
        //名條內容
        if (rblContent.SelectedValue == "1")
        {
            strSql = strSql.Replace("WhereClause1", " d.ZipCode as [郵遞區號], (Case When d.City='' Then d.Address Else Case When A.mValue<>B.mValue Then A.mValue+B.mValue+d.Address Else A.mValue+d.Address End End) as [地址]");
            strSql += " and City+Area+d.Address <> '' ";
        }
        if (rblContent.SelectedValue == "2")
        {
            strSql = strSql.Replace("WhereClause1", " d.Invoice_ZipCode as [郵遞區號], (Case When d.Invoice_City='' Then d.Invoice_Address Else Case When A.mValue<>B.mValue Then A.mValue+B.mValue+d.Invoice_Address Else A.mValue+d.Invoice_Address End End) as [地址]");
            strSql += " and Invoice_City+Invoice_Area+Invoice_Address <> '' ";
        }
        //排序方式
        int k = int.Parse(rblSort.SelectedValue);
        {
            switch (k)
            {
                //20140425 修改 by Ian_Kao
                case 1:
                    if (rblContent.SelectedValue == "1")
                    {
                        strSql = strSql.Replace("WhereClause2", " from (select top 10000 * from DONOR Where 1=1 " + strSql2 + " order by ZipCode)d ");
                        strSql += " order by d.ZipCode ";
                    }
                    else if (rblContent.SelectedValue == "2")
                    {
                        strSql = strSql.Replace("WhereClause2", " from (select top 10000 * from DONOR Where 1=1 " + strSql2 + " order by Invoice_ZipCode)d ");
                        strSql += " order by d.Invoice_ZipCode ";
                    }
                    break;
                case 2:
                    strSql = strSql.Replace("WhereClause2", " from DONOR d ");
                    strSql += strSql2;
                    strSql += " order by d.Donor_Id";
                    break;
            }
        }
        Session["strSql"] = strSql;
    }
}