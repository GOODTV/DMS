using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class EmailMgr_OtherList : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "OtherList";
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
        Authrity.CheckButtonRight("_Print", btnMailPrint);
        Authrity.CheckButtonRight("_Print", btnAddress);
        Authrity.CheckButtonRight("_Print", btnPreview);
    }
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.Items.Insert(0, new ListItem("", ""));
        ddlDept.SelectedIndex = 1;

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

        //生日月份
        Util.FillDropDownList(ddlBirthMonth, Util.GetDataTable("CaseCode", "GroupName", "生日月份", "", ""), "CaseName", "CaseID", false);
        ddlBirthMonth.Items.Insert(0, new ListItem("", ""));
        ddlBirthMonth.SelectedIndex = 0;

        //多久未捐款
        Util.FillDropDownList(ddlHowLong, Util.GetDataTable("CaseCode", "GroupName", "多久未捐款", "", ""), "CaseName", "CaseID", false);
        ddlHowLong.Items.Insert(0, new ListItem("", ""));
        ddlHowLong.SelectedIndex = 0;

        //其他信函內容
        Util.FillDropDownList(ddlEmail_subject, Util.GetDataTable("EmailMgr", "EmailMgr_Type", "其他信函", "", ""), "EmailMgr_Subject", "EmailMgr_Subject", false);
        ddlEmail_subject.Items.Insert(0, new ListItem("", ""));
        if (ddlEmail_subject.Items.Count == 2)
        {
            ddlEmail_subject.SelectedIndex = 1;
        }
        else
        {
            ddlEmail_subject.SelectedIndex = 0;
        }


        //名條格式
        Util.FillDropDownList(ddlFormat, Util.GetDataTable("CaseCode", "GroupName", "郵遞標籤", "", ""), "CaseName", "CaseID", false);
        ddlFormat.Items.Insert(0, new ListItem("", ""));
        ddlFormat.SelectedIndex = 8;

        ddlArea.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
        ddlArea2.Items.Insert(0, new ListItem("鄉鎮市區", "鄉鎮市區"));
    }
    public void LoadCheckBoxListData()
    {
        //類別
        //Util.FillCheckBoxList(cblDonor_Category, Util.GetDataTable("CaseCode", "GroupName", "捐款類別", "", ""), "CaseName", "CaseID", false);
        //cblDonor_Category.Items[0].Selected = false;
        //身份別
        Util.FillCheckBoxList(cblDonor_Type, Util.GetDataTable("CaseCode", "GroupName", "身份別", "CaseID", ""), "CaseName", "CaseID", false);
        cblDonor_Type.Items[0].Selected = false;
        //文宣品
        Util.FillCheckBoxList(cblPropaganda, Util.GetDataTable("CaseCode", "GroupName", "文宣品", "", ""), "CaseName", "CaseID", false);
        cblPropaganda.Items[0].Selected = false;
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
    protected void btnMailPrint_Click(object sender, EventArgs e)
    {
        Session["strSql"] = Sql();
        //其他信函內容
        Session["Email_Subject"] = ddlEmail_subject.SelectedItem.Text;
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "MailPrint();", true);
    }
    protected void btnAddress_Click(object sender, EventArgs e)
    {
        Session["strSql_Print"] = Sql();
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Address();", true);
    }
    protected void btnPreview_Click(object sender, EventArgs e)
    {
        //其他信函內容
        Session["Email_Subject"] = ddlEmail_subject.SelectedItem.Text;
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Preview();", true);
    }
    //---------------------------------------------------------------------------
    private string Sql()
    {
        string strSql = @" WhereClause1
                           from Donate do Left Join Donor dr on do.Donor_Id = dr.Donor_Id
                           Left Join CODECITY As A On dr.City=A.mCode Left Join CODECITY As B On dr.Area=B.mCode
                           Left Join CODECITY As C On dr.Invoice_City=C.mCode Left Join CODECITY As D On dr.Invoice_Area=D.mCode
                           where dr.DeleteDate is null WhereClause2 ";
        if (rblContent.SelectedValue == "1")
        {
            strSql = strSql.Replace("WhereClause1", " select Distinct dr.Donor_Id as [捐款人編號], dr.ZipCode as [郵遞區號], (Case When dr.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+B.mValue+Address Else A.mValue+Address End End) as [地址], dr.Donor_Name as [捐款人], case when dr.Title<>'' then dr.title else '君' end as [稱謂] ");
            strSql = strSql.Replace("WhereClause2", "and  dr.City + dr.Area + dr.Address<>''");
        }
        else if (rblContent.SelectedValue == "2")
        {
            strSql = strSql.Replace("WhereClause1", " select Distinct dr.Donor_Id as [捐款人編號], dr.Invoice_ZipCode as [郵遞區號], (Case When dr.Invoice_City='' Then Invoice_Address Else Case When C.mValue<>D.mValue Then C.mValue+D.mValue+Invoice_Address Else C.mValue+Invoice_Address End End) as [地址], dr.Invoice_Title [捐款人], case when dr.Title2<>'' then dr.title2 else '君' end as [稱謂] ");
            strSql = strSql.Replace("WhereClause2", "and dr.Invoice_City+dr.Invoice_Area+dr.Invoice_Address<>''");
        }
        
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += @" and dr.Dept_Id ='" + ddlDept.SelectedValue + "'";
        }
        ////類別-----------------------------------------------------------------------------------//
        //bool first = true; int cnt = 0;
        //for (int i = 0; i < cblDonor_Category.Items.Count; i++)
        //{
        //    if (cblDonor_Category.Items[i].Selected && first)
        //    {
        //        strSql += @"and (dr.Category='" + cblDonor_Category.Items[i].ToString() + "' "; first = false; cnt++;
        //    }
        //    else if (cblDonor_Category.Items[i].Selected)
        //    {
        //        strSql += @"or dr.Category='" + cblDonor_Category.Items[i].ToString() + "' "; cnt++;
        //    }
        //}
        //if (cnt != 0) strSql += ')';
        //身分別-----------------------------------------------------------------------------------//
        bool first = true; int cnt = 0;
        for (int i = 0; i < cblDonor_Type.Items.Count; i++)
        {
            if (cblDonor_Type.Items[i].Selected && first)
            {
                strSql += @"and (dr.Donor_Type='" + cblDonor_Type.Items[i].ToString() + "' "; first = false; cnt++;
            }
            else if (cblDonor_Type.Items[i].Selected)
            {
                strSql += @"or dr.Donor_Type='" + cblDonor_Type.Items[i].ToString() + "' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //----------------------------------------------------------------------------------------//
        if (ddlCity.SelectedIndex != 0)
        {
            strSql += @"and dr.City= '" + ddlCity.SelectedValue + "'";
        }
        if (ddlArea.SelectedIndex != 0 && ddlArea.Items.Count != 0)
        {
            strSql += @"and dr.Area like '%" + ddlArea.SelectedValue + "'";
        }
        if (ddlCity2.SelectedIndex != 0)
        {
            strSql += @"and dr.invoice_City= '" + ddlCity2.SelectedValue + "'";
        }
        if (ddlArea2.SelectedIndex != 0 && ddlArea2.Items.Count != 0)
        {
            strSql += @"and dr.invoice_Area like '%" + ddlArea2.SelectedValue + "'";
        }
        //捐款日期
        if (tbxDonate_DateS.Text.Trim() != "")
        {
            strSql += @" and do.Donate_Date >='" + tbxDonate_DateS.Text.Trim() + "'";
        }
        if (tbxDonate_DateE.Text.Trim() != "")
        {
            strSql += @" and Donate_Date <='" + tbxDonate_DateE.Text.Trim() + "'";
        }
        //捐款總金額
        if (tbxDonate_TotalS.Text.Trim() != "")
        {
            strSql += @"and cast(Donate_Total as int) >='" + tbxDonate_TotalS.Text.Trim() + "'";
        }
        if (tbxDonate_TotalE.Text.Trim() != "")
        {
            strSql += @"and cast(Donate_Total as int) <='" + tbxDonate_TotalE.Text.Trim() + "'";
        }
        //募款活動
        if (ddlActName.SelectedIndex != 0)
        {
            strSql += @" and do.Act_Id ='" + ddlActName.SelectedValue + "' ";
        }
        //捐款總次數
        if (tbxDonate_NoS.Text.Trim() != "")
        {
            strSql += @"and dr.Donate_No >='" + tbxDonate_NoS.Text.Trim() + "'";
        }
        if (tbxDonate_NoE.Text.Trim() != "")
        {
            strSql += @"and dr.Donate_No <='" + tbxDonate_NoE.Text.Trim() + "'";
        }
        //捐款人
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += @" and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%'";
            strSql += @" and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%'";
        }
        //捐款人編號
        if (tbxDonor_IdS.Text.Trim() != "")
        {
            strSql += @" and do.Donor_Id >='" + tbxDonor_IdS.Text.Trim() + "'";
        }
        if (tbxDonor_IdE.Text.Trim() != "")
        {
            strSql += @" and do.Donor_Id <='" + tbxDonor_IdE.Text.Trim() + "'";
        }
        if (ddlBirthMonth.SelectedIndex != 0)
        {
            strSql += "and cast(MONTH([Birthday])as nvarchar)+'月'='" + ddlBirthMonth.SelectedItem.Text + "'";
        }
        if (ddlHowLong.SelectedIndex != 0)
        {
            strSql += @"and case when DATEDIFF(YYYY,[Last_DonateDate],GETDATE())<3 and DATEDIFF(YYYY,[Last_DonateDate],GETDATE())>=1 then '一年以上三年以下' else(
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
        //文宣品-----------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblPropaganda.Items.Count; i++)
        {
            if (cblPropaganda.Items[i].Selected && first)
            {
                strSql += "and ("; first = false;
                //strSql += cblPropaganda.Items[i].Value + "='Y' ";
                switch (cblPropaganda.Items[i].Text)
                {
                    case "紙本月刊": strSql += " IsSendNewsNum='Y' "; break;
                    case "DVD": strSql += " IsDVD='Y' "; break;
                    case "電子文宣": strSql += " IsSendEpaper='Y' "; break;
                    case "生日卡": strSql += " IsBirthday='Y' "; break;
                }
                cnt++;
            }
            else if (cblPropaganda.Items[i].Selected)
            {
                strSql += "or";
                switch (cblPropaganda.Items[i].Text)
                {
                    case "紙本月刊": strSql += " IsSendNewsNum='Y' "; break;
                    case "DVD": strSql += " IsDVD='Y' "; break;
                    case "電子文宣": strSql += " IsSendEpaper='Y' "; break;
                    case "生日卡": strSql += " IsBirthday='Y' "; break;
                }
                cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //----------------------------------------------------------------------------------------//
        if (cbxIsAbroad.Checked)
        {
            strSql += "or IsAbroad='Y'";
        }
        //排序方式
        if (rblSort.SelectedValue == "1")
        {
            strSql += @"order by [郵遞區號]";
        }
        else if (rblSort.SelectedValue == "2")
        {
            strSql += @"order by [捐款人編號] desc";
        }
        return strSql;
    }
}