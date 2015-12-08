using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonateMgr_Invoice_Yearly_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "Invoice_Yearly_Print";
        //權控處理
        AuthrityControl();
        if (!IsPostBack)
        {
            LoadDropDownListData();
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
        Authrity.CheckButtonRight("_Print", btnReport);
        Authrity.CheckButtonRight("_Print", btnPost);
    }
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);

        //捐款年度
        for (int i = 0; i < 16; i++)
        {
            ddlDonate_Date_Year.Items.Insert(i, new ListItem((int.Parse(DateTime.Now.Year.ToString())-i).ToString(), (int.Parse(DateTime.Now.Year.ToString())-i).ToString()));
        }
        //名條格式
        Util.FillDropDownList(ddlFormat, Util.GetDataTable("CaseCode", "GroupName", "郵遞標籤", "", ""), "CaseName", "CaseID", false);
        ddlFormat.Items.Insert(0, new ListItem("", ""));
        ddlFormat.SelectedIndex = 8;
    }
    //---------------------------------------------------------------------------
    protected void btnReport_Click(object sender, EventArgs e)
    {
        Sql_Data();
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Report_OnClick();", true);
    }
    protected void btnAddress_Click(object sender, EventArgs e)
    {
        Sql_Print();
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Address_OnClick();", true);
    }
    protected void btnPost_Click(object sender, EventArgs e)
    {
        Sql_Print();
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Post_OnClick();", true);
    }
    private void Sql_Data()
    {
        /*
        string strSql_temp = @"select distinct dr.Donor_Name as [捐款人], dr.Donor_Id as [捐款人編號], dr.Title as [稱謂], dr.Invoice_Title, IsNull(dr.Title2,'') as Title2, IsNull(dr.ZipCode,'') as [郵遞區號], IsNull(A.mValue,'') as [City], IsNull(B.mValue,'') as [Area],dr.Address, dr.IDNo,dr.Invoice_Attn as [Attn],dr.IsAbroad
                          from Donate do 
                              Left Join Donor dr on do.Donor_Id =  dr.Donor_Id
                              Left Join CODECITY As A On dr.City=A.mCode Left Join CODECITY As B On dr.Area=B.mCode
                              Left Join CODECITY As C On dr.Invoice_City=C.mCode Left Join CODECITY As D On dr.Invoice_Area=D.mCode
                          where dr.DeleteDate is null and Issue_Type != 'D' ";
        string strSql = Condition(strSql_temp);
        */
        // 2014/7/11 增加跨頁(超過20筆跳頁)
        string strSql1_temp = @"select dr.Donor_Name as [捐款人], dr.Donor_Id as [捐款人編號], dr.Title as [稱謂], 
               dr.Invoice_Title, IsNull(dr.Title2,'') as Title2, IsNull(dr.Invoice_ZipCode,'') as [郵遞區號], 
               IsNull(C.mValue,'') as [City], IsNull(D.mValue,'') as [Area],dr.Invoice_Address as [Address], dr.Invoice_IDNo,
               dr.Invoice_Attn as [Attn], dr.IsAbroad, do.cnt as [row_cnt]
               from ( select Donor_Id ,count(*) as cnt from Donate where Issue_Type <> 'D' ";
        string strSql = Condition1(strSql1_temp);
        strSql += @"group by Donor_Id ) as do ";
        string strSql2_temp = @"Left Join Donor dr on do.Donor_Id =  dr.Donor_Id
                              Left Join CODECITY As A On dr.City=A.mCode Left Join CODECITY As B On dr.Area=B.mCode
                              Left Join CODECITY As C On dr.Invoice_City=C.mCode Left Join CODECITY As D On dr.Invoice_Area=D.mCode
                          where dr.DeleteDate is null and ISNULL(IsErrAddress,'') !='Y' ";
        strSql += Condition2(strSql2_temp);

        Session["Condition"] = rblContent.SelectedItem.Text;
        Session["Donate_Date_Year"] = ddlDonate_Date_Year.SelectedItem.Text;
        //20140430 Add by GoodTV Tanya:增加單一捐款人時可「收據抬頭改印」
        Session["Invoice_Title_New"] = tbxInvoice_Title_New.Text;
        Session["strSql_Data"] = strSql;
    }
    private void Sql_Print()
    {
        string strSql_temp = @"select Distinct dr.Donor_Id as [捐款人編號],
                                 dr.Donor_Name as [捐款人],dr.Title as [稱謂],
                                 WhereClause1     
                                 from Donate do 
                                     Left Join Donor dr on do.Donor_Id =  dr.Donor_Id
                                     Left Join CODECITY As A On dr.City=A.mCode Left Join CODECITY As B On dr.Area=B.mCode
                                     Left Join CODECITY As C On dr.Invoice_City=C.mCode Left Join CODECITY As D On dr.Invoice_Area=D.mCode
                                 where dr.DeleteDate is null and Issue_Type != 'D' and ISNULL(IsErrAddress,'') !='Y' ";
        string strSql = Condition(strSql_temp);
        Session["strSql_Print"] = strSql;
    }

    private string Condition(string strSql)
    {
        //機構
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and do.Dept_Id = '" + ddlDept.SelectedValue + "'";
        }

        //捐款年度
        if (ddlDonate_Date_Year.SelectedItem.Text != "")
        {
            strSql += " and Year(Donate_Date) = '" + ddlDonate_Date_Year.SelectedValue + "'";
        }

        //捐款人
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and (dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%' or dr.Invoice_Title=N'"+ tbxDonor_Name.Text.Trim() +"') ";
            strSql += " and (dr.Donor_Name like N'%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%' or dr.Invoice_Title = N'" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "') ";
        }

        //捐款人編號
        if (txtDonor_IdS.Text.Trim() != "")
        {
            strSql += " and do.Donor_Id >='" + txtDonor_IdS.Text.Trim() + "'";
        }
        if (txtDonor_IdE.Text.Trim() != "")
        {
            strSql += " and do.Donor_Id <='" + txtDonor_IdE.Text.Trim() + "'";
        }
        //地址 
        if (rblAddress.SelectedValue == "1")
        {
            strSql += " and  dr.IsAbroad = 'N'";
        }
        else if (rblAddress.SelectedValue == "2")
        {
            strSql += " and  dr.IsAbroad = 'Y'";
        }
        else
        {
            strSql += " and  IsNull(dr.IsAbroad,'') <> ''";
        }
        //名條內容 
        if (rblContent.SelectedValue == "1")
        {
            strSql = strSql.Replace("WhereClause1", "IsNull(dr.Invoice_ZipCode,'') as [郵遞區號],(Case When IsNull(dr.Invoice_City,'')='' Then dr.Invoice_Address Else Case When IsNull(C.mValue,'')<>IsNull(D.mValue,'') Then IsNull(C.mValue,'')+IsNull(D.mValue,'')+dr.Invoice_Address Else IsNull(C.mValue,'')+dr.Invoice_Address End End) as [地址],dr.IsAbroad_Invoice as [IsAbroad],dr.Invoice_Attn as [Attn],dr.Invoice_OverseasCountry as [OverseasCountry],dr.Invoice_OverseasAddress as [OverseasAddress]");
            strSql += " and IsNull(C.mValue,'') + IsNull(D.mValue,'') + dr.Invoice_Address<>''";
        }
        if (rblContent.SelectedValue == "2")
        {
            strSql = strSql.Replace("WhereClause1", "IsNull(dr.Invoice_ZipCode,'') as [郵遞區號],(Case When IsNull(dr.City,'')='' Then dr.Address Else Case When IsNull(A.mValue,'')<>IsNull(B.mValue,'') Then IsNull(A.mValue,'')+IsNull(B.mValue,'')+dr.Address Else IsNull(A.mValue,'')+dr.Address End End) as [地址],dr.IsAbroad as [IsAbroad],dr.Attn as [Attn],dr.OverseasCountry as [OverseasCountry],dr.OverseasAddress as [OverseasAddress]");
            strSql += " and  IsNull(A.mValue,'') + IsNull(B.mValue,'') + dr.Address<>''";
        }

        // 2014/7/11 增加全部收據
        //收據重印與全部收據
        if (!CbxAllInvoice.Checked)
        {
            //收據重印
            if (cbxInvoice.Checked)
            {
                strSql += " and do.Invoice_Print_Yearly = '1' ";
            }
            else
            {
                strSql += " and (do.Invoice_Print_Yearly = '0' or do.Invoice_Print_Yearly is null) ";
            }
        }

        //排序方式
        for (int i = 0; i < rblSort.Items.Count; i++)
        {
            if (rblSort.Items[i].Selected)
            {
                switch (i)
                {
                    case 0: strSql += " order by [郵遞區號]"; break;
                    case 1: strSql += " order by [捐款人編號]"; break;
                }
            }
        }
        return strSql;
    }

    private string Condition1(string strSql)
    {
        //機構
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and Dept_Id = '" + ddlDept.SelectedValue + "'";
        }

        //捐款年度
        if (ddlDonate_Date_Year.SelectedItem.Text != "")
        {
            strSql += " and Year(Donate_Date) = '" + ddlDonate_Date_Year.SelectedValue + "'";
        }

        //捐款人編號
        if (txtDonor_IdS.Text.Trim() != "")
        {
            strSql += " and Donor_Id >='" + txtDonor_IdS.Text.Trim() + "'";
        }
        if (txtDonor_IdE.Text.Trim() != "")
        {
            strSql += " and Donor_Id <='" + txtDonor_IdE.Text.Trim() + "'";
        }

        // 2014/7/11 增加全部收據
        //20150521 增加年度證明列印記錄
        //收據重印與全部收據
        if (!CbxAllInvoice.Checked)
        {
            //收據重印
            if (cbxInvoice.Checked)
            {
                strSql += " and Invoice_Print_Yearly = '1' ";
                Session["RePrint"] = "Y";
            }
            else
            {
                strSql += " and (Invoice_Print_Yearly = '0' or Invoice_Print_Yearly is null) ";
                Session["RePrint"] = "N";
            }
        }
        else
        {
            Session["RePrint"] = "Y";
        }

        return strSql;
    }


    private string Condition2(string strSql)
    {

        //捐款人
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and (dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%' or dr.Invoice_Title=N'"+ tbxDonor_Name.Text.Trim() +"') ";
            strSql += " and (dr.Donor_Name like N'%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%' or dr.Invoice_Title = N'" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "') ";
        }

        //地址 
        if (rblAddress.SelectedValue == "1")
        {
            strSql += " and  dr.IsAbroad = 'N'";
        }
        else if (rblAddress.SelectedValue == "2")
        {
            strSql += " and  dr.IsAbroad = 'Y'";
        }
        else
        {
            strSql += " and  IsNull(dr.IsAbroad,'') <> ''";
        }
        //名條內容 
        if (rblContent.SelectedValue == "1")
        {
            strSql = strSql.Replace("WhereClause1", "IsNull(dr.Invoice_ZipCode,'') as [郵遞區號],(Case When IsNull(dr.Invoice_City,'')='' Then dr.Invoice_Address Else Case When IsNull(C.mValue,'')<>IsNull(D.mValue,'') Then IsNull(C.mValue,'')+IsNull(D.mValue,'')+dr.Invoice_Address Else IsNull(C.mValue,'')+dr.Invoice_Address End End) as [地址],dr.IsAbroad_Invoice as [IsAbroad],dr.Invoice_Attn as [Attn],dr.Invoice_OverseasCountry as [OverseasCountry],dr.Invoice_OverseasAddress as [OverseasAddress]");
            strSql += " and IsNull(C.mValue,'') + IsNull(D.mValue,'') + dr.Invoice_Address<>''";
        }
        if (rblContent.SelectedValue == "2")
        {
            strSql = strSql.Replace("WhereClause1", "IsNull(dr.Invoice_ZipCode,'') as [郵遞區號],(Case When IsNull(dr.City,'')='' Then dr.Address Else Case When IsNull(A.mValue,'')<>IsNull(B.mValue,'') Then IsNull(A.mValue,'')+IsNull(B.mValue,'')+dr.Address Else IsNull(A.mValue,'')+dr.Address End End) as [地址],dr.IsAbroad as [IsAbroad],dr.Attn as [Attn],dr.OverseasCountry as [OverseasCountry],dr.OverseasAddress as [OverseasAddress]");
            strSql += " and  IsNull(A.mValue,'') + IsNull(B.mValue,'') + dr.Address<>''";
        }

        //排序方式
        for (int i = 0; i < rblSort.Items.Count; i++)
        {
            if (rblSort.Items[i].Selected)
            {
                switch (i)
                {
                    case 0: strSql += " order by [郵遞區號]"; break;
                    case 1: strSql += " order by [捐款人編號]"; break;
                }
            }
        }
        return strSql;
    }

}