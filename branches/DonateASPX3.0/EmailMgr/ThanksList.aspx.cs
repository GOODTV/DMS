using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class EmailMgr_ThanksList : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "ThanksList";
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
        Authrity.CheckButtonRight("_Print", btnMailPrint);
        Authrity.CheckButtonRight("_Print", btnAddress);
        Authrity.CheckButtonRight("_Print", btnPreview);
    }
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);

        //募款活動
        Util.FillDropDownList(ddlActName, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        ddlActName.Items.Insert(0, new ListItem("", ""));
        ddlActName.SelectedIndex = 0;

        //感謝函內容
        Util.FillDropDownList(ddlEmail_subject, Util.GetDataTable("EmailMgr", "EmailMgr_Type", "首捐感謝函", "", ""), "EmailMgr_Subject", "EmailMgr_Subject", false);
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
    }
    //---------------------------------------------------------------------------
    protected void btnMailPrint_Click(object sender, EventArgs e)
    {
        string strSql = Sql();
        //感謝函重印
        if (cbxThanks.Checked)
        {
            strSql = strSql.Replace("WhereClause3", " and (dr.IsThanks = '1' or dr.IsThanks = '0' or dr.IsThanks is null) ");
        }
        else
        {
            strSql = strSql.Replace("WhereClause3", " and (dr.IsThanks = '0' or dr.IsThanks is null) ");
        }
        Session["strSql"] = strSql;
        //感謝函內容
        Session["Email_Subject"] = ddlEmail_subject.SelectedItem.Text;
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "MailPrint();", true);
    }
    protected void btnAddress_Click(object sender, EventArgs e)
    {
        string strSql = Sql();
        //地址名條重印
        if (cbxAddress.Checked)
        {
            strSql = strSql.Replace("WhereClause3", " and (dr.IsThanks = '1' or dr.IsThanks = '0' or dr.IsThanks is null) ");
        }
        else
        {
            strSql = strSql.Replace("WhereClause3", " and (dr.IsThanks_Add = '0' or dr.IsThanks_Add is null) ");
        }
        Session["strSql_Print"] = strSql;
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Address();", true);
    }
    protected void btnPreview_Click(object sender, EventArgs e)
    {
    
        //感謝函內容
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
                           where dr.DeleteDate is null WhereClause2 WhereClause3";
        if (rblContent.SelectedValue == "1")
        {
            strSql = strSql.Replace("WhereClause1", " select Distinct dr.Donor_Id as [捐款人編號], dr.ZipCode as [郵遞區號], (Case When dr.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+B.mValue+Address Else A.mValue+Address End End)  as [地址], dr.Donor_Name as [捐款人], case when dr.Title<>'' then dr.title else '君' end as [稱謂]\n");
            strSql = strSql.Replace("WhereClause2", "and  dr.City + dr.Area + dr.Address<>''");
        }
        else if (rblContent.SelectedValue == "2")
        {
            strSql = strSql.Replace("WhereClause1", " select Distinct dr.Donor_Id as [捐款人編號], dr.Invoice_ZipCode as [郵遞區號], (Case When dr.Invoice_City='' Then Invoice_Address Else Case When C.mValue<>D.mValue Then C.mValue+D.mValue+Invoice_Address Else C.mValue+Invoice_Address End End) as [地址], dr.Invoice_Title [捐款人], case when dr.Title2<>'' then dr.title else '君' end as [稱謂]\n");
            strSql = strSql.Replace("WhereClause2", "and dr.Invoice_City+dr.Invoice_Area+dr.Invoice_Address<>''");
        }
        
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and dr.Dept_Id ='" + ddlDept.SelectedValue + "'";
        }
        if (tbxDonor_Name.Text.Trim()!="")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%'";
            strSql += " and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%'";
        }
        //捐款人編號
        if (tbxDonor_IdS.Text.Trim() != "")
        {
            strSql += " and do.Donor_Id >='" + tbxDonor_IdS.Text.Trim() + "'";
        }
        if (tbxDonor_IdE.Text.Trim() != "")
        {
            strSql += " and do.Donor_Id <='" + tbxDonor_IdE.Text.Trim() + "'";
        }
        //捐款日期
        if (tbxDonate_DateS.Text.Trim() != "")
        {
            strSql += " and do.Donate_Date >='" + tbxDonate_DateS.Text.Trim() + "'";
        }
        if (tbxDonate_DateE.Text.Trim() != "")
        {
            strSql += " and Donate_Date <='" + tbxDonate_DateE.Text.Trim() + "'";
        }
        //募款活動
        if (ddlActName.SelectedIndex != 0)
        {
            strSql += " and do.Act_Id ='" + ddlActName.SelectedValue + "' ";
        }
        //排序方式
        if (rblSort.SelectedValue == "1")
        {
            strSql += " order by [郵遞區號]\n";
        }
        else if (rblSort.SelectedValue == "2")
        {
            strSql += " order by [捐款人編號] desc\n";
        }
        return strSql;
    }
}