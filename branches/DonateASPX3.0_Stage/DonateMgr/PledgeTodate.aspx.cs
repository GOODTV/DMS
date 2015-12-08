using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonateMgr_PledgeTodate : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "PledgeTodate";
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
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.Items.Insert(0, new ListItem("", ""));
        ddlDept.SelectedIndex = 1;

        //捐款授權到期日
        Util.FillDropDownList(ddlYear_Donate_ToDate, Int32.Parse(DateTime.Now.Year.ToString()), Int32.Parse(DateTime.Now.Year.ToString()) + 10, true, 0);
        ddlYear_Donate_ToDate.SelectedIndex = 1;
        Util.FillDropDownList(ddlMonth_Donate_ToDate, 1, 12, true, 0);
        ddlMonth_Donate_ToDate.Text = (DateTime.Now.Month + 1).ToString();

        //授權狀態
        ddlStatus.Items.Insert(0, new ListItem("", ""));
        ddlStatus.Items.Insert(1, new ListItem("授權中", "授權中"));
        ddlStatus.Items.Insert(2, new ListItem("停止", "停止"));
        ddlStatus.SelectedIndex = 1;

        //授權方式
        ddlDonate_Type.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Type.Items.Insert(1, new ListItem("長期捐款", "長期捐款"));
        ddlDonate_Type.Items.Insert(2, new ListItem("單次捐款", "單次捐款"));
        ddlDonate_Type.SelectedIndex = 1;

        //到期信函內容
        Util.FillDropDownList(ddlContent, Util.GetDataTable("EmailMgr", "EmailMgr_Type", "固定捐款到期通知", "", ""), "EmailMgr_Subject", "EmailMgr_Subject", false);
        ddlContent.Items.Insert(0, new ListItem("", ""));
        if (ddlContent.Items.Count == 2)
        {
            ddlContent.SelectedIndex = 1;
        }
        else
        {
            ddlContent.SelectedIndex = 0;
        }

        //名條格式
        Util.FillDropDownList(ddlFormat, Util.GetDataTable("CaseCode", "GroupName", "郵遞標籤", "", ""), "CaseName", "CaseID", false);
        ddlFormat.Items.Insert(0, new ListItem("", ""));
        ddlFormat.SelectedIndex = 8;

    }
    //---------------------------------------------------------------------------
    protected void btnMailPrint_Click(object sender, EventArgs e)
    {
        Session["strSql"] = Sql();
        Session["Email_Subject"] = ddlContent.SelectedItem.Text;
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "MailPrint();", true);
    }
    protected void btnAddress_Click(object sender, EventArgs e)
    {
        Session["strSql_Print"] = Sql();
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Address();", true);
    }
    protected void btnPreview_Click(object sender, EventArgs e)
    {
        Session["Email_Subject"] = ddlContent.SelectedItem.Text;
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Preview();", true);
    }
    //---------------------------------------------------------------------------
    private string Sql()
    {
        string strSql = @"WhereClause1
                          from Pledge P Left Join Donor D on P.Donor_Id = D.Donor_Id 
                          Left Join CODECITY As A On D.City=A.mCode Left Join CODECITY As B On D.Area=B.mCode
                          Left Join CODECITY As C On D.Invoice_City=C.mCode Left Join CODECITY As E On D.Invoice_Area=E.mCode
                          where D.DeleteDate is null WhereClause2";
        if (rblContent.SelectedValue == "1")
        {
            strSql = strSql.Replace("WhereClause1", " select  Pledge_Id as [Pledge_Id], P.Donor_Id as [捐款人編號], D.ZipCode as [郵遞區號], (Case When D.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+B.mValue+Address Else A.mValue+Address End End) as [地址], D.Donor_Name as [捐款人], case when D.Title<>'' then D.title else '君' end as [稱謂], Year(P.Donate_ToDate) as Year, Month(P.Donate_ToDate) as Month\n");
            strSql = strSql.Replace("WhereClause2", " and  D.City + D.Area + D.Address<>'' ");
        }
        else if (rblContent.SelectedValue =="2")
        {
            strSql = strSql.Replace("WhereClause1", " select  Pledge_Id as [Pledge_Id], P.Donor_Id as [捐款人編號], D.Invoice_ZipCode as [郵遞區號], (Case When D.Invoice_City='' Then Invoice_Address Else Case When C.mValue<>E.mValue Then C.mValue+E.mValue+Invoice_Address Else C.mValue+Invoice_Address End End) as [地址], D.Invoice_Title [捐款人], case when D.Title2<>'' then D.title2 else '君' end as [稱謂], Year(P.Donate_ToDate) as Year, Month(P.Donate_ToDate) as Month\n");
            strSql = strSql.Replace("WhereClause2", " and D.Invoice_City+D.Invoice_Area+D.Invoice_Address<>'' ");
        }
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and P.Dept_Id ='" + ddlDept.SelectedValue + "'";            
        }
        if (ddlYear_Donate_ToDate.SelectedIndex != 0)
        {
            strSql += " and Year(P.Donate_ToDate)='" + ddlYear_Donate_ToDate.SelectedItem.Text + "'";
        }
        if (ddlMonth_Donate_ToDate.SelectedIndex != 0)
        {
            strSql += " and Month(P.Donate_ToDate)='" + ddlMonth_Donate_ToDate.SelectedItem.Text + "'";
        }
        if (ddlStatus.SelectedIndex != 0)
        {
            strSql += " and P.Status ='" + ddlStatus.SelectedItem.Text + "'";
        }
        if (ddlDonate_Type.SelectedIndex != 0)
        {
            strSql += " and Donate_Type ='" + ddlDonate_Type.SelectedItem.Text + "'";
        }
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and D.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%'";
            strSql += " and D.Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%'";
        }
        if (tbxPledge_Id_From.Text.Trim() != "" && tbxPledge_Id_To.Text.Trim() != "")
        {
            strSql += " and P.Pledge_Id between '" + tbxPledge_Id_From.Text.Trim() + "' and '" + tbxPledge_Id_To.Text.Trim() + "' ";
        }
        if (rblSort.SelectedValue == "1")
        {
            strSql += "order by [郵遞區號]\n";
        }
        else if (rblSort.SelectedValue == "2")
        {
            strSql += "order by [捐款人編號] desc\n";
        }
        else if (rblSort.SelectedValue == "3")
        {
            strSql += "order by [Pledge_Id] desc\n";
        }
        return strSql;
    }
}