using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonateMgr_VerifyList : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "VerifyList";
        //權控處理
        AuthrityControl();
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
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
        Authrity.CheckButtonRight("_Print", btnPrint);
        Authrity.CheckButtonRight("_Print", btnToxls);
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.Items.Insert(0, new ListItem("", ""));
        ddlDept.SelectedIndex = 1;

        //募款活動
        Util.FillDropDownList(ddlAct_Id, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        ddlAct_Id.Items.Insert(0, new ListItem("", ""));
        ddlAct_Id.SelectedIndex = 0;

    }
    //---------------------------------------------------------------------------
    protected void btnPrint_Click(object sender, EventArgs e)
    {
        Condition();
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Print();", true);
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Condition();
        Response.Redirect(Util.RedirectByTime("../DonateMgr/VerifyList_Print_Excel_style"+rblDonate_Purpose_Type.SelectedValue+".aspx"));
    }
    private void Condition()
    {
        string strSql="";
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and do.Dept_Id = '" + ddlDept.SelectedValue + "' ";
        }
        Session["Donor_Name"] = "";
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%' ";
            //Session["Donor_Name"] = " and Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%' ";
            strSql += " and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%' ";
            Session["Donor_Name"] = " and Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%' ";
        }
        if (tbxDonateDateS.Text.Trim() != "" && tbxDonateDateE.Text.Trim() == "")
        {
            strSql += " and do.Donate_Date between '" + tbxDonateDateS.Text.Trim() + "' and getdate() ";
        }
        else if (tbxDonateDateS.Text.Trim() == "" && tbxDonateDateE.Text.Trim() != "")
        {
            strSql += " and (do.Donate_Date between '1900/01/01' and'" + tbxDonateDateE.Text.Trim() + "') ";
        }
        else if (tbxDonateDateS.Text.Trim() != "" && tbxDonateDateE.Text.Trim() != "")
        {
            strSql += " and do.Donate_Date between '" + tbxDonateDateS.Text.Trim() + "' and '" + tbxDonateDateE.Text.Trim() + "' ";
            Session["Donate_Date"] = " and Donate.Donate_Date between '" + tbxDonateDateS.Text.Trim() + "' and '" + tbxDonateDateE.Text.Trim() + "' ";
        }
        if (ddlAct_Id.SelectedIndex != 0)
        {
            strSql += " and do.Act_Id = '" + ddlAct_Id.SelectedValue + "' ";
        }
        Session["condition"] = strSql;
    }
}