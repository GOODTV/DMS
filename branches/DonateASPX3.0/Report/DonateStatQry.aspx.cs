using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Report_DonateStatQry : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "DonateStatQry";
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
        Authrity.CheckButtonRight("_Print", btnPrint);
        Authrity.CheckButtonRight("_Print", btnToxls);
    }
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.Items.Insert(0, new ListItem("", ""));
        ddlDept.SelectedIndex = 1;

        //募款活動
        Util.FillDropDownList(ddlActName, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        ddlActName.Items.Insert(0, new ListItem("", ""));
        ddlActName.SelectedIndex = 0;
    }
    protected void btnPrint_Click(object sender, EventArgs e)
    {
        string strSql = Sql();
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        string strSql = Sql();
    }
    private string Sql()
    {
        string strSql = "";

        return strSql;
    }
}