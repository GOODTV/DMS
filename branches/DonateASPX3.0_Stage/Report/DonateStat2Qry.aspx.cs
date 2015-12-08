using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Report_DonateStat2Qry : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "DonateStat2Qry";
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

        //統計項目X
        string[] CategoryX = { "類別", "性別", "年齡", "教育程度", "職業別", "婚姻狀況", "宗教信仰", "通訊縣市", "收據縣市", "捐款方式", "捐款用途", "捐款金額" };
        for (int i = 0; i < CategoryX.Length; i++)
        {
            ddlCategoryX.Items.Insert(i, new ListItem(CategoryX[i], CategoryX[i]));
        }
        ddlCategoryX.SelectedIndex = 0;
        //統計項目Y
        string[] CategoryY = { "類別", "性別", "年齡", "教育程度", "職業別", "婚姻狀況", "宗教信仰", "通訊縣市", "收據縣市", "捐款方式", "捐款用途", "捐款金額" };
        for (int i = 0; i < CategoryY.Length; i++)
        {
            ddlCategoryY.Items.Insert(i, new ListItem(CategoryY[i], CategoryY[i]));
        }
        ddlCategoryY.SelectedIndex = 1;
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