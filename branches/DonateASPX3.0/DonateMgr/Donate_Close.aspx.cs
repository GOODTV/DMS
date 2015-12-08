using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonateMgr_Donate_Close : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "Donate_Close";
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            LoadFormData();
        }
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = "select DeptID as Dept_Id, DeptShortName as [機構簡稱], CONVERT(VARCHAR(10) , Donate_Close, 111 ) as 捐款關帳日, CONVERT(VARCHAR(10) , Contribute_Close, 111 ) as 物品捐贈關帳日 \n";
        strSql += " from Dept where 1=1 ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count != 0)
        {
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.Keys.Add("Dept_Id");
            npoGridView.DisableColumn.Add("Dept_Id");
            npoGridView.ShowPage = false;
            npoGridView.EditLink = Util.RedirectByTime("Donate_Close_Edit.aspx", "Dept_Id=");
            lblGridList.Text = npoGridView.Render();
        }
        else
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
        }
    }
}