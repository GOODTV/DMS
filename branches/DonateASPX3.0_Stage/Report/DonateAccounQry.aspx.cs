using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Report_DonateAccounReport : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "DonateAccounQry";

        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            LoadDropDownListData();
            LoadCheckBoxListData();
            DateTime dt = DateTime.Now;
            DateTime startMonth = dt.AddDays(1 - dt.Day);
            txtDonateDateS.Text = startMonth.ToShortDateString();
            txtDonateDateE.Text = startMonth.AddMonths(1).AddDays(-1).ToShortDateString();
        }
    }
    //----------------------------------------------------------------------
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
    public void LoadCheckBoxListData()
    {
        //捐款方式
        Util.FillCheckBoxList(cblDonate_Payment, Util.GetDataTable("CaseCode", "GroupName", "捐款方式", "ABS(CaseID)", ""), "CaseName", "CaseName", false);
        cblDonate_Payment.Items[0].Selected = false;

        //捐款用途
        Util.FillCheckBoxList(cblDonate_Purpose, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseName", false);
        cblDonate_Purpose.Items[0].Selected = false;
    }
    //----------------------------------------------------------------------
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Sql();
        Response.Redirect("DonateAccounQry_Print_Excel.aspx");
    }
    protected void btnPrint_Click(object sender, EventArgs e)
    {
        Sql();
        ////check-----------------------------------------------------------------------------------//
        //DataTable dt;
        //Dictionary<string, object> dict = new Dictionary<string, object>();
        //dt = NpoDB.GetDataTableS(Session["strSql"].ToString(), dict);
        //NPOGridView npoGridView = new NPOGridView();
        //npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        //npoGridView.dataTable = dt;
        //lblGridList.Text = npoGridView.Render();
        ////----------------------------------------------------------------------------------------//
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Print('');", true);
    }
    private void Sql()
    {
        string strSql;
        strSql = @"select '一般捐款' as [會計科目]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '一般捐款' then Donate_Amt else 0 end)),1),'.00','') as [捐款金額]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '一般捐款' then Donate_Fee else 0 end)),1),'.00','') as [手續費]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '一般捐款' then Donate_Accou else 0 end)),1),'.00','') as [實收金額]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '一般捐款' then 1 else 0 end)),1),'.00','') as [捐款筆數]
                           ,'活動捐款' as [會計科目]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '活動捐款' then Donate_Amt else 0 end)),1),'.00','') as [捐款金額]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '活動捐款' then Donate_Fee else 0 end)),1),'.00','') as [手續費]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '活動捐款' then Donate_Accou else 0 end)),1),'.00','') as [實收金額]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '活動捐款' then 1 else 0 end)),1),'.00','') as [捐款筆數]
                           ,'急難救助' as [會計科目]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '急難救助' then Donate_Amt else 0 end)),1),'.00','') as [捐款金額]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '急難救助' then Donate_Fee else 0 end)),1),'.00','') as [手續費]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '急難救助' then Donate_Accou else 0 end)),1),'.00','') as [實收金額]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '急難救助' then 1 else 0 end)),1),'.00','') as [捐款筆數]
                           ,'科目不詳' as [會計科目]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '' or Accounting_Title is null then Donate_Amt else 0 end)),1),'.00','') as [捐款金額]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '' or Accounting_Title is null then Donate_Fee else 0 end)),1),'.00','') as [手續費]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '' or Accounting_Title is null then Donate_Accou else 0 end)),1),'.00','') as [實收金額]
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Accounting_Title = '' or Accounting_Title is null then 1 else 0 end)),1),'.00','') as [捐款筆數]
                           ,CONVERT(VARCHAR, GETDATE(), 120 ) as 列印日期
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(Donate_Amt)),1),'.00','') as 總捐款金額
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(Donate_Fee)),1),'.00','') as 總手續費
                           ,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(Donate_Accou)),1),'.00','') as 總實收金額
                           ,count(Accounting_Title) as 總筆數
                    from Donate D left join Donor Dr on D.Donor_Id = Dr.Donor_Id
                    where Dr.DeleteDate is null ";
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and Dr.Dept_Id='" + ddlDept.SelectedValue + "' ";
        }
        if (txtDonateDateS.Text.Trim() != "")
        {
            strSql += " and D.Donate_Date >= '" + txtDonateDateS.Text.Trim() + "' ";
        }
        if (txtDonateDateE.Text.Trim() != "")
        {
            strSql += " and D.Donate_Date <='" + txtDonateDateE.Text.Trim() + "' ";
        }
        if (ddlActName.SelectedIndex != 0)
        {
            strSql += " and D.Act_Id ='" + ddlActName.SelectedValue + "' ";
        }
        //捐款方式-----------------------------------------------------------------------------------//
        bool first = true; int cnt = 0;
        for (int i = 0; i < cblDonate_Payment.Items.Count; i++)
        {
            if (cblDonate_Payment.Items[i].Selected && first)
            {
                strSql += "and (D.Donate_Payment='" + cblDonate_Payment.Items[i].ToString() + "' "; first = false; cnt++;
            }
            else if (cblDonate_Payment.Items[i].Selected)
            {
                strSql += "or D.Donate_Payment='" + cblDonate_Payment.Items[i].ToString() + "' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //-------------------------------------------------------------------------------------------//
        //捐款用途-----------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblDonate_Purpose.Items.Count; i++)
        {
            if (cblDonate_Purpose.Items[i].Selected && first)
            {
                strSql += "and (D.Donate_Purpose='" + cblDonate_Purpose.Items[i].ToString() + "' "; first = false; cnt++;
            }
            else if (cblDonate_Purpose.Items[i].Selected)
            {
                strSql += "or D.Donate_Purpose='" + cblDonate_Purpose.Items[i].ToString() + "' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //-------------------------------------------------------------------------------------------//

        //日期範圍
        if (txtDonateDateS.Text.Trim() != "" && txtDonateDateE.Text.Trim() == "")
        {
            Session["Date"] = txtDonateDateS.Text.Trim() + "~";
        }
        else if (txtDonateDateS.Text.Trim() == "" && txtDonateDateE.Text.Trim() != "")
        {
            Session["Date"] = "~" + txtDonateDateE.Text.Trim();
        }
        else if (txtDonateDateS.Text.Trim() != "" && txtDonateDateE.Text.Trim() != "")
        {
            Session["Date"] = txtDonateDateS.Text.Trim() + "~" + txtDonateDateE.Text.Trim();
        }
        else
        {
            Session["Date"] = "無限制";
        }

        Session["strSql"] = strSql;
    }
}