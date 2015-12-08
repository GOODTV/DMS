using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Report_DonateMonthReport : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "DonateMonthQry";

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
        //Util.FillDropDownList(ddlActName, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        //ddlActName.Items.Insert(0, new ListItem("", ""));
        //ddlActName.SelectedIndex = 0;
    }
    public void LoadCheckBoxListData()
    {
        //捐款方式
        Util.FillCheckBoxList(cblDonate_Payment, Util.GetDataTable("CaseCode", "GroupName", "捐款方式", "ABS(CaseID)", ""), "CaseName", "CaseName", false);
        cblDonate_Payment.Items[0].Selected = false;

        //捐款用途
        Util.FillCheckBoxList(cblDonate_Purpose, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseName", false);
        cblDonate_Purpose.Items[0].Selected = false;

        //付款方式
        Util.FillCheckBoxList(cblPayment_type, Util.GetDataTable("Donate_IePayType", "Display", "1", "", ""), "CodeName", "CodeID", false);
        cblPayment_type.Items[0].Selected = false;
    }
    //----------------------------------------------------------------------
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        string strSql = Sql();
        strSql += Condition();
        strSql += Sql2();
        strSql += Condition();

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
            Session["Date"] = DateTime.Now.ToString();
        }
        strSql += " order by D.Invoice_No";
        Session["strSql"] = strSql;
        Response.Redirect("DonateMonthQry_Print_Excel.aspx");
    }
    protected void btnPrint_Click(object sender, EventArgs e)
    {
        string strSql = Sql();
        strSql += Condition();
        strSql += Sql2();
        strSql += Condition();

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
            Session["Date"] = DateTime.Now.ToString();
        }
        strSql += " order by D.Invoice_No";
        Session["strSql"] = strSql;
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
    private string Sql()
    {
        string strSql = @"select ROW_NUMBER() OVER(ORDER BY D.Invoice_No) AS 序號 , D.Invoice_No as 收據編號, CONVERT(VARCHAR(10) , D.Donate_Date, 111 ) as 捐款日期,
                        Dr.Donor_Name as 捐款人, D.Donor_Id as 編號, 
                        Case when Donate_Forign <> '' then REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Donate_Amt),1),'.00','')+'('+Donate_Forign+ISNULL(CONVERT(VARCHAR,CONVERT(MONEY,Donate_ForignAmt),1),'')+')' else REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,D.Donate_Amt),1),'.00','') end as 捐款金額,
                        D.Donate_Purpose as 捐款用途, D.Donate_Payment as 捐款方式,D.Invoice_Pre,dp.CodeName as 付款方式
                        from Donate D left join Donor Dr on D.Donor_Id = Dr.Donor_Id 
                        left join (select * from (select ROW_NUMBER() OVER(PARTITION by orderid ORDER BY Ser_No) AS ROWID,*
                        from DONATE_IEPAY) as P1 where ROWID = 1) as di on d.od_sob = di.orderid 
                        left join Donate_IePayType dp on di.paytype = dp.CodeID 
                        where Dr.DeleteDate is null and D.Issue_Type != 'D'";

        return strSql;
    }
    private string Sql2()
    {
        string strSql = @"union
                        select '99999999' as [序號] , '列印日期：' as [收據編號], CONVERT(VARCHAR, GETDATE(), 120 ) as [捐款日期],'總計' as [捐款人],'' as [編號],
                        REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(D.Donate_Amt)),1),'.00','') as [捐款金額], '' as [捐款用途],'','' ,''  
                         from Donate D left join Donor Dr on D.Donor_Id = Dr.Donor_Id 
                         left join (select * from (select ROW_NUMBER() OVER(PARTITION by orderid ORDER BY Ser_No) AS ROWID,*
                         from DONATE_IEPAY) as P1 where ROWID = 1) as di on d.od_sob = di.orderid 
                         left join Donate_IePayType dp on di.paytype = dp.CodeID 
                         where Dr.DeleteDate is null and D.Issue_Type != 'D'";
                        return strSql;
    }
    private string Condition()
    {
        string strSql="";
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and Dr.Dept_Id='" + ddlDept.SelectedValue + "' ";
        }
        if (txtInvoice_NoB.Text.Trim() != "")
        {
            strSql += " and D.Invoice_No >= '" + txtInvoice_NoB.Text.Trim() + "' ";
        }
        if (txtInvoice_NoE.Text.Trim() != "")
        {
            strSql += " and D.Invoice_No <= '" + txtInvoice_NoE.Text.Trim() + "' ";
        }
        if (txtDonateDateS.Text.Trim() != "" && txtDonateDateE.Text.Trim() == "")
        {
            strSql += " and D.Donate_Date between '" + txtDonateDateS.Text.Trim() + "' and getdate() ";
        }
        if (txtDonateDateS.Text.Trim() == "" && txtDonateDateE.Text.Trim() != "")
        {
            strSql += " and D.Donate_Date between '1900/01/01' and'" + txtDonateDateE.Text.Trim() + "' ";
        }
        if (txtDonateDateS.Text.Trim() != "" && txtDonateDateE.Text.Trim() != "")
        {
            strSql += " and D.Donate_Date between '" + txtDonateDateS.Text.Trim() + "' and '" + txtDonateDateE.Text.Trim() + "' ";
        }
        //if (ddlActName.SelectedIndex != 0)
        //{
        //    strSql += " and D.Act_Id ='" + ddlActName.SelectedValue + "' ";
        //}
        if (tbxDonate_AmtB.Text.Trim() != "")
        {
            strSql += " and D.Donate_Amt >= '" + tbxDonate_AmtB.Text.Trim() + "' ";
        }
        if (tbxDonate_AmtE.Text.Trim() != "")
        {
            strSql += " and D.Donate_Amt <= '" + tbxDonate_AmtE.Text.Trim() + "' ";
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
        //付款方式-----------------------------------------------------------------------------------//
        first = true; cnt = 0;
        for (int i = 0; i < cblPayment_type.Items.Count; i++)
        {
            if (cblPayment_type.Items[i].Selected && first)
            {
                strSql += "and (di.paytype='" + cblPayment_type.Items[i].Value.ToString() + "' "; first = false; cnt++;
            }
            else if (cblPayment_type.Items[i].Selected)
            {
                strSql += "or di.paytype='" + cblPayment_type.Items[i].Value.ToString() + "' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //-------------------------------------------------------------------------------------------//
        //區域 20141029新增
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
        //-------------------------------------------------------------------------------------------//
        //匯入台銀回覆檔檔名 2015/1/8 新增
        if (tbPledgeBatchFileName.Text.Trim() != "")
        {
            strSql += " and D.PledgeBatchFileName = '" + tbPledgeBatchFileName.Text.Trim() + "' ";
        }

        return strSql;
    }

}
