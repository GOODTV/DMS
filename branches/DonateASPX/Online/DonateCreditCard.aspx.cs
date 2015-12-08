using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateCreditCard : System.Web.UI.Page
{
    DataTable dtOnce = new DataTable();
    DataTable dtPeriod = new DataTable();
    string strRemoteAddr = "";

    protected void Page_Load(object sender, EventArgs e)
    {
        strRemoteAddr = Request.ServerVariables["REMOTE_ADDR"];
        if (strRemoteAddr == "::1")
        {
            strRemoteAddr = "127.0.0.1";
        }

        if (Session["ItemOnce"] != null)
        {
            dtOnce = (DataTable)Session["ItemOnce"];
        }
        if (Session["ItemPeriod"] != null)
        {
            dtPeriod = (DataTable)Session["ItemPeriod"];
        }

        if (!IsPostBack)
        {
            Session["InsertPeriod"] = "N";

            LoadFormData();
            if (Session["DonorName"] != null)
            {
                lblTitle.Text = Session["DonorName"].ToString() + lblTitle.Text;
            }
            LoadDropDownList();
        }
    }
    //-------------------------------------------------------------------------------------------------------------
    private void LoadDropDownList()
    {
        ddlCardType.Items.Clear();
        ddlCardType.Items.Add(new ListItem("", ""));
        ddlCardType.Items.Add(new ListItem("VISA", "VISA"));
        ddlCardType.Items.Add(new ListItem("MASTER", "MASTER"));
        ddlCardType.Items.Add(new ListItem("JCB", "JCB"));

        ddlValidMonth.Items.Add(new ListItem("", ""));
        for (int i = 1; i < 13; i++)
        {
            ddlValidMonth.Items.Add(new ListItem(i.ToString("00"), i.ToString("00")));
        }

        ddlValidYear.Items.Add(new ListItem("", ""));
        for (int i = 0; i < 16; i++)
        {
            ddlValidYear.Items.Add(new ListItem((DateTime.Now.Year + i - 2000).ToString(), (DateTime.Now.Year + i - 2000).ToString()));
        }
    }
    //---------------------------------------------------------------------------
    private void LoadFormData()
    {

    }
    //---------------------------------------------------------------------------
    private void InsertPledgeTemp()
    {
        for (int i = 0; i < dtPeriod.Rows.Count; i++)
        {
            //string strPayment = dtPeriod.Rows[i]["奉獻金額"].ToString();
            Dictionary<string, object> dict = new Dictionary<string, object>();
            List<ColumnData> list = new List<ColumnData>();
            SetColumnData(list, dtPeriod.Rows[i]);
            string strSql = Util.CreateInsertCommand("PLEDGE_Temp", list, dict);
            NpoDB.ExecuteSQLS(strSql, dict);
            Session["Msg"] = "定期定額奉獻成功！";
            Session["InsertPeriod"] = "Y";
        }
    }
    //-------------------------------------------------------------------------------------------------------------
    private void SetColumnData(List<ColumnData> list, DataRow dr)
    {
        //Donor_Id	int	Checked 捐款人編號
        list.Add(new ColumnData("Donor_Id", Session["DonorID"].ToString(), true, false, true));
        //Donate_Payment	nvarchar(20)	Checked
        list.Add(new ColumnData("Donate_Payment", "信用卡授權書", true, false, false));
        //Donate_Purpose	nvarchar(20)	Checked 捐款用途
        list.Add(new ColumnData("Donate_Purpose", dr["奉獻項目"].ToString(), true, false, false));
        //Donate_Purpose_Type	nvarchar(10)	Checked 捐款用途ID
        list.Add(new ColumnData("Donate_Purpose_Type", "D", true, false, false));
        //Donate_Type	nvarchar(8)	Checked  D:Donate  
        list.Add(new ColumnData("Donate_Type", "長期捐款", true, false, false));
        //Donate_Amt	money	Checked
        list.Add(new ColumnData("Donate_Amt", dr["奉獻金額"].ToString(), true, false, false));
        //Donate_FirstAmt	money	Checked
        list.Add(new ColumnData("Donate_FirstAmt", "0", true, false, false));

        string strFromDate = "";
        string strToDate = "";
        string strBeginYear = "";
        string strBeginMonth = "";
        string strEndYear = "";
        string strEndMonth= "";
        switch (dr["繳費方式"].ToString())
        {
            case "月繳":
                strFromDate = dr["開始年"].ToString() + "/" + dr["開始月"].ToString() + "/01";
                strToDate = dr["結束年"].ToString() + "/" + dr["結束月"].ToString() + "/01";
                strBeginYear = dr["開始年"].ToString();
                strBeginMonth = dr["開始月"].ToString();
                strEndYear = dr["結束年"].ToString();
                strEndMonth = dr["結束月"].ToString();
                break;
            case "季繳":
                //int iFrom= (Convert.ToInt16(dr["開始月"].ToString()) - 1) * 3 + 1;
                //int iTo = Convert.ToInt16(dr["開始月"].ToString()) * 3;
                //strFromDate = dr["開始年"].ToString() + "/" + iFrom.ToString("00") + "/01";
                //strToDate = dr["結束年"].ToString() + "/" + iTo.ToString("00") + "/01";
                strFromDate = dr["開始年"].ToString() + "/" + dr["開始月"].ToString() + "/01";
                strToDate = dr["結束年"].ToString() + "/" + dr["結束月"].ToString() + "/01";
                strBeginYear = dr["開始年"].ToString();
                strBeginMonth = dr["開始月"].ToString();
                strEndYear = dr["結束年"].ToString();
                strEndMonth = dr["結束月"].ToString();
                break;
            case "年繳":
                strFromDate = dr["開始年"].ToString() + "/01/01";
                strToDate = dr["結束年"].ToString() + "/12/01";
                strBeginYear = dr["開始年"].ToString();
                strBeginMonth = "1";
                strEndYear = dr["結束年"].ToString();
                strEndMonth = "12";
                break;
        }
        
        //Donate_FromDate	datetime	Checked
        list.Add(new ColumnData("Donate_FromDate", strFromDate, true, false, false));
        //Donate_ToDate	datetime	Checked
        list.Add(new ColumnData("Donate_ToDate", strToDate, true, false, false));
        //Donate_Period	nvarchar(10)	Checked
        list.Add(new ColumnData("Donate_Period", dr["繳費方式"].ToString(), true, false, false));
        //Next_DonateDate	datetime	Checked
        //Card_Bank	nvarchar(20)	Checked
        list.Add(new ColumnData("Card_Bank", txtCardBank.Text.Trim(), true, false, false));
        //Card_Type	nvarchar(20)	Checked
        list.Add(new ColumnData("Card_Type", ddlCardType.SelectedValue, true, false, false));
        //Account_No	nvarchar(20)	Checked
        list.Add(new ColumnData("Account_No", txtAccountNo.Text.Trim(), true, false, false));
        //Valid_Date	nvarchar(4)	Checked
        list.Add(new ColumnData("Valid_Date", ddlValidMonth.SelectedValue + ddlValidYear.SelectedValue, true, false, false));
        //Card_Owner	nvarchar(20)	Checked
        list.Add(new ColumnData("Card_Owner", txtCardOwner.Text.Trim(), true, false, false));
        //Owner_IDNo	nvarchar(10)	Checked
        list.Add(new ColumnData("Owner_IDNo", txtOwnerIdNo.Text.Trim(), true, false, false));
        //Relation	nvarchar(10)	Checked
        list.Add(new ColumnData("Relation", txtRelation.Text.Trim(), true, false, false));
        //Authorize	nvarchar(10)	Checked        
        list.Add(new ColumnData("Authorize", txtAuthorize.Text.Trim(), true, false, false));
        //Create_Date	datetime	Checked
        list.Add(new ColumnData("Create_Date", Util.GetDBDateTime(), true, false, false));
        //Create_DateTime	datetime	Checked
        list.Add(new ColumnData("Create_DateTime", Util.GetDBDateTime(), true, false, false));
        //Create_User	nvarchar(20)	Checked
        list.Add(new ColumnData("Create_User", "Online", true, false, false));
        //Create_IP	nvarchar(16)	Checked
        list.Add(new ColumnData("Create_IP", strRemoteAddr, true, false, false));
        //Pledge_BeginYear	int	Checked
        list.Add(new ColumnData("Pledge_BeginYear", strBeginYear, true, false, false));
        //Pledge_BeginMonth	int	Checked
        list.Add(new ColumnData("Pledge_BeginMonth", strBeginMonth, true, false, false));
        //Pledge_EndYear	int	Checked
        list.Add(new ColumnData("Pledge_EndYear", strEndYear, true, false, false));
        //Pledge_EndMonth	int	Checked
        list.Add(new ColumnData("Pledge_EndMonth", strEndMonth, true, false, false));
        
        //Dept_Id  
        list.Add(new ColumnData("Dept_Id", "C001", true, false, false));
        //Status
        list.Add(new ColumnData("Status", "授權中", true, false, false));

    }
    //---------------------------------------------------------------------------
    protected void btnConfirm_Click(object sender, EventArgs e)
    {
        if (dtOnce.Rows.Count == 0 && dtPeriod.Rows.Count == 0)
        {
            Session["Msg"] = "無奉獻項目，請選擇奉獻項目！";
            Response.Redirect(Util.RedirectByTime("DonateOnlineAll.aspx"));            
        }
        InsertPledgeTemp();
        Session["ItemPeriod"] = null;
        if (dtOnce.Rows.Count > 0)
        {            
            Response.Redirect(Util.RedirectByTime("DonateInfoConfirm.aspx"));
        }
        else
        {
            Response.Redirect(Util.RedirectByTime("DonateDone.aspx"));
        }
    }
    //---------------------------------------------------------------------------
    protected void btnPrev_Click(object sender, EventArgs e)
    {
        //Response.Redirect("ShoppingCart.aspx?Mode=FromDefault");
        Response.Redirect(Util.RedirectByTime("DonateInfoConfirm.aspx"));
    }
}