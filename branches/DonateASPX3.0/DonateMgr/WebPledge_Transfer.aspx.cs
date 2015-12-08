using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonateMgr_WebPledge_Transfer : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Pledge_Id");
            Transfer();
        }
    }
    public void Transfer()
    {
        bool flag = false;
        try
        {
            AddNew();
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            Del();
            SetSysMsg("線上轉帳授權書授權成功！");
            Response.Redirect(Util.RedirectByTime("WebPledgeList.aspx"));
        }
    }
    public void AddNew()
    {
        //查詢線上授權書的資料
        string strSql = "select * from Pledge_Temp where Pledge_Id = '" + HFD_Uid.Value + "'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];

        //新增線上授權書的資料到轉帳授權書的表中
        strSql = "insert into  Pledge\n";
        strSql += "( Pledge_Id, Donor_Id, Donate_Payment, Donate_Purpose, Donate_Purpose_Type, Donate_Type, Donate_Amt, Donate_FirstAmt, Donate_FromDate,\n";
        strSql += " Donate_ToDate, Donate_Period, Next_DonateDate, Card_Bank, Card_Type, Account_No, Valid_Date, Card_Owner,\n";
        strSql += " Owner_IDNo, Relation, Authorize, Post_Name, Post_IDNo, Post_SavingsNo, Post_AccountNo, Dept_Id, Invoice_Type,\n";
        strSql += " Invoice_Title, Accoun_Bank, Accounting_Title, Act_id, Status, Comment, Create_Date, Create_DateTime,\n";
        strSql += " Create_User, Create_IP,LastUpdate_Date,LastUpdate_DateTime,LastUpdate_User,LastUpdate_IP,\n";
        strSql += " Pledge_BeginYear,Pledge_BeginMonth,Pledge_EndYear,Pledge_EndMonth,P_BANK, P_RCLNO, P_PID) values\n";
        strSql += "( (SELECT MAX(Pledge_Id) FROM Pledge)+1,@Donor_Id,@Donate_Payment,@Donate_Purpose,@Donate_Purpose_Type,@Donate_Type,@Donate_Amt,@Donate_FirstAmt,@Donate_FromDate,\n";
        strSql += "@Donate_ToDate,@Donate_Period,@Next_DonateDate,@Card_Bank,@Card_Type,@Account_No,@Valid_Date,@Card_Owner,\n";
        strSql += "@Owner_IDNo,@Relation,@Authorize,@Post_Name,@Post_IDNo,@Post_SavingsNo,@Post_AccountNo,@Dept_Id,@Invoice_Type,\n";
        strSql += "@Invoice_Title,@Accoun_Bank,@Accounting_Title,@Act_id,@Status,@Comment, @Create_Date,@Create_DateTime,\n";
        strSql += "@Create_User,@Create_IP,@LastUpdate_Date,@LastUpdate_DateTime,@LastUpdate_User,@LastUpdate_IP,@Pledge_BeginYear,\n";
        strSql += "@Pledge_BeginMonth,@Pledge_EndYear,@Pledge_EndMonth,@P_BANK,@P_RCLNO,@P_PID) ";
        strSql += "\n";
        strSql += "SELECT MAX(Pledge_Id) FROM Pledge";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", dr["Donor_Id"].ToString());
        dict.Add("Donate_Payment", dr["Donate_Payment"].ToString());
        dict.Add("Donate_Purpose", dr["Donate_Purpose"].ToString());
        dict.Add("Donate_Purpose_Type", dr["Donate_Purpose_Type"].ToString());
        dict.Add("Donate_Type", dr["Donate_Type"].ToString());
        dict.Add("Donate_Amt", dr["Donate_Amt"].ToString());
        dict.Add("Donate_FirstAmt", dr["Donate_FirstAmt"].ToString());
        dict.Add("Donate_FromDate", Util.FixDateTime(dr["Donate_FromDate"].ToString()));
        dict.Add("Donate_ToDate", Util.FixDateTime(dr["Donate_ToDate"].ToString()));
        dict.Add("Donate_Period", dr["Donate_Period"].ToString());
        dict.Add("Next_DonateDate", Util.FixDateTime(dr["Next_DonateDate"].ToString()));
        dict.Add("Card_Bank", dr["Card_Bank"].ToString());
        dict.Add("Card_Type", dr["Card_Type"].ToString());
        dict.Add("Account_No", dr["Account_No"].ToString());
        dict.Add("Valid_Date", dr["Valid_Date"].ToString());
        dict.Add("Card_Owner", dr["Card_Owner"].ToString());
        dict.Add("Owner_IDNo", dr["Owner_IDNo"].ToString());
        dict.Add("Relation", dr["Relation"].ToString());
        dict.Add("Authorize", dr["Authorize"].ToString());
        dict.Add("Post_Name", dr["Post_Name"].ToString());
        dict.Add("Post_IDNo", dr["Post_IDNo"].ToString());
        dict.Add("Post_SavingsNo", dr["Post_SavingsNo"].ToString());
        dict.Add("Post_AccountNo", dr["Post_AccountNo"].ToString());
        dict.Add("P_BANK", dr["P_BANK"].ToString());
        dict.Add("P_RCLNO", dr["P_RCLNO"].ToString());
        dict.Add("P_PID", dr["P_PID"].ToString());
        dict.Add("Dept_Id", dr["Dept_Id"].ToString());
        dict.Add("Invoice_Type", dr["Invoice_Type"].ToString());
        dict.Add("Invoice_Title", dr["Invoice_Title"].ToString());
        dict.Add("Accoun_Bank", dr["Accoun_Bank"].ToString());
        dict.Add("Accounting_Title", dr["Accounting_Title"].ToString());
        dict.Add("Act_Id", dr["Act_Id"].ToString());
        dict.Add("Status", "授權中");
        dict.Add("Comment", dr["Comment"].ToString());
        dict.Add("Create_Date", Util.FixDateTime(dr["Create_Date"].ToString()));
        dict.Add("Create_DateTime", Util.FixDateTime(dr["Create_DateTime"].ToString()));
        dict.Add("Create_User", dr["Create_User"].ToString());
        dict.Add("Create_IP", dr["Create_IP"].ToString());
        dict.Add("LastUpdate_Date", dr["LastUpdate_Date"].ToString());
        dict.Add("LastUpdate_DateTime", dr["LastUpdate_DateTime"].ToString());
        dict.Add("LastUpdate_User", dr["LastUpdate_User"].ToString());
        dict.Add("LastUpdate_IP", dr["LastUpdate_User"].ToString());
        dict.Add("Pledge_BeginYear", dr["Pledge_BeginYear"].ToString());
        dict.Add("Pledge_BeginMonth", dr["Pledge_BeginMonth"].ToString());
        dict.Add("Pledge_EndYear", dr["Pledge_EndYear"].ToString());
        dict.Add("Pledge_EndMonth", dr["Pledge_EndMonth"].ToString());
                
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void Del()
    {
        string strSql = "delete from Pledge_Temp where Pledge_Id=@Pledge_Id";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Pledge_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
}