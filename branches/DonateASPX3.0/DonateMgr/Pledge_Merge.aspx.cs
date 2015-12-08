using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;

public partial class DonateMgr_Pledge_Merge : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "Pledge_Merge";
        //權控處理
        AuthrityControl();
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            //資料合併按鈕
            btntransfer.Visible = false;
        }
    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_Query", btnQuery);
        Authrity.CheckButtonRight("_Update", btntransfer);
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        //轉出人
        string strSql_From;
        DataTable dt_From;
        strSql_From = @"Select * ,(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) as [通訊地址]
                        From Donor 
                            Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode
                        where Donor_Id = @Donor_Id and DeleteDate is null";

        Dictionary<string, object> dict_From = new Dictionary<string, object>();
        dict_From.Add("Donor_Id", tbxFrom_Donor_Id.Text.Trim());
        dt_From = NpoDB.GetDataTableS(strSql_From, dict_From);
        DataRow dr_From;

        if (dt_From.Rows.Count == 0)
        {
            SetSysMsg("您輸入的『轉出』捐款人編號不存在，請您重新確認！");
            Response.Redirect("Donate_Merge.aspx");
        }
        else
        {
            //轉出人基本資料
            dr_From = dt_From.Rows[0];
            lblDonor_From.Text = "轉出捐款人：" + dr_From["Donor_Name"].ToString() + "<br>";
            lblDonor_From.Text += "　通訊地址：" + dr_From["通訊地址"].ToString() + "<br>";
            lblDonor_From.Text += "欲轉出資料：<span class=\"necessary\">( 請先勾選然後按下『 資料合併 』按鈕) </span>";

            //轉出人的轉帳授權書資料
            strSql_From = @"select P.Pledge_Id as [Pledge_Id], Donate_Payment as [授權方式],REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','') as [扣款金額], 
                                   CONVERT(VARCHAR(10) , Donate_FromDate, 111 ) as [授權起日期], CONVERT(VARCHAR(10) , [Donate_ToDate], 111 ) as [授權迄日期],
                                   De.DeptShortName as [機構]
                            from Pledge P
                                   left join Donor Dr on P.Donor_Id = Dr.Donor_Id 
                                   left join Dept De on Dr.Dept_Id = De.DeptId
                            where P.Donor_Id = @Donor_Id and Dr.DeleteDate is null";
            dict_From = new Dictionary<string, object>();
            dict_From.Add("Donor_Id", tbxFrom_Donor_Id.Text.Trim());
            dt_From = NpoDB.GetDataTableS(strSql_From, dict_From);
            if (dt_From.Rows.Count == 0)
            {
                ShowSysMsg("您輸入的『轉出』捐款人編號無轉帳授權書紀錄，請您重新確認！");
                btntransfer.Visible = false;
            }
            else
            {
                btntransfer.Visible = true;
            }
            lblData_From.Text = CheckBoxData(strSql_From, "Pledge_Id", 0, dt_From);
        }

        //轉入人
        string strSql_To;
        DataTable dt_To;
        strSql_To = @"Select * ,(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) as [通訊地址]
                      From Donor 
                          Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode
                      where Donor_Id = @Donor_Id and DeleteDate is null";

        Dictionary<string, object> dict_To = new Dictionary<string, object>();
        dict_To.Add("Donor_Id", tbxTo_Donor_Id.Text.Trim());
        dt_To = NpoDB.GetDataTableS(strSql_To, dict_To);
        DataRow dr_To;

        if (dt_To.Rows.Count == 0)
        {

            SetSysMsg("您輸入的『轉入』捐款人編號不存在，請您重新確認！");
            Response.Redirect("Pledge_Merge.aspx");
        }
        else
        {
            dr_To = dt_To.Rows[0];
            //轉入人基本資料
            dr_To = dt_To.Rows[0];
            lblDonor_To.Text = "轉入捐款人：" + dr_To["Donor_Name"].ToString() + "<br>";
            lblDonor_To.Text += "　通訊地址：" + dr_To["通訊地址"].ToString() + "<br>";
            lblDonor_To.Text += "欲轉入資料：";

            //轉入人的捐款資料
            strSql_To = @"select P.Pledge_Id as [Pledge_Id], Donate_Payment as [授權方式],REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','') as [扣款金額], 
                                   CONVERT(VARCHAR(10) , Donate_FromDate, 111 ) as [授權起日期], CONVERT(VARCHAR(10) , [Donate_ToDate], 111 ) as [授權迄日期],
                                   De.DeptShortName as [機構]
                            from Pledge P
                                   left join Donor Dr on P.Donor_Id = Dr.Donor_Id 
                                   left join Dept De on Dr.Dept_Id = De.DeptId
                            where P.Donor_Id = @Donor_Id and Dr.DeleteDate is null";
            dict_To = new Dictionary<string, object>();
            dict_To.Add("Donor_Id", tbxTo_Donor_Id.Text.Trim());
            dt_To = NpoDB.GetDataTableS(strSql_To, dict_To);
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt_To;
            npoGridView.ShowPage = false;
            npoGridView.DisableColumn.Add("Pledge_Id");
            lblData_To.Text = npoGridView.Render();
        }

    }
    //---------------------------------------------------------------------------
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    protected void btntransfer_Click(object sender, EventArgs e)
    {
        string Pledge_Id = CatchPledge_Id();

        if (Pledge_Id == "")
        {
            ShowSysMsg("您尚未勾選任何轉出資料");
            return;
        }
        else
        {
            //有幾筆資料
            string[] Pledge_Ids = Pledge_Id.Split(' ');

            for (int i = 0; i < Pledge_Ids.Length; i++)
            {
                string strSql = @" update Pledge 
                                        set Donor_Id = @Donor_Id
                                        ,LastUpdate_Date = @LastUpdate_Date
                                        ,LastUpdate_DateTime = @LastUpdate_DateTime
                                        ,LastUpdate_User = @LastUpdate_User
                                        ,LastUpdate_IP= @LastUpdate_IP 
                                        where Pledge_Id = @Pledge_Id";
                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Donor_Id", tbxTo_Donor_Id.Text);
                dict.Add("LastUpdate_Date", DateTime.Now.ToString("yyyy-MM-dd"));
                dict.Add("LastUpdate_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
                dict.Add("LastUpdate_User", SessionInfo.UserName);
                dict.Add("LastUpdate_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
                dict.Add("Pledge_Id", Pledge_Ids[i]);

                NpoDB.ExecuteSQLS(strSql, dict);
            }
            int Count = Pledge_Ids.Length - 1;
            ShowSysMsg("轉帳授權書資料已合併成功，共計" + Count + " 筆");
            LoadFormData();
        }
    }
    public string CheckBoxData(string strSql, string FName, int DataNum, DataTable dt)
    {

        StringBuilder sb = new StringBuilder();

        //k用來決定第幾欄以後的內容是特殊內容(例如文字框(textbox))
        int k;

        k = dt.Columns.Count - DataNum;
        sb.AppendLine(@"<table width='100%' border='0' cellpadding='0' cellspacing='1' class='table_h'>");
        sb.AppendLine(@"<tr>");
        //20140430 Modify by GoodTV Tanya:增加「全選」的功能
        //sb.AppendLine(@"<th nowrap><span style='font-size: 9pt; font-family: 新細明體'>選擇</span></th>");
        sb.AppendLine(@"<th nowrap><span style='font-size: 9pt; font-family: 新細明體'><Input Type='checkbox' Name='checkboxAll' Id='checkboxAll' checked></span></th>");


        int c = 0;
        foreach (DataColumn dc in dt.Columns)
        {
            if (c == 0)
            {
                c++;
            }
            else
            {
                sb.AppendLine(@"<th nowrap><span style='font-size: 9pt; font-family: 新細明體'>" + dc.ColumnName + "</span></th>");
            }
        }
        sb.AppendLine("</tr>");
        int j = 1;
        int i = 1;
        foreach (DataRow dr in dt.Rows)
        {
            sb.AppendLine(@"<tr>");
            //20140430 Modify by GoodTV Tanya:增加「全選」的功能
            sb.AppendLine(@"<td nowrap><span style='font-size: 9pt; font-family: 新細明體'><Input Type='checkbox' Id='checkbox' Name='" + FName + dr[FName].ToString() + @"' value='" + j + "' checked></span></td>");
            i = 1;
            c = 0;
            foreach (DataColumn dc in dt.Columns)
            {
                if (c != 0)
                {
                    if (i > k)
                        sb.AppendLine(@"<td nowrap><span style='font-size: 9pt; ; font-family: 新細明體'><input type='text' name='" + i + "_" + j + @"' size='4' style='font-size: 9pt; font-family: 新細明體' value='" + dr[dc.ColumnName] + @"'></span></td>");
                    else
                        sb.AppendLine(@"<td nowrap><span style='font-size: 9pt; font-family: 新細明體'>" + dr[dc.ColumnName] + "</span></td>");
                }

                c++;
                i++;
            }
            j++;
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</table>");
        return sb.ToString();
    }
    //抓轉帳授權書編號
    private string CatchPledge_Id()
    {
        string Pledge_Id = "";
        string strPledge_Id = "";

        foreach (string str in Request.Form.Keys)
        {
            if (str.StartsWith("Pledge_Id"))
            {
                strPledge_Id = str.Split('d')[2];
                Pledge_Id += strPledge_Id + " ";
            }
        }
        return Pledge_Id;
    }
}