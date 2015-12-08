using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;

public partial class DonateMgr_Donate_Merge : BasePage
{

    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "Donate_Merge";
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
        strSql_From = @"Select * ,(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+Address Else DONOR.ZipCode+A.mValue+Address End End) as [通訊地址]
                        From Donor 
                            Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode
                        where Donor_Id = @Donor_Id and DeleteDate is null";

        Dictionary<string, object> dict_From = new Dictionary<string, object>();
        dict_From.Add("Donor_Id", tbxFrom_Donor_Id.Text.Trim());
        dt_From = NpoDB.GetDataTableS(strSql_From, dict_From);
        DataRow dr_From;

        if(dt_From.Rows.Count==0)
        {
            SetSysMsg("您輸入的『轉出』捐款人編號不存在，請您重新確認！");
            Response.Redirect("Donate_Merge.aspx");
        }
        else
        {
            //轉出人基本資料
            dr_From = dt_From.Rows[0] ;
            lblDonor_From.Text = "轉出捐款人：" + dr_From["Donor_Name"].ToString() + "<br>";
            lblDonor_From.Text += "　通訊地址：" + dr_From["通訊地址"].ToString() + "<br>";
            if (!String.IsNullOrEmpty(dr_From["Donor_Pwd"].ToString()))
            {
                lblDonor_From.Text += "　線上奉獻會員帳號：" + dr_From["Email"].ToString() + "&nbsp;&nbsp;";
                lblDonor_From.Text += "<font color='blue'><INPUT id='ckEmail' type='checkbox' name='ckEmail'>帳號是否轉出(覆蓋轉入帳號)？</font><br>";
            }
            lblDonor_From.Text += "欲轉出資料：<span class=\"necessary\">( 請先勾選然後按下『 資料合併 』按鈕) </span>";
            
            //轉出人的捐款資料
            strSql_From = @"select Do.Donate_Id as [Donate_Id], CONVERT(VARCHAR(10) , Do.Donate_Date, 111 ) as [捐款日期], Do.Donate_Payment as [捐款方式], 
                                   REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Do.Donate_Amt),1),'.00','') as [捐款金額], 
                                   De.DeptShortName as [機構], IsNull(Do.Invoice_Pre,'') + Do.Invoice_No as [收據編號] 
                            from Donate Do
                                   left join Donor Dr on Do.Donor_Id = Dr.Donor_Id 
                                   left join Dept De on Do.Dept_Id = De.DeptId
                            where Do.Donor_Id = @Donor_Id and Dr.DeleteDate is null --and IsNull(Do.Issue_Type,'') <> 'D' 
                            order by Do.Donate_Date";
            dict_From = new Dictionary<string, object>();
            dict_From.Add("Donor_Id", tbxFrom_Donor_Id.Text.Trim());
            dt_From = NpoDB.GetDataTableS(strSql_From, dict_From);
            if (dt_From.Rows.Count == 0)
            {
                ShowSysMsg("您輸入的『轉出』捐款人編號無捐款紀錄，請您重新確認！");
                btntransfer.Visible = false;
            }
            else
            {
                btntransfer.Visible = true;
            }
            lblData_From.Text = CheckBoxData(strSql_From, "Donate_Id", 0, dt_From);
        }

        //轉入人
        string strSql_To;
        DataTable dt_To;
        strSql_To = @"Select * ,(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+Address Else DONOR.ZipCode+A.mValue+Address End End) as [通訊地址]
                      From Donor 
                          Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode
                      where Donor_Id = @Donor_Id and DeleteDate is null";

        Dictionary<string, object> dict_To = new Dictionary<string, object>();
        dict_To.Add("Donor_Id", tbxTo_Donor_Id.Text.Trim());
        dt_To = NpoDB.GetDataTableS(strSql_To, dict_To);
        DataRow dr_To;

        if(dt_To.Rows.Count==0)
        {
            
            SetSysMsg("您輸入的『轉入』捐款人編號不存在，請您重新確認！");
            Response.Redirect("Donate_Merge.aspx");
        }
        else
        {
            dr_To = dt_To.Rows[0] ;
            //轉入人基本資料
            //20140430 Modify by GoodTV：修改標題內容
            lblDonor_To.Text = "轉入捐款人：" + dr_To["Donor_Name"].ToString() + "<br>";
            lblDonor_To.Text += "　通訊地址：" + dr_To["通訊地址"].ToString() + "<br>";
            if (!String.IsNullOrEmpty(dr_To["Donor_Pwd"].ToString()))
            {
                lblDonor_To.Text += "　線上奉獻會員帳號：" + dr_To["Email"].ToString() + "<br>";
            }
            lblDonor_To.Text += "欲轉入資料：";

            // 2014/7/9 增加存入捐款人姓名以利轉入的捐款紀錄更新收據抬頭
            if (dr_To["Invoice_Title"].ToString() != "")
            {
                hfInvoice_Title.Value = dr_To["Invoice_Title"].ToString();
            }
            else
            {
                hfInvoice_Title.Value = dr_To["Donor_Name"].ToString();
            }
            // 2014/10/3 增加存入收據開立以利轉入的捐款記錄更新收據開立
            if (dr_To["Invoice_Type"].ToString() != "")
            {
                hfInvoice_Type.Value = dr_To["Invoice_Type"].ToString();
            }

            //轉入人的捐款資料
            strSql_To = @"select Do.Donate_Id as [Donate_Id], CONVERT(VARCHAR(10) , Do.Donate_Date, 111 ) as [捐款日期], Do.Donate_Payment as [捐款方式], 
                                   REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Do.Donate_Amt),1),'.00','') as [捐款金額], 
                                   De.DeptShortName as [機構], IsNull(Do.Invoice_Pre,'') + Do.Invoice_No as [收據編號] 
                            from Donate Do
                                   left join Donor Dr on Do.Donor_Id = Dr.Donor_Id 
                                   left join Dept De on Do.Dept_Id = De.DeptId
                            where Do.Donor_Id = @Donor_Id and Dr.DeleteDate is null --and IsNull(Do.Issue_Type,'') <> 'D' 
                            order by Do.Donate_Date";
            dict_To = new Dictionary<string, object>();
            dict_To.Add("Donor_Id", tbxTo_Donor_Id.Text.Trim());
            dt_To = NpoDB.GetDataTableS(strSql_To, dict_To);
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt_To;
            npoGridView.ShowPage = false;
            npoGridView.DisableColumn.Add("Donate_Id");
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
        string Donate_Id = CatchDonate_Id();

        if (Donate_Id == "")
        {
            ShowSysMsg("您尚未勾選任何轉出資料");
            return;
        }
        else
        {
            //有幾筆資料
            string[] Donate_Ids = Donate_Id.Split(' ');

            for (int i = 0; i < Donate_Ids.Length; i++)
            {
                string strSql = @" update Donate 
                                    set Donor_Id = @Donor_Id
                                        ,Invoice_Title = @Invoice_Title
                                        ,Invoice_Type = @Invoice_Type 
                                        ,LastUpdate_Date = @LastUpdate_Date
                                        ,LastUpdate_DateTime = @LastUpdate_DateTime
                                        ,LastUpdate_User = @LastUpdate_User
                                        ,LastUpdate_IP= @LastUpdate_IP
                                    where Donate_Id = @Donate_Id";
                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Donor_Id", tbxTo_Donor_Id.Text);
                dict.Add("Invoice_Title", hfInvoice_Title.Value);
                dict.Add("Invoice_Type", hfInvoice_Type.Value);
                dict.Add("LastUpdate_Date", DateTime.Now.ToString("yyyy-MM-dd"));
                dict.Add("LastUpdate_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
                dict.Add("LastUpdate_User", SessionInfo.UserName);
                dict.Add("LastUpdate_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
                dict.Add("Donate_Id", Donate_Ids[i]);

                NpoDB.ExecuteSQLS(strSql, dict);

            }

            // 2014/7/9 更正 捐款紀錄合併後，重新統計轉出與轉入捐款人的數據
            UpdateDonor(tbxFrom_Donor_Id.Text);
            UpdateDonor(tbxTo_Donor_Id.Text);

            // 2015/1/6 轉出轉入線上奉獻會員帳號
            if (hfCheckEmail.Value == "Emailok")
            {
                UpdateDonorEmail(tbxFrom_Donor_Id.Text, tbxTo_Donor_Id.Text);
            }

            int Count = Donate_Ids.Length - 1;
            ShowSysMsg("捐款資料已合併成功，共計" + Count + " 筆");
            LoadFormData();
        }
    }

    private void UpdateDonor(string strDonor_Id)
    {

        // 2014/7/9 參照ASP2.0
        string strSQL = @"DECLARE @Begin_DonateDate datetime " +
       "DECLARE @Last_DonateDate datetime " +
       "DECLARE @Donate_No numeric " +
       "DECLARE @Donate_Total numeric " +
       "Select Top 1 @Begin_DonateDate=Donate_Date From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 Order By Donate_Date " +
       "Select Top 1 @Last_DonateDate=Donate_Date From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 Order By Donate_Date Desc " +
       "Select @Donate_No=Count(*) From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 " +
       "Select @Donate_Total=IsNull(Sum(Donate_Amt),0) From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 " +
       "Update DONOR Set Begin_DonateDate=@Begin_DonateDate,Last_DonateDate=@Last_DonateDate,Donate_No=@Donate_No,Donate_Total=@Donate_Total Where Donor_Id=@Donor_Id ; ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", strDonor_Id);
        NpoDB.ExecuteSQLS(strSQL, dict);

    }

    private void UpdateDonorEmail(string strDonor_Id_From, string strDonor_Id_To)
    {

        string strSQL = @"update d2 set [Email]=d1.[Email],[Donor_Pwd]=d1.[Donor_Pwd] 
        from [Donor] as d1,[Donor] as d2 where d1.Donor_Id = @Donor_Id_From and d2.Donor_Id = @Donor_Id_To ;
        Update DONOR Set Donor_Pwd = null Where Donor_Id=@Donor_Id_From ; ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id_To", strDonor_Id_To);
        dict.Add("Donor_Id_From", strDonor_Id_From);
        NpoDB.ExecuteSQLS(strSQL, dict);

    }

    public string CheckBoxData(string strSql, string FName, int DataNum,  DataTable dt)
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
    //抓捐款編號
    private string CatchDonate_Id()
    {
        string Donate_Id = "";
        string strDonate_Id = "";

        foreach (string str in Request.Form.Keys)
        {
            if (str.StartsWith("Donate_Id"))
            {
                strDonate_Id = str.Split('d')[1];
                Donate_Id += strDonate_Id + " ";
            }
        }
        return Donate_Id;
    }
}