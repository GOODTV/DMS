using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;
using System.Web.UI.HtmlControls;

public partial class DonateMgr_Pledge_Transfer : BasePage
{

    string TableWidth = "1450px";

    #region NpoGridView 處理換頁相關程式碼
    Button btnNextPage, btnPreviousPage, btnGoPage;
    HiddenField HFD_CurrentPage, HFD_CurrentQuerye;

    override protected void OnInit(EventArgs e)
    {
        CreatePageControl();
        base.OnInit(e);
    }

    private void CreatePageControl()
    {
        // Create dynamic controls here.
        btnNextPage = new Button();
        btnNextPage.ID = "btnNextPage";
        Form1.Controls.Add(btnNextPage);
        btnNextPage.Click += new System.EventHandler(btnNextPage_Click);

        btnPreviousPage = new Button();
        btnPreviousPage.ID = "btnPreviousPage";
        Form1.Controls.Add(btnPreviousPage);
        btnPreviousPage.Click += new System.EventHandler(btnPreviousPage_Click);

        btnGoPage = new Button();
        btnGoPage.ID = "btnGoPage";
        Form1.Controls.Add(btnGoPage);
        btnGoPage.Click += new System.EventHandler(btnGoPage_Click);

        HFD_CurrentPage = new HiddenField();
        HFD_CurrentPage.Value = "1";
        HFD_CurrentPage.ID = "HFD_CurrentPage";
        Form1.Controls.Add(HFD_CurrentPage);

        HFD_CurrentQuerye = new HiddenField();
        HFD_CurrentQuerye.Value = "Query";
        HFD_CurrentQuerye.ID = "HFHFD_CurrentQuerye";
        Form1.Controls.Add(HFD_CurrentQuerye);
    }

    protected void btnPreviousPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.MinusStringNumber(HFD_CurrentPage.Value);
        LoadFormData();
    }

    protected void btnNextPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.AddStringNumber(HFD_CurrentPage.Value);
        LoadFormData();
    }

    protected void btnGoPage_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    #endregion NpoGridView 處理換頁相關程式碼

    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {

        Session["ProgID"] = "Pledge_Transfer";
        //權控處理
        AuthrityControl();
        ////有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
        if (!IsPostBack)
        {
            //Util.ShowSysMsgWithScript("");
            LoadDropDownListData();
            //20140518 修改by Ian 移動至PostBack外
            //Transfer_Date();
            // 2014/4/10 修正進入時預設不載入資料
            //LoadFormData();
            HFD_Dept_Id.Value = SessionInfo.DeptID;
            //20140515 預設筆數和金額為0
            lblAccount_Amount.Text = "0";
            lblAccount_Count.Text = "0";
            lblACH_Amount.Text = "0";
            lblACH_Count.Text = "0";
            lblCard_Amount.Text = "0";
            lblCard_Count.Text = "0";
            //20140518 修改by Ian 
            Transfer_Date();
            //ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >Window_OnLoad();</script>");

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
        Authrity.CheckButtonRight("_Query", btnPrint);
        Authrity.CheckButtonRight("_AddNew", btnInput);
        Authrity.CheckButtonRight("_AddNew", btnImport);
        Authrity.CheckButtonRight("_AddNew", btnImportCathayTXT);
        Authrity.CheckButtonRight("_Query", btnExportErrorExcel);
        Authrity.CheckButtonRight("_Query", btnImport_CheckFalse);
    }

    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.Items.Insert(0, new ListItem("", ""));
        ddlDept.SelectedIndex = 1;

        //TXT格式
        string strSql = "Select Transfer_AspxName,Transfer_BankName From Donate_Transfer Where Transfer_StopFlg='N' Order By Transfer_Seq,Ser_No Desc";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        Util.FillDropDownList(ddlFormat, dt, "Transfer_BankName", "Transfer_AspxName", true);
        ddlFormat.SelectedIndex = 0;
        
        //轉帳週期
        Util.FillDropDownList(ddlDonate_Period, Util.GetDataTable("CaseCode", "GroupName", "轉帳週期", "", ""), "CaseName", "CaseName", false);
        ddlDonate_Period.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Period.SelectedIndex = 0;

        //授權方式
        Util.FillDropDownList(ddlDonate_Payment, Util.GetDataTable("CaseCode", "GroupName", "授權方式", "", ""), "CaseName", "CaseName", false);
        ddlDonate_Payment.Items.Insert(0, new ListItem("", ""));
        // 2014/4/10 修改一個信用卡授權書變成二個不同信用卡授權書的名單，及排序
        ddlDonate_Payment.SelectedIndex = 0;
    }

    public void LoadFormData()
    {
        string strSql = Sql();
        strSql += SqlCondition();
        string strSqlOrder = " order by Pledge_Id";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql + strSqlOrder, dict);

        if (dt.Rows.Count == 0)
        {
            lblAccount_Count.Text = "0";
            lblAccount_Amount.Text = "0";
            lblCard_Count.Text = "0";
            lblCard_Amount.Text = "0";
            lblACH_Count.Text = "0";
            lblACH_Amount.Text = "0";
            lblGridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
            HFD_Flag.Value = "True";
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.Keys.Add("Pledge_Id");
            //2014/4/21 修改by Ian_Kao 不要分頁以免勾選全選後無法其他分頁的資料
            //npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            npoGridView.ShowPage = false;
            npoGridView.DisableColumn.Add("Pledge_Id");
            npoGridView.DisableColumn.Add("disabled");
            //-------------------------------------------------------------------------
            // 2014/4/11 修正標題換成為checkbox
            //NPOGridViewColumn col = new NPOGridViewColumn("選擇");
            NPOGridViewColumn col = new NPOGridViewColumn();
            col.ColumnType = NPOColumnType.Checkbox;
            col.ControlKeyColumn.Add("Pledge_Id");
            col.CaptionWithControl = true;

            // 2014/4/10 修正為checked
            //col.ColumnName.Add("disabled");//決定是否check
            col.ColumnName.Add("Pledge_Id");

            col.ControlText.Add("");
            col.ControlName.Add("PledgeId");
            col.ControlValue.Add("");
            col.ControlId.Add("checkbox");
            // 2014/4/10 修正可點選checkbox
            //col.DisableValue = "0";
            col.DisableValue = "";
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("授權編號");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("授權編號");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("捐款人");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("捐款人");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("授權方式");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("授權方式");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("扣款金額");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("扣款金額");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("授權起日");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("授權起日");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("授權迄日");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("授權迄日");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("轉帳週期");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("轉帳週期");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("下次扣款日");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("下次扣款日");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("有效月年");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("有效月年");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("最後扣款日");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("最後扣款日");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            //更改資料列顏色設定
            npoGridView.ColumnNameToChangeColor = "有效年月日期";
            npoGridView.ColumnDataToChangeColor = DateTime.Parse(tbxDonateDate.Text).Year + "/" + DateTime.Parse(tbxDonateDate.Text).Month + "/01";
            lblGridList.Text = npoGridView.Render();

            //比數和金額
            Tuple<string, string> Count_Account = Account();
            lblAccount_Count.Text = Count_Account.Item1;
            lblAccount_Amount.Text = Count_Account.Item2;
            Tuple<string, string> Count_Card = Card();
            lblCard_Count.Text = Count_Card.Item1;
            lblCard_Amount.Text = Count_Card.Item2;
            Tuple<string, string> Count_ACH = ACH();
            lblACH_Count.Text = Count_ACH.Item1;
            lblACH_Amount.Text = Count_ACH.Item2;
        }
        Session["strSql"] = strSql;

    }

    //---------------------------------------------------------------------------
    protected void btnPrint_Click(object sender, EventArgs e)
    {
        string Pledge_Id = "";
        string strId = "";
        foreach (string str in Request.Form.Keys)
        {
            if (str.StartsWith("PledgeId") && str != "PledgeIdAll")
            {

                strId = str.Split('_')[1];
                Pledge_Id += strId + " ";
            }
        }

        Session["Pledge_Id"] = Pledge_Id;
        Session["Transfer_AspxName"] = ddlFormat.SelectedValue;
        //Add by GoodTV Tanya:匯出日期為畫面上的扣款日期
        Session["Donate_Date"] = tbxDonateDate.Text;
        Response.Redirect(Util.RedirectByTime("../DonateMgr/CreditCard/" + ddlFormat.SelectedValue));
    }

    protected void btnInput_Click(object sender, EventArgs e)
    {

        // 2014/4/11 少Session
        Session["List"] = lblGridList.Text;
        string strSql = "Select * From DEPT Where DeptId='" + SessionInfo.DeptID + "'";
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string LastTransfer_Date_S = DateTime.Parse(dr["LastTransfer_Date"].ToString()).ToString("yyyy/MM/dd");
        DateTime LastTransfer_Date = Convert.ToDateTime(LastTransfer_Date_S);
        DateTime DonateDate = DateTime.Parse(tbxDonateDate.Text);
        
        //2014/04/19 修改 by Ian_Kao
        string Pledge_Id = "";
        string strId = "";
        foreach (string str in Request.Form.Keys)
        {
            if (str.StartsWith("PledgeId") && str != "PledgeIdAll")
            {
                strId = str.Split('_')[1];//將勾選的Pledge_Id串成字串
                Pledge_Id += strId + " ";
            }
        }

        Session["Pledge_Id"] = Pledge_Id;
        Response.Redirect(Util.RedirectByTime("Pledge_Check.aspx", "DonateDate=" + tbxDonateDate.Text + "&LastTransfer_Check=" + HFD_LastTransfer_Check.Value + "&file=" + lblPledgeBatchFileName.Text));

    }

    protected void btnQuery_Click(object sender, EventArgs e)
    {
        HFD_Query_Flag.Value = "Y";
        LoadFormData();
    }

    //---------------------------------------------------------------------------
    private Tuple<string, string> Account()
    {
        string strSql = @"Select Count(*) as Donate_No,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY, Sum(Donate_Amt)),1),'.00','') as Donate_Amt 
                          From PLEDGE P Join DONOR D On P.Donor_Id=D.Donor_Id and D.DeleteDate is null
                          Where Status='授權中' And Post_SavingsNo<>'' And Post_AccountNo<>'' And Donate_Payment like '%帳戶轉帳授權書%'";
        strSql += SqlCondition();
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Donate_No = dr["Donate_No"].ToString();
        string Donate_Amt = dr["Donate_Amt"].ToString() == "" ? "0" : dr["Donate_Amt"].ToString();
        Tuple<string, string> Count = new Tuple<string, string>(Donate_No, Donate_Amt);
        return Count;
    }

    private Tuple<string, string> Card()
    {
        string Year = tbxDonateDate.Text == "" ? "0" : (Convert.ToDateTime(tbxDonateDate.Text).Year).ToString().Substring(2, 2);
        string Month = tbxDonateDate.Text == "" ? "0" : (Convert.ToDateTime(tbxDonateDate.Text).Month).ToString();
        string strSql = @"Select Count(*) as Donate_No, REPLACE(CONVERT(VARCHAR,CONVERT(MONEY, Sum(Donate_Amt)),1),'.00','') as Donate_Amt 
                          From PLEDGE P Join DONOR D On P.Donor_Id=D.Donor_Id and D.DeleteDate is null
                          Where Status='授權中' And Account_No<>'' And Donate_Payment like '%信用卡授權書%' And (Convert(int,RIGHT(Valid_Date,2))>" + Year +
                                " Or (Convert(int,RIGHT(Valid_Date,2))=" + Year + " And Convert(int,Left(Valid_Date,2))>=0))";
        strSql += SqlCondition();
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Donate_No = dr["Donate_No"].ToString();
        string Donate_Amt = dr["Donate_Amt"].ToString() == "" ? "0" : dr["Donate_Amt"].ToString();
        Tuple<string, string> Count = new Tuple<string, string>(Donate_No, Donate_Amt);
        return Count;
    }

    private Tuple<string, string> ACH()
    {
        string strSql = @"Select Count(*) as Donate_No, REPLACE(CONVERT(VARCHAR,CONVERT(MONEY, Sum(Donate_Amt)),1),'.00','') as Donate_Amt 
                          From PLEDGE P Join DONOR D On P.Donor_Id=D.Donor_Id and D.DeleteDate is null
                          Where Status='授權中' And Donate_Payment like '%ACH轉帳授權書%' And P_BANK<>'' And P_RCLNO<>'' And P_PID<>''";
        strSql += SqlCondition();
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Donate_No = dr["Donate_No"].ToString();
        string Donate_Amt = dr["Donate_Amt"].ToString() == "" ? "0" : dr["Donate_Amt"].ToString();
        Tuple<string, string> Count = new Tuple<string, string>(Donate_No, Donate_Amt);
        return Count;
    }

    public void Transfer_Date()
    {
        string strSql = "Select * From DEPT Where DeptId='" + SessionInfo.DeptID + "'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Transfer_Date = dr["Transfer_Date"].ToString();
        DateTime LastTransfer_Date = Convert.ToDateTime(dr["LastTransfer_Date"].ToString());
        //20140518 新增by Ian 若是第一次執行則把值丟入tbxDonateDate.Text中,之後則是查詢的值故不丟入
        if (!IsPostBack)
        {
            tbxDonateDate.Text = DateTime.Parse(LastTransfer_Date.ToString()).AddMonths(1).ToString("yyyy/MM/dd");
        }
        // 2014/4/14 預先載入前端的變數值
        DateTime DonateDate = DateTime.Parse(tbxDonateDate.Text);
        //20140518 修改by Ian 邏輯判斷修改
        //if (LastTransfer_Date.Year == DonateDate.Year && LastTransfer_Date.Month == DonateDate.Month && HFD_LastTransfer_Check.Value == "1")
        if (LastTransfer_Date.Year == DonateDate.Year && LastTransfer_Date.Month == DonateDate.Month)
        {
            HFD_LastTransfer_Check.Value = "1";
        }
        else
        {
            HFD_LastTransfer_Check.Value = "0";
        }

    }

    public string Sql()
    {
        string strSql = "";
        //20140513 修改SQL 有可能沒有信用卡的有效月年 若沒有則帶入2099/01/01以免發生Bug
        strSql = @"Select '0' as 'disabled',Pledge_Id,Pledge_Id as '授權編號',D.Donor_Name as 捐款人,Donate_Payment as 授權方式,
                          REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','')  as 扣款金額,
                          CONVERT(VarChar,Donate_FromDate,111) as 授權起日, CONVERT(VarChar,Donate_ToDate,111) as 授權迄日,Donate_Period as 轉帳週期,CONVERT(VarChar,Next_DonateDate,111) as 下次扣款日,
                          (Case When Valid_Date<>'' Then Left(Valid_Date,2)+'/'+Right(Valid_Date,2) Else '' End) as 有效月年, CONVERT(VarChar,Transfer_Date,111)  as 最後扣款日, (Case When Valid_Date<>'' Then '20' + Right(Valid_Date,2)+'/'+Left(Valid_Date,2)+'/01' Else '2099/01/01' End) as 有效年月日期
                   From PLEDGE P Join DONOR D On P.Donor_Id=D.Donor_Id and D.DeleteDate is null
                   Where Status='授權中' And ((Post_SavingsNo<>'' And Post_AccountNo<>'') Or Account_No<>'' Or P_RCLNO<>'') ";
        return strSql;
    }

    public string SqlCondition()
    {
        string strSql = "";
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " And D.dept_id='" + ddlDept.SelectedValue + "' ";
        }
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " And (D.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%' Or D.NickName like '%" + tbxDonor_Name.Text.Trim() + "%') ";
            strSql += " And (D.Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%' Or D.NickName like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%') ";
        }
        if (tbxDonateDate.Text != "")
        {
            strSql += " And Donate_FromDate <= '" + tbxDonateDate.Text + "' And Donate_ToDate>= '" + tbxDonateDate.Text +"' And ((Year(Next_DonateDate)<'" + Convert.ToDateTime(tbxDonateDate.Text).Year + "') Or (Year(Next_DonateDate)='" + Convert.ToDateTime(tbxDonateDate.Text).Year + "' And Month(Next_DonateDate)<='" + Convert.ToDateTime(tbxDonateDate.Text).Month + "')) ";
        }
        if (ddlDonate_Payment.SelectedIndex != 0)
        {
            strSql += " And Donate_Payment = '" + ddlDonate_Payment.SelectedItem.Text + "' ";
        }
        if (ddlDonate_Period.SelectedIndex != 0)
        {
            strSql += " And Donate_Period = '" + ddlDonate_Period.SelectedItem.Text + "' ";
        }
        //20151203新增下次扣款日
        if (tbxNext_DonateDateB.Text != "")
        {
            strSql += " And Next_DonateDate >= '" + tbxNext_DonateDateB.Text + "'";
        }
        if (tbxNext_DonateDateE.Text != "")
        {
            strSql += " And Next_DonateDate <= '" + tbxNext_DonateDateE.Text + "'";
        }
        return strSql;
    }

    protected void Import_Check()
    {

        lblGridList.Text = "";//先清空再填入
        var flag = false;
        //撈出Temp Table:PLEDGE_SEND_RETURN裡的Pledge_Id串成陣列 Pledge_Ids
        string strSql = "Select Pledge_Id,Donate_Amt From Pledge_Send_Return Where Return_Status='Y' ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        DataRow dr;

        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count != 0)
        {
            string[] Pledge_Ids = new string[dt.Rows.Count];
            int intCount = Convert.ToInt32(dt.Compute("Count(Pledge_Id)", "true"));
            long longSum = Convert.ToInt64(dt.Compute("Sum(Donate_Amt)", "true"));

            bool first = true;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dr = dt.Rows[i];
                Pledge_Ids[i] = dr["Pledge_Id"].ToString();
            }

            //判斷是否PLEDGE_SEND_RETURN裡的Pledge_Id有在PLEDGE中
            reload_list(true, first, Pledge_Ids);//有
            if (lblGridList.Text != "")
            {
                flag = true;
            }

            // 2014/9/16 增加授權失敗的清單
            Pledge_Err_List();

            // 2014/9/17 故決定不顯示第三層: 扣款日期前累計等待授權紀錄
            //reload_list(false, first, Pledge_Ids); //沒有

            if (flag == false)
            {
                ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"授權清單與回覆檔案資料不符！\");</script>");
            }
            else
            {
                lblPledgeBatchFileTile.Text = "匯入" + (HFD_Pledge_Import.Value == "cathayok" ? "國泰世華" : "台銀") + "回覆檔檔名：";
                lblPledgeBatchFileName.Text = FileUpload.FileName.Substring(FileUpload.FileName.LastIndexOf(@"\") + 1);
                SendEmail_Excel();
                SendEmail_Excel_OK();
                //strSql = "Select Count(*) AS Count,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,SUM(Donate_Amt)),1),'.00','') AS Sum_Amt From dbo.PLEDGE_SEND_RETURN Where Return_Status='Y'";
                //DataTable dt2 = NpoDB.GetDataTableS(strSql, null);
                //DataRow dr2 = dt2.Rows[0];
                //ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"回覆檔資料匯入成功，共" + dr2["Count"] + "筆，金額" + dr2["Sum_Amt"] + "元。\\n授權成功資料已自動勾選！\");</script>");
                ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"回覆檔資料匯入成功，共" + String.Format("{0:#,0}", intCount) + "筆，金額" + String.Format("{0:#,0}", longSum) + "元。\\n授權成功資料已自動勾選！\");</script>");
            }
        }

    }

    private void Pledge_Err_List()
    {

        string strSql = "";
        if (HFD_Pledge_Import.Value == "cathayok")
        {
            strSql = @"
            Select ROW_NUMBER() OVER(ORDER BY P.Pledge_Id) AS [序號],p.Pledge_Id as '授權編號',D.Donor_Name as 捐款人
            ,Donate_Payment as 授權方式,cast(p.Donate_Amt as int)  as 扣款金額
            ,CONVERT(VarChar,Donate_FromDate,111) as 授權起日, CONVERT(VarChar,Donate_ToDate,111) as 授權迄日
            ,Donate_Period as 轉帳週期,Authorize as [末三碼CVV]
            ,(Case When Valid_Date<>'' Then Left(Valid_Date,2)+'/'+Right(Valid_Date,2) Else '' End) as 有效月年
            ,r.[Return_Status_No] as [授權失敗碼]
			,(case when D.DeleteDate is null then r.Authorization_Messages else '捐款人已刪除' end) as [授權失敗原因],p.[Status]
            From PLEDGE P 
            Join DONOR D On P.Donor_Id=D.Donor_Id
            join Pledge_Send_Return r on p.Pledge_Id=r.Pledge_Id and r.[Return_Status] = 'N'
            order by p.Pledge_Id
         ;   ";
        }
        else
        {
            strSql = @"
            Select ROW_NUMBER() OVER(ORDER BY P.Pledge_Id) AS [序號],p.Pledge_Id as '授權編號',D.Donor_Name as 捐款人
            ,Donate_Payment as 授權方式,cast(p.Donate_Amt as int)  as 扣款金額
            ,CONVERT(VarChar,Donate_FromDate,111) as 授權起日, CONVERT(VarChar,Donate_ToDate,111) as 授權迄日
            ,Donate_Period as 轉帳週期,Authorize as [末三碼CVV]
            ,(Case When Valid_Date<>'' Then Left(Valid_Date,2)+'/'+Right(Valid_Date,2) Else '' End) as 有效月年
            ,r.[Return_Status_No] as [授權失敗碼]
			,(case when D.DeleteDate is null then s.Note else '捐款人已刪除' end) as [授權失敗原因],p.[Status]
            From PLEDGE P 
            Join DONOR D On P.Donor_Id=D.Donor_Id
            join Pledge_Send_Return r on p.Pledge_Id=r.Pledge_Id and r.[Return_Status] = 'N'
            left join Pledge_Status S on R.Return_Status_No = S.ErrorCode
            order by p.Pledge_Id
         ;   ";
        }
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count > 0)
        {

            //lblGridList.ForeColor = System.Drawing.Color.Black;
            string strBody = "<table class='table_h' width='100%'>";
            strBody += @"<TR><TH noWrap><SPAN>序號</SPAN></TH>
                            <TH noWrap><SPAN>授權編號</SPAN></TH>
                            <TH noWrap><SPAN>捐款人</SPAN></TH>
                            <TH noWrap><SPAN>授權方式</SPAN></TH>
                            <TH noWrap><SPAN>扣款金額</SPAN></TH>
                            <TH noWrap><SPAN>授權起日</SPAN></TH>
                            <TH noWrap><SPAN>授權迄日</SPAN></TH>
                            <TH noWrap><SPAN>轉帳週期</SPAN></TH>
                            <TH noWrap><SPAN>有效月年</SPAN></TH>
                            <TH noWrap><SPAN>末三碼CVV</SPAN></TH>
                            <TH noWrap><SPAN>失敗碼</SPAN></TH>
                            <TH noWrap><SPAN>授權失敗原因</SPAN></TH>
                        </TR>";
            foreach (DataRow dr in dt.Rows)
            {

                strBody += "<TR " + (Convert.ToInt64(dr["扣款金額"].ToString()) >= 10000 ? "style=\"COLOR: white; BACKGROUND-COLOR: darkorange\" " : "") + "align='center'><TD noWrap align='right'><SPAN>" + dr["序號"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["授權編號"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["捐款人"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["授權方式"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD noWrap align='right'><SPAN>" + String.Format("{0:#,0}", Convert.ToInt64(dr["扣款金額"].ToString())) + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["授權起日"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["授權迄日"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["轉帳週期"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["有效月年"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["末三碼CVV"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["授權失敗碼"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap align='left'><SPAN" + (dr["授權失敗碼"].ToString() == "05" && Convert.ToInt64(dr["扣款金額"].ToString()) < 10000 ? " style=\"COLOR: red\"" : "") + ">" + dr["授權失敗原因"].ToString() + "</SPAN></TD></TR>";

            }

            strBody += "</table>";
            lblGridList.Text += "<BR/><TABLE cellSpacing='0' width='100%' border=0 height='30px'><TR style='BACKGROUND-COLOR: mistyrose'>" +
            "<TD width='5%' align=right><DIV style='HEIGHT: 15px; WIDTH: 20px; BACKGROUND-COLOR: darkorange'></DIV></TD><TD width='25%' align=left><SPAN style='FONT-SIZE: 20px; FONT-WEIGHT: bold; COLOR: red;'>大額警示</SPAN></TD>" +
            "<TD align=center><SPAN style='FONT-SIZE: 18px; COLOR: blue;'>本次匯入授權失敗</SPAN></TD><TD width='25%'></TD><TD width='5%'></TD></TR></TABLE>" + strBody;
            //Session["strSql"] = strSql;
        }
    }

    private void reload_list(bool check, bool first, string[] pledge_id)
    {
        string strSql = Sql();
        strSql += SqlCondition();
        if (check == true)
        {
            strSql += " AND Pledge_Id IN ( ";
        }
        else
        {
            strSql += " AND Pledge_Id NOT IN ( ";
        }
        for (int i = 0; i < pledge_id.Length; i++)
        {
            if (i == pledge_id.Length - 1)
            {
                strSql += "'" + pledge_id[i] + "'";
            }
            else
            {
                strSql += "'" + pledge_id[i] + "',";
            }
        }
        strSql += " ) ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (check == true && dt.Rows.Count == 0)
        {
            lblGridList.Text = "";
            return;
        }
        else if (check == false && dt.Rows.Count == 0)
        {
            return;
        }
        //Grid initial
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("Pledge_Id");
        npoGridView.DisableColumn.Add("Pledge_Id");
        npoGridView.DisableColumn.Add("disabled");
        npoGridView.ShowPage = false;
        //-------------------------------------------------------------------------
        // 2014/4/11 修正標題換成為checkbox
        //NPOGridViewColumn col = new NPOGridViewColumn("選擇");
        NPOGridViewColumn col = new NPOGridViewColumn();
        col.ColumnType = NPOColumnType.Checkbox;
        col.ControlKeyColumn.Add("Pledge_Id");
        if (check == true)//決定是否check
        {
            col.ColumnName.Add("Pledge_Id");
        }
        else
        {
            col.ColumnName.Add("disabled");
        }
        col.ControlText.Add("");
        col.ControlName.Add("PledgeId");
        col.ControlValue.Add("");
        col.ControlId.Add("checkbox");
        if (first == false)//第一次顯示標頭,之後不顯示,但是width會亂掉
        {
           // col.ShowTitle = false;
        }
        // 2014/4/10 修正可點選checkbox
        //col.DisableValue = "0";
        col.DisableValue = "";
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權編號");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權編號");
        if (first == false)
        {
           // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("捐款人");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("捐款人");
        if (first == false)
        {
           // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權方式");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權方式");
        if (first == false)
        {
           // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("扣款金額");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("扣款金額");
        if (first == false)
        {
           // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權起日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權起日");
        if (first == false)
        {
           // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權迄日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權迄日");
        if (first == false)
        {
           // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("轉帳週期");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("轉帳週期");
        if (first == false)
        {
           // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("下次扣款日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("下次扣款日");
        if (first == false)
        {
           // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("有效月年");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("有效月年");
        if (first == false)
        {
           // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("最後扣款日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("最後扣款日");
        if (first == false)
        {
           // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        //更改資料列顏色設定
        npoGridView.ColumnNameToChangeColor = "有效年月日期";
        npoGridView.ColumnDataToChangeColor = DateTime.Parse(tbxDonateDate.Text).Year + "/" + DateTime.Parse(tbxDonateDate.Text).Month + "/01";
        lblGridList.Text += (check == true ? "<BR/><DIV style=\"FONT-SIZE: 18px; COLOR: blue; LINE-HEIGHT: 30px; BACKGROUND-COLOR: skyblue\">本次匯入授權成功</DIV>" : "<BR/><DIV style=\"FONT-SIZE: 18px; COLOR: blue; LINE-HEIGHT: 30px; BACKGROUND-COLOR: palegreen\">扣款日期前累計等待授權紀錄</DIV>") + npoGridView.Render();
        Session["List"] = lblGridList.Text;
    }

    private void SendEmail_Excel()
    {

        string strSql = "";
        if (HFD_Pledge_Import.Value == "cathayok")
        {
            strSql = @"select ROW_NUMBER() OVER(ORDER BY P.Pledge_Id) AS '序號'
                        ,P.Pledge_Id as '授權編號',CAST(SR.Donate_Amt as int) as '授權金額'
                        ,D.Donor_Id as '捐款人編號', D.Donor_Name as '捐款人姓名'
						,p.[Card_Bank] as [發卡銀行],p.[Account_No] as [信用卡卡號],p.[Valid_Date] as [效期],Authorize as [末三碼CVV]
                        ,IsNull(D.ZipCode,'') AS '郵遞區號'
                        ,Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') 
		                        else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
                        ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
			                        (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
			                        Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
		                        Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
			                        Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話'  
                        ,IsNull(Cellular_Phone,'') AS '手機', SR.Return_Status_No as '授權失敗碼' 
                        ,(case when D.DeleteDate is null then SR.Authorization_Messages else '捐款人已刪除' end) as '授權失敗原因'
                        from Pledge_Send_Return SR
                        left join Pledge P on SR.Pledge_Id = P.Pledge_Id
                        left join Donor D on P.Donor_Id = D.Donor_Id
                        LEFT JOIN dbo.CODECITYNew C1 ON D.City = C1.ZipCode
                        LEFT JOIN dbo.CODECITYNew C2 ON D.Area = C2.ZipCode
                        where SR.Return_Status='N'
                        order by SR.Pledge_Id";
        }
        else
        {
            strSql = @"select ROW_NUMBER() OVER(ORDER BY P.Pledge_Id) AS '序號'
                        ,P.Pledge_Id as '授權編號',CAST(SR.Donate_Amt as int) as '授權金額'
                        ,D.Donor_Id as '捐款人編號', D.Donor_Name as '捐款人姓名'
						,p.[Card_Bank] as [發卡銀行],p.[Account_No] as [信用卡卡號],p.[Valid_Date] as [效期],Authorize as [末三碼CVV]
                        ,IsNull(D.ZipCode,'') AS '郵遞區號'
                        ,Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') 
		                        else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
                        ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
			                        (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
			                        Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
		                        Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
			                        Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話'  
                        ,IsNull(Cellular_Phone,'') AS '手機', PS.ErrorCode as '授權失敗碼' 
                        ,(case when D.DeleteDate is null then PS.Note else '捐款人已刪除' end) as '授權失敗原因'
                        from Pledge_Send_Return SR
                        left join Pledge P on SR.Pledge_Id = P.Pledge_Id
                        left join Donor D on P.Donor_Id = D.Donor_Id
                        LEFT JOIN dbo.CODECITYNew C1 ON D.City = C1.ZipCode
                        LEFT JOIN dbo.CODECITYNew C2 ON D.Area = C2.ZipCode
                        left join Pledge_Status PS on SR.Return_Status_No = PS.ErrorCode
                        where SR.Return_Status='N'
                        order by SR.Pledge_Id";
        }
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            //ShowSysMsg("查無資料!!!");
            return;
        }

        //string style = "<style> .text { mso-number-format:\\@; } </style> ";
        string style = "<style type=text/css>td{mso-number-format:\"\\@\";}</style>";
        GridList.Text = style + GetTitle(dt.Rows[0], false) + GetTable1(dt.Rows[0], strSql, false);
        long longSum = Convert.ToInt64(dt.Compute("Sum([授權金額])", "true"));
        byte[] byteArray = Encoding.UTF8.GetBytes(GridList.Text);
        MemoryStream sm = new MemoryStream(byteArray);

        string StrEmailToDonations = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];
        //發送內部通知mail*****************************************
        string strBody = "親愛的同工平安,<BR/><BR/>";
        //"2014-01-01"
        strBody += "以下是本次扣款日期: " + tbxDonateDate.Text + " 回覆之授權失敗記錄 <BR/><BR/>";
        strBody += "<font color='blue'>合計授權失敗的筆數為：" + String.Format("{0:#,0}", dt.Rows.Count) + " 筆，金額為: " + String.Format("{0:#,0}", longSum) + " 元 </font><BR/><BR/>";
        strBody += GridList.Text;
        strBody += "<BR/>請盡速追蹤處理!<BR/><BR/>";

        SendEMailObject MailObject = new SendEMailObject();
        string MailSubject = "信用卡授權失敗提示(批次) - 扣款日期:" + tbxDonateDate.Text;
        string MailBody = strBody;
        string strFilename = "PledgeReturnError.xls";
        string result = MailObject.SendEmailAttachment(StrEmailToDonations, strFilename, MailSubject, strBody, sm);

        if (MailObject.ErrorCode != 0)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        }

    }

    //---------------------------------------------------------------------------
    private string GetTitle(DataRow dr, bool OKERR)
    {
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "15px");
        css.Add("font-family", "標楷體");
        css.Add("width", TableWidth);

        string strFontSize = "24px";

        string strTitle = OKERR ? (HFD_Pledge_Import.Value == "cathayok" ? "國泰世華" : "台銀") + "授權成功回覆檔暨捐款人資料<br/> " : (HFD_Pledge_Import.Value == "cathayok" ? "國泰世華" : "台銀") + "授權失敗回覆檔暨捐款人資料<br/> ";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = OKERR ? 12 : 17;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }

    //---------------------------------------------------------------------------
    private string GetTable1(DataRow dr, string strSql, bool OKERR)
    {
        //組 table
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        Dictionary<string, object> dict = new Dictionary<string, object>();

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;

        DataTable dtRet = OKERR ? CaseUtil.PledgeReturnOK_Print(dt) : CaseUtil.PledgeReturnError_Print(dt);

        table.Border = 1;
        table.CellPadding = 3;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "15px");
        css.Add("font-family", "新細明體");
        css.Add("width", TableWidth);
        css.Add("line-height", "25px");

        row = new HtmlTableRow();

        int iCtrl = 0;
        foreach (DataColumn dc in dtRet.Columns)
        {
            cell = new HtmlTableCell();
            cell.Style.Add("background-color", " #FFE4C4 ");
            /*
            cell.Style.Add("border-left", ".5pt solid windowtext");
            cell.Style.Add("border-top", ".5pt solid windowtext");
            cell.Style.Add("border-right", ".5pt solid windowtext");
            cell.Style.Add("border-bottom", ".5pt solid windowtext");
            */

            if (iCtrl == 0)
            {
                cell.Width = "30";
            }
            else
            {
                cell.Width = "90mm";
            }
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp;" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

        foreach (DataRow drow in dtRet.Rows)
        {

            row = new HtmlTableRow();
            iCtrl = 0;
            foreach (object objItem in drow.ItemArray)
            {
                cell = new HtmlTableCell();
                /*
                cell.Style.Add("border-left", ".5pt solid windowtext");
                cell.Style.Add("border-top", ".5pt solid windowtext");
                cell.Style.Add("border-right", ".5pt solid windowtext");
                cell.Style.Add("border-bottom", ".5pt solid windowtext");
                */

                if (iCtrl == 2)
                {
                    cell.InnerHtml = objItem.ToString() == "" ? "&nbsp;" : String.Format("{0:#,0}", Convert.ToInt64(objItem.ToString()));
                }
                else
                {
                    cell.InnerHtml = objItem.ToString() == "" ? "&nbsp;" : objItem.ToString();
                }
                //cell.InnerHtml = objItem.ToString() == "" ? "&nbsp;" : objItem.ToString();
                row.Cells.Add(cell);
                iCtrl++;
            }
            table.Rows.Add(row);
        }

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }

    private void SendEmail_Excel_OK()
    {

        string strSql = "";
        if (HFD_Pledge_Import.Value == "cathayok")
        {
            strSql = @"select ROW_NUMBER() OVER(ORDER BY P.Pledge_Id) AS '序號'
                        ,P.Pledge_Id as '授權編號',CAST(SR.Donate_Amt as int) as '授權金額'
                        ,D.Donor_Id as '捐款人編號', D.Donor_Name as '捐款人姓名'
						,p.[Card_Bank] as [發卡銀行],p.[Account_No] as [信用卡卡號],p.[Valid_Date] as [效期]
                        ,IsNull(D.ZipCode,'') AS '郵遞區號'  
                        ,Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') 
		                        else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
                        ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
			                        (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
			                        Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
		                        Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
			                        Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話'  
                        ,IsNull(Cellular_Phone,'') AS '手機'
                        from Pledge_Send_Return SR
                        left join Pledge P on SR.Pledge_Id = P.Pledge_Id
                        join Donor D on P.Donor_Id = D.Donor_Id and D.DeleteDate is null
                        LEFT JOIN dbo.CODECITYNew C1 ON D.City = C1.ZipCode
                        LEFT JOIN dbo.CODECITYNew C2 ON D.Area = C2.ZipCode
                        where SR.Return_Status='Y'
                        order by SR.Pledge_Id";
        }
        else
        {
            strSql = @"select ROW_NUMBER() OVER(ORDER BY P.Pledge_Id) AS '序號'
                        ,P.Pledge_Id as '授權編號',CAST(SR.Donate_Amt as int) as '授權金額'
                        ,D.Donor_Id as '捐款人編號', D.Donor_Name as '捐款人姓名'
						,p.[Card_Bank] as [發卡銀行],p.[Account_No] as [信用卡卡號],p.[Valid_Date] as [效期]
                        ,IsNull(D.ZipCode,'') AS '郵遞區號'  
                        ,Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') 
		                        else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
                        ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
			                        (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
			                        Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
		                        Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
			                        Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話'  
                        ,IsNull(Cellular_Phone,'') AS '手機'
                        from Pledge_Send_Return SR
                        left join Pledge P on SR.Pledge_Id = P.Pledge_Id
                        join Donor D on P.Donor_Id = D.Donor_Id and D.DeleteDate is null
                        LEFT JOIN dbo.CODECITYNew C1 ON D.City = C1.ZipCode
                        LEFT JOIN dbo.CODECITYNew C2 ON D.Area = C2.ZipCode
                        left join Pledge_Status PS on SR.Return_Status_No = PS.ErrorCode
                        where SR.Return_Status='Y'
                        order by SR.Pledge_Id";
        }
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            //ShowSysMsg("查無資料!!!");
            return;
        }
       
        //string style = "<style> .text { mso-number-format:\\@; } </style> ";
        string style = "<style type=text/css>td{mso-number-format:\"\\@\";}</style>";
        string strReplaceOldValue = GetTable1(dt.Rows[0], strSql, true);
        long longSum = Convert.ToInt64(dt.Compute("Sum([授權金額])", "true"));
        string strReplaceNewValue = "<tr><td colspan='2'>合計</td><td>" + String.Format("{0:#,0}", longSum) + "</td><td colspan='9'>&nbsp;</td></tr></table>";
        GridList.Text = style + GetTitle(dt.Rows[0], true) + strReplaceOldValue.Replace("</table>", strReplaceNewValue);
        byte[] byteArray = Encoding.UTF8.GetBytes(GridList.Text);
        MemoryStream sm = new MemoryStream(byteArray);

        string StrEmailToDonations = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];
        //發送內部通知mail*****************************************
        string strBody = "親愛的同工平安,<BR/><BR/>";
        //"2014-01-01"
        strBody += "以下是本次扣款日期: " + tbxDonateDate.Text + " 回覆之授權成功記錄 <BR/>";
        strBody += "<font color='blue'>合計授權成功的筆數為：" + String.Format("{0:#,0}", dt.Rows.Count) + " 筆，金額為： " + String.Format("{0:#,0}", longSum) + " 元 </font><BR/><BR/>";
        strBody += GridList.Text;
        strBody += "<BR/>請盡速進行轉捐款收據作業!<BR/><BR/>";

        SendEMailObject MailObject = new SendEMailObject();
        string MailSubject = "信用卡授權成功通知(批次) - 扣款日期:" + tbxDonateDate.Text;
        string MailBody = strBody;
        string strFilename = "PledgeReturnOK.xls";
        string result = MailObject.SendEmailAttachment(StrEmailToDonations, strFilename, MailSubject, strBody, sm);

        if (MailObject.ErrorCode != 0)
        {
            ClientScript.RegisterStartupScript(this.GetType(),"s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        }

    }

    //信用卡批次授權(一般)回覆檔匯入
    protected void btnFileImport_Click(object sender, EventArgs e)
    {

        if (!FileUpload.HasFile)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript'>alert('請選擇檔案');</script>");
            return;
        }
        if (FileNameValid() == false)
            return;
        if (!AllowSave())
            return;

        string txt_filePath = "";
        lblPledgeBatchFileTile.Text = "";
        lblPledgeBatchFileName.Text = "";

        txt_filePath = SaveFileAndReturnPath();//先上傳TXT檔案給Server
        if (HFD_Pledge_Import.Value == "cathay")
        {
            //讀取國泰世華的TXT中的資料寫入DB
            SaveOrInsertCathayDB();
        }
        else if (HFD_Pledge_Import.Value == "bot")
        {
            //讀取台銀世華的TXT中的資料寫入DB
            SaveOrInsertDB();
        }

    }

    //---------------------------------------------------------------------------
    public bool FileNameValid()
    {
        string strfilePath = FileUpload.FileName;
        string strfileName = strfilePath.Substring(strfilePath.LastIndexOf(@"\") + 1);
        if (strfileName.Contains(" ") || strfileName.Contains("　"))
        {
            ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('檔案名稱不能有空白');</script>");
            return false;
        }
        return true;
    }

    //檢查及限定上傳檔案類型與大小
    public bool AllowSave()
    {
        int DenyMbSize = 4;//限制大小30MB
        bool flag = false;
        //判斷副檔名
        switch (System.IO.Path.GetExtension(FileUpload.PostedFile.FileName).ToUpper())
        {
            case ".TXT":
                flag = true;
                break;
            case ".01R":
                flag = true;
                break;
            case ".02R":
                flag = true;
                break;
            case ".03R":
                flag = true;
                break;
            case ".04R":
                flag = true;
                break;
            case ".05R":
                flag = true;
                break;
            case ".06R":
                flag = true;
                break;
            case ".07R":
                flag = true;
                break;
            case ".08R":
                flag = true;
                break;
            case ".09R":
                flag = true;
                break;
            default:
                ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('此檔案類型不允許上傳');</script>");
                flag = false;
                break;
        }
        if (Util.GetFileSizeMB(FileUpload.FileBytes.Length) > DenyMbSize)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('檔案大小不得超過" + DenyMbSize + @"MB');</script>");
            flag = false;
        }
        return flag;
    }

    //儲存TXT檔案給Server
    private string SaveFileAndReturnPath()
    {
        string return_file_path = "";//上傳的Excel檔在Server上的位置
        if (FileUpload.FileName != "")
        {
            String appPath = Request.PhysicalApplicationPath;
            return_file_path = System.IO.Path.Combine(appPath + System.Configuration.ConfigurationManager.AppSettings["UploadPath"], Guid.NewGuid().ToString() + ".txt");

            FileUpload.SaveAs(return_file_path);
        }
        return return_file_path;
    }

    //把TXT資料Insert into Table
    private void SaveOrInsertDB()
    {
        //先將PLEDGE_SEND_RETURN中資料刪除
        string strSql = "Delete from Pledge_Send_Return where Pledge_Id>0";
        NpoDB.ExecuteSQLS(strSql, null);
        //讀取TXT檔
        bool flag = false;
        System.IO.StreamReader red = new System.IO.StreamReader(this.FileUpload.PostedFile.InputStream, System.Text.Encoding.Default);
        long SumAmt = 0;
        long Total_Amt = 0;
        string Donate_Date_yyyyMMDD = "";
        int ErrorCut = 0;

        while (red.Peek() > 0)
        {
            string strContent = red.ReadLine();

            //防呆 2015/3/20
            if (strContent.IndexOf("T", 0) == -1)
            {
                if (strContent.Length >= 84)
                {
                    string Store_Id = strContent.Substring(0, 19).Trim();
                    if (Store_Id != "0048158205490018398")
                    {
                        ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('台銀回覆檔格式有問題！！');</script>");
                        return;
                    }
                }
                else
                {
                    ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('台銀回覆檔格式有問題！！');</script>");
                    return;
                }
            }

            //基本資料
            /* 沒用到
            string SourcePath = "";
            string SourceName = "";
            string UploadName = "";
            string UploadSize = "";
            string ExtName = "";
            if (strContent.Trim() != "")
            {
                UploadName = Guid.NewGuid().ToString() + ".txt";
                SourcePath = FileUpload.PostedFile.FileName.Replace(UploadName, "");
                SourceName = FileUpload.FileName;
                UploadSize = string.Format((FileUpload.PostedFile.ContentLength / 1024).ToString(), "0.00");
                ExtName = System.IO.Path.GetExtension(FileUpload.PostedFile.FileName).ToUpper();
            }
            */

            //Data
            string Account_No = "";
            string Donate_Amt = "";
            string Return_Status = "";
            string Pledge_Id = "";
            if (strContent.IndexOf("T", 0) == -1)
            {

                Account_No = (Mid(strContent, 19, 19).Trim());
                Donate_Amt = Trim0(Mid(strContent, 45, 10));
                Return_Status = (Mid(strContent, 57, 2).Trim());
                Pledge_Id = Trim0(Mid(strContent, 76, 8));

                SumAmt = SumAmt + Convert.ToInt64(Donate_Amt);
                //insert Data
                strSql = "insert into  Pledge_Send_Return\n";
                strSql += "( Pledge_Id, Account_No, Return_Status, Donate_Amt,Return_Status_No ) values\n";
                strSql += "( @Pledge_Id,@Account_No,@Return_Status,@Donate_Amt,@Return_Status_No ) ; ";

                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Pledge_Id", Pledge_Id);
                dict.Add("Account_No", Account_No);
                if (Return_Status == "00")
                {
                    dict.Add("Return_Status", "Y");
                }
                else
                {
                    dict.Add("Return_Status", "N");
                }
                dict.Add("Donate_Amt", Donate_Amt);
                dict.Add("Return_Status_No", Return_Status);
                NpoDB.ExecuteSQLS(strSql, dict);
                flag = true;

            }
            else
            {
                Total_Amt = Convert.ToInt64(Trim0(Mid(strContent, 7, 10)));
                Donate_Date_yyyyMMDD = Mid(strContent, 19, 8);
            }

        }

        string Donate_Year = Convert.ToDateTime(tbxDonateDate.Text).Year.ToString();
        string Donate_Month = Convert.ToDateTime(tbxDonateDate.Text).Month.ToString();
        if (Donate_Month.Length == 1)
        {
            Donate_Month = "0" + Donate_Month;
        }
        string Donate_Day = Convert.ToDateTime(tbxDonateDate.Text).Day.ToString();
        if (Donate_Day.Length == 1)
        {
            Donate_Day = "0" + Donate_Day;
        }

        if (SumAmt != Total_Amt || ErrorCut > 0)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('台銀回覆檔格式有問題！！');</script>");
        }
        else if ((Donate_Year + Donate_Month + Donate_Day) != Donate_Date_yyyyMMDD)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('台銀回覆檔不屬於目前的扣款日期!！');</script>");
        }
        else
        {
            if (flag == true)
            {

                HFD_Pledge_Import.Value = "botok";
                Import_Check();

            }
            else
            {
                ShowSysMsg("無符合的資料！");
            }
        }
    }

    //增加國泰世華
    //把TXT資料Insert into Table
    private void SaveOrInsertCathayDB()
    {
        //先將PLEDGE_SEND_RETURN中資料刪除
        string strSql = "Delete from Pledge_Send_Return where Pledge_Id>0";
        NpoDB.ExecuteSQLS(strSql, null);
        //讀取TXT檔
        bool flag = false;
        System.IO.StreamReader red = new System.IO.StreamReader(this.FileUpload.PostedFile.InputStream, System.Text.Encoding.Default);
        string Donate_Date_yyyyMMDD = "";
        int intErrorCount = 0;

        while (red.Peek() > 0)
        {
            string strContent = red.ReadLine();

            //防呆 2015/3/20
            if (strContent.Length >= 82)
            {
                //商家代碼(目前代碼有9位)  15位	靠左，右補白
                string Store_Id = strContent.Substring(0, 9);
                if (Store_Id != "010970018")
                {
                    ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('國泰世華回覆檔格式有問題！！');</script>");
                    return;
                }
            }
            else
            {
                ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('國泰世華回覆檔格式有問題！！');</script>");
                return;
            }

            //Data
            //商家代碼  15位	靠左，右補白
            //不儲存，只做前面的防呆檢查
            //交易上傳日期	  10位	YYYY/MM/DD
            string Upload_Date = strContent.Substring(15, 10);
            //交易上傳批次編號	  8 位	靠左右補白同一天上傳第幾次的流水編號
            string Upload_Batch_Number = strContent.Substring(25, 8).Trim();
            //訂單編號	  20位	靠右左補白
            //order_no使用Account_No欄位，可省一個欄位
            string Account_No = strContent.Substring(33, 20).Trim();
            string Pledge_Id = Trim0(Mid(Account_No, 8, 8));
            Donate_Date_yyyyMMDD = Account_No.Substring(0, 8);
            //授權時間	14位	為YYYYMMDDhhmmss
            string Authorization_DateTime = strContent.Substring(53, 14);
            //授權金額	      	  8 位	靠右左補0
            long Donate_Amt = Convert.ToInt64(strContent.Substring(67, 8));
            //授權是否成功			1位	N:為失敗   Y:為成功
            string Return_Status = strContent.Substring(75, 1);
            //授權碼			    	6位	若授權失敗 則為000000
            string Return_Status_No = strContent.Substring(76, 6);
            //錯誤訊息				靠左右補白
            string Authorization_Messages = strContent.Substring(82).Trim();

            //防呆(授權是否成功)
            if (Return_Status != "Y" && Return_Status != "N") intErrorCount++;

            //insert Data
            strSql = "insert into  Pledge_Send_Return ";
            strSql += "( Pledge_Id, Account_No, Donate_Amt, Return_Status, Return_Status_No, Upload_Date, Upload_Batch_Number, Authorization_DateTime, Authorization_Messages ) values ";
            strSql += "(@Pledge_Id,@Account_No,@Donate_Amt,@Return_Status,@Return_Status_No,@Upload_Date,@Upload_Batch_Number,@Authorization_DateTime,@Authorization_Messages ) ; ";

            Dictionary<string, object> dict = new Dictionary<string, object>();
            dict.Add("Pledge_Id", Pledge_Id);
            dict.Add("Account_No", Account_No);
            dict.Add("Donate_Amt", Donate_Amt);
            dict.Add("Return_Status", Return_Status);
            dict.Add("Return_Status_No", Return_Status_No);
            dict.Add("Upload_Date", Upload_Date);
            dict.Add("Upload_Batch_Number", Upload_Batch_Number);
            dict.Add("Authorization_DateTime", Authorization_DateTime);
            dict.Add("Authorization_Messages", Authorization_Messages);
            NpoDB.ExecuteSQLS(strSql, dict);
            flag = true;

        }

        string Donate_Year = Convert.ToDateTime(tbxDonateDate.Text).Year.ToString();
        string Donate_Month = Convert.ToDateTime(tbxDonateDate.Text).Month.ToString();
        if (Donate_Month.Length == 1)
        {
            Donate_Month = "0" + Donate_Month;
        }
        string Donate_Day = Convert.ToDateTime(tbxDonateDate.Text).Day.ToString();
        if (Donate_Day.Length == 1)
        {
            Donate_Day = "0" + Donate_Day;
        }

        if (intErrorCount > 0)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('國泰世華回覆檔格式有問題！');</script>");
        }
        else if ((Donate_Year + Donate_Month + Donate_Day) != Donate_Date_yyyyMMDD)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('國泰世華回覆檔不屬於目前的扣款日期！');</script>");
        }
        else
        {
            if (flag == true)
            {

                HFD_Pledge_Import.Value = "cathayok";
                Import_Check();

            }
            else
            {
                ShowSysMsg("無符合的資料！");
            }
        }
    }

    // 取得起始位置至結束之字串
    public static string Mid(string sSource, int iStart, int iLength)
    {
        if (sSource.Trim().Length > 0)
        {
            int iStartPoint = iStart > sSource.Length ? sSource.Length : iStart;
            return sSource.Substring(iStartPoint, iStartPoint + iLength > sSource.Length ? sSource.Length - iStartPoint : iLength);
        }
        else
            return "";
    }

    public static string Trim0(string sSource)
    {
        int r = 0;
        for (int i = 0; i < sSource.Length; i++)
        {
            if (sSource[i].ToString().Equals("0"))
            {
                r += 1;
            }
            else
            {
                break;
            }
        }
        return sSource.Substring(r);
    }  


}
