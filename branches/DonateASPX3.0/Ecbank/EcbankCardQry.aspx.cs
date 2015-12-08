using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Ecbank_EcbankCardQry : BasePage
{
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
        Session["ProgID"] = "EcbankCardQry";
        //權控處理
        AuthrityControl();

        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            LoadDropDownListData();
            tbxDonateCreateDateS.Text = DateTime.Now.Year.ToString() + "/" + DateTime.Now.AddMonths(-1).Month.ToString() + "/1";
            LoadFormData();
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
        Authrity.CheckButtonRight("_Print", btnToxls);
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.SelectedIndex = 0;

        //捐款用途
        Util.FillDropDownList(ddlDonate_Purpose, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseName", false);
        ddlDonate_Purpose.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Purpose.SelectedIndex = 0;

        //收據開立
        Util.FillDropDownList(ddlInvoice_Type, Util.GetDataTable("CaseCode", "GroupName", "收據開立", "", ""), "CaseName", "CaseName", false);
        ddlInvoice_Type.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Type.SelectedIndex = 0;

        //交易狀態
        ddlClose_Type.Items.Insert(0, new ListItem("", ""));
        ddlClose_Type.Items.Insert(1, new ListItem("成功", "L"));
        ddlClose_Type.Items.Insert(2, new ListItem("失敗", "S"));
        ddlClose_Type.Items.Insert(3, new ListItem("已請款", "P"));
        ddlClose_Type.SelectedIndex = 0;

        //是否匯出
        ddlDonate_Export.Items.Insert(0, new ListItem("", ""));
        ddlDonate_Export.Items.Insert(1, new ListItem("未匯出", "N"));
        ddlDonate_Export.Items.Insert(2, new ListItem("已匯出", "Y"));
        ddlDonate_Export.SelectedIndex = 1;

        

    }
    public void LoadFormData()
    {
        string strSql_temp;
        DataTable dt;
        strSql_temp= @"Select A.Ser_No as Ser_No,A.Donate_DonorName as 捐款人 ,'信用卡' as '交易方式',A.Donate_CreateDateTime as 交易日期,A.od_sob as 交易序號
                          ,(Case When CONVERT(numeric,B.amount)>0 Then CONVERT(numeric,B.amount) Else CONVERT(numeric,A.Donate_Amount) End) as 交易金額
                          ,(Case When B.succ='1' Then '成功' Else '失敗' End) as 交易狀態
                          ,B.response_msg as 授權狀態,B.auth_code as 授權碼,close_type as 觸發請款
                   From DONATE_WEB A Left Join DONATE_ECPAY As B On A.od_sob=B.od_sob 
                   Where Donate_Type='creditcard'";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = Condition(strSql_temp);
        dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.DisableColumn.Add("Ser_No");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            lblGridList.Text = npoGridView.Render();
        }
    }
    //---------------------------------------------------------------------------
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        string strSql_tmep = @"Select A.Ser_No,A.Donate_DonorName as 捐款人,'信用卡' as 交易方式,A.Donate_CreateDateTime as 交易日期
                                      ,A.od_sob as 交易序號,(Case When CONVERT(numeric,D.amount)>0 Then CONVERT(numeric,D.amount) Else CONVERT(numeric,A.Donate_Amount) End) as 交易金額
                                      ,(Case When D.succ='1' Then '成功' Else '失敗' End) as 交易狀態,D.response_msg as 授權狀態,D.auth_code as 授權碼
                                      ,(Case When D.Close_Type='ok' Then 'V' Else '' End) as 是否請款,Donate_CellPhone as 手機,Donate_TelOffice as 聯絡電話
                                      ,Donate_Email as 電子郵件,Donate_Purpose as 捐款用途,Donate_Invoice_Type as 收據開立,Donate_Invoice_Title as 收據抬頭
                                      ,(Case When A.Donate_Invoice_CityCode='' Then Donate_Invoice_Address Else Case When B.mValue<>C.mValue Then B.mValue+A.Donate_Invoice_ZipCode+C.mValue+A.Donate_Invoice_Address Else B.mValue+A.Donate_Invoice_ZipCode+Donate_Invoice_Address End End) as 收據地址
                                      ,(Case When A.Donate_CityCode='' Then Donate_Address Else Case When F.mValue<>G.mValue Then F.mValue+A.Donate_ZipCode+G.mValue+A.Donate_Address Else F.mValue+A.Donate_ZipCode+Donate_Address End End) as 通訊地址
                                      ,Donate_Sex as 性別,Donate_IDNO as 身分證,Donate_Birthday as 出生日期,Donate_Education as 教育程度
                                      ,Donate_Occupation as 職業 ,Donate_Export 
                               From DONATE_WEB A 
                                    Left Join CODECITY As B On A.Donate_Invoice_CityCode=B.mCode
                                    Left Join CODECITY As C On A.Donate_Invoice_AreaCode=C.mCode 
                                    Left Join CODECITY As F On A.Donate_CityCode=F.mCode 
                                    Left Join CODECITY As G On A.Donate_AreaCode=G.mCode
                                    Left Join DONATE_ECPAY As D On A.od_sob=D.od_sob 
                               Where A.Donate_Type='creditcard' ";
        Session["strSql"] = Condition(strSql_tmep);
        Response.Redirect("EcbankCardQry_Print_Excel.aspx");
    }
    private string Condition(string strSql)
    {
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and Dept_Id = '" + ddlDept.SelectedValue + "'";
        }
        if (tbxDonor_Name.Text.Trim() != "")
        {
            strSql += " and Donate_DonorName like '%" + tbxDonor_Name.Text.Trim() + "%'";
        }
        if (ddlDonate_Purpose.SelectedIndex != 0)
        {
            strSql += " and Donate_Purpose = '" + ddlDonate_Purpose.SelectedItem.Text + "'";
        }
        if (ddlInvoice_Type.SelectedIndex != 0)
        {
            strSql += " and Donate_Invoice_Type = '" + ddlInvoice_Type.SelectedItem.Text + "'";
        }
        if (tbxDonateCreateDateS.Text.Trim() != "")
        {
            strSql += "and Donate_CreateDate >='" + Util.DateTime2String(tbxDonateCreateDateS.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "'";
        }
        if (tbxDonateCreateDateE.Text.Trim() != "")
        {
            strSql += "and Donate_CreateDate <='" + Util.DateTime2String(tbxDonateCreateDateE.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty) + "'";
        }
        if (ddlClose_Type.SelectedIndex != 0)
        {
            if (ddlClose_Type.SelectedValue == "L")
            {
                strSql += " And (B.succ='0' Or B.succ Is Null) ";
            }
            else if (ddlClose_Type.SelectedValue == "S")
            {
                strSql += " And (B.succ='1') ";
            }
            else if (ddlClose_Type.SelectedValue == "P")
            {
                strSql += " And (B.Close_Type='ok') ";
            }
        }
        if (ddlDonate_Export.SelectedIndex != 0)
        {
            strSql += " and Donate_Export = '" + ddlDonate_Export.SelectedItem.Text + "'";
        }
        return strSql;
    }
}