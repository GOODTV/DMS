using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;

public partial class DonateMgr_QuestionnaireDataList : BasePage
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
    protected void Page_Load(object sender, EventArgs e)
    {
        //有 npoGridView 時才需要
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        //權控處理
        AuthrityControl();
        if (!IsPostBack)
        {
            Session["cType"] = "";
            HFD_Uid.Value = Util.GetQueryString("Donor_Id");
            MemberMenu.Text = CaseUtil.MakeMenu(HFD_Uid.Value, 5); //傳入目前選擇的 tab
            Form_DataBind();
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
        //Authrity.CheckButtonRight("_AddNew", btnAdd);
    }
    //----------------------------------------------------------------------
    //帶入資料
    public void Form_DataBind()
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;

        //****變數設定****//
        uid = HFD_Uid.Value;

        //****設定查詢****//
        strSql = @" select *  , (Case When ISNULL(DONOR.Invoice_City,'')='' Then Invoice_Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.Invoice_ZipCode+B.mValue+Invoice_Address Else A.mValue+DONOR.Invoice_ZipCode+Invoice_Address End End) as [地址]
                    from DONOR 
                        Left Join CODECITY As A On Donor.Invoice_City=A.mCode Left Join CODECITY As B On Donor.Invoice_Area=B.mCode 
                    where Donor_id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("DonorInfo.aspx");

        DataRow dr = dt.Rows[0];

        //捐款人
        txtDonor_Name.Text = dr["Donor_Name"].ToString() + " " + dr["Title"].ToString(); ;
        //捐款人編號
        tbxDonor_Id.Text = dr["Donor_Id"].ToString();
        //收據開立
        tbxInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString();
        //收據地址
        tbxAddress.Text = dr["地址"].ToString();
    }
    public void LoadFormData()
    {
        DataTable dt;
        string strSql = "";
        strSql = @"SELECT [SerNo]
	                      ,CONVERT(VarChar,Create_Date,111) as [問卷日期]
                          ,Case When DonateMotive1='' Then '' Else 'V' End [支持媒體<br>宣教大平台，<br>可廣傳福音]
                          ,Case When DonateMotive2='' Then '' Else 'V' End [個人靈命<br>得造就]
                          ,Case When DonateMotive3='' Then '' Else 'V' End [支持優質<br>節目製作]
                          ,Case When DonateMotive4='' Then '' Else 'V' End [支持GOOD TV<br>家庭事工]
	                      ,Case When DonateMotive5='' Then '' Else 'V' End [感恩奉獻]
                          ,Case When WatchMode1='' Then '' Else 'V' End [GOOD TV<br>電視頻道]
                          ,Case When WatchMode2='' Then '' Else 'V' End [官網]
                          ,Case When WatchMode3='' Then '' Else 'V' End [Facebook]
                          ,Case When WatchMode4='' Then '' Else 'V' End [Youtube]
                          ,Case When WatchMode5='' Then '' Else 'V' End [好消息<br>月刊]
                          ,Case When WatchMode6='' Then '' Else 'V' End [GOOD TV<br>簡介刊物]
                          ,Case When WatchMode7='' Then '' Else 'V' End [教會牧者]
                          ,Case When WatchMode8='' Then '' Else 'V' End [親友]
                          ,Case When WatchMode9='' Then '' Else 'V' End [報章雜誌]
                          ,ToGOODTV as [給GOODTV<br>的話]
                          ,DonateWay as [問卷來源]
                      FROM [Donate_OnlineQuestion]";
        strSql += " where Donor_Id ='" + HFD_Uid.Value + "'";
        strSql += " order by CONVERT(VarChar,Create_DateTime,111) desc,SerNo desc";

        dt = NpoDB.QueryGetTable(strSql);

        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
            // 換掉原先的 Html Table
//            string strBody = "<table class='table_h' width='100%'>";
//            strBody += @"<TR><TH noWrap>問卷日期</SPAN></TH>
//                            <TH noWrap><SPAN>支持媒體<br>宣教大平台，<br>可廣傳福音</SPAN></TH>
//                            <TH noWrap><SPAN>個人靈命<br>得造就</SPAN></TH>
//                            <TH noWrap><SPAN>支持優質<br>節目製作</SPAN></TH>
//                            <TH noWrap><SPAN>支持GOOD TV<br>家庭事工</SPAN></TH>
//                            <TH noWrap><SPAN>感恩奉獻</SPAN></TH>
//                            <TH noWrap><SPAN>GOOD TV<br>電視頻道</SPAN></TH>
//                            <TH noWrap><SPAN>官網</SPAN></TH>
//                            <TH noWrap><SPAN>Facebook</SPAN></TH>
//                            <TH noWrap><SPAN>Youtube</SPAN></TH>
//                            <TH noWrap><SPAN>好消息<br>月刊</SPAN></TH>
//                            <TH noWrap><SPAN>GOOD TV<br>簡介刊物</SPAN></TH>
//                            <TH noWrap><SPAN>教會牧者</SPAN></TH>
//                            <TH noWrap><SPAN>親友</SPAN></TH>
//                            <TH noWrap><SPAN>報章雜誌</SPAN></TH>
//                            <TH noWrap><SPAN>給GOODTV<br>的話</SPAN></TH>
//                            <TH noWrap><SPAN>問卷來源</SPAN></TH></TR>";
//            foreach (DataRow dr in dt.Rows)
//            {
//                strBody += "<TR><TD><SPAN>" + dr["問卷日期"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["支持媒體<br>宣教大平台，<br>可廣傳福音"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["個人靈命<br>得造就"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["支持優質<br>節目製作"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["支持GOOD TV<br>家庭事工"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["感恩奉獻"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["GOOD TV<br>電視頻道"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["官網"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["Facebook"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["Youtube"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["好消息<br>月刊"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["GOOD TV<br>簡介刊物"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["教會牧者"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["親友"].ToString() + "</SPAN></TD>";
//                strBody += "<TD><SPAN>" + dr["報章雜誌"].ToString() + "</SPAN></TD>";
//                strBody += "<TD width='70px'><SPAN>" + dr["給GOODTV<br>的話"].ToString() + "</SPAN></TD>";
//                strBody += "<TD width='70px'><SPAN>" + dr["問卷來源"].ToString() + "</SPAN></TD>";

//                strBody += "</TR>";
//            }
//            strBody += "</table>";
//            lblGridList.Text = strBody;


            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.Keys.Add("SerNo");
            npoGridView.DisableColumn.Add("SerNo");
            npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            Session["cType"] = "QuestionnaireDataList";
            //npoGridView.EditLink = Util.RedirectByTime("Pledge_Edit.aspx", "Pledge_Id=");
            lblGridList.Text = npoGridView.Render();
        }
    }
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Session["cType"] = "QuestionnaireDataList";
        Response.Redirect(Util.RedirectByTime("Questionnaire_Add.aspx", "Donor_Id=" + HFD_Uid.Value));
    }
}