using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateInfoConfirm: System.Web.UI.Page
{
    DataTable dtOnce = new DataTable();
    DataTable dtPeriod = new DataTable();
    //for insert DONATE_WEB
    string strDonorID;
    string strDonorName;
    string strSex;
    string strIDNo;
    string strBirthday;
    string strEducation;
    string strOccupation;
    string strMarriage;
    string strCellPhone;
    string strTelOfficeRegion;
    string strTelOffice;
    string strTelOfficeExt;
    string strTelHome;
    string strZipCode;
    string strCityCode;
    string strAreaCode;
    string strAddress;
    string strEmail;
    string strInvoiceZipCode;
    string strInvoiceCityCode;
    string strInvoiceAreaCode;
    string strInvoiceAddress;
	string strPhone;
    //string strDonate_Type;
    string strInvoiceType;
    // 2014/7/3 增加金流網站變數
    string strIpayHttp = System.Configuration.ConfigurationManager.AppSettings["ipay_http"];
    string StrEmailToDonations = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];
    string MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
    string strIsAnonymous;
    string strCity;
    string strArea;
    string strIsAbroad;
    string strOverseasCountry;
    string strOverseasAddress;

    protected void Page_Load(object sender, EventArgs e)
    {

        if (Session["ItemOnce"] == null && Session["ItemPeriod"] == null)
        {
            Util.ShowMsg("Timeout! Please donation again.");
            Response.Redirect(Util.RedirectByTime("DonateOnlineAll.aspx"));
        }

        // 2014/10/29 修改至金流網站的變數值(重要，請勿刪除) Transaction Site value begin
        storeid.Value = System.Configuration.ConfigurationManager.AppSettings["storeid"];
        password.Value = System.Configuration.ConfigurationManager.AppSettings["password"];
        customer.Value = System.Configuration.ConfigurationManager.AppSettings["customer"];
        // Transaction Site value end

        //顯示前頁msg
        if (Session["Msg"] != null)
        {
            Util.ShowMsg(Session["Msg"].ToString());
            Session["Msg"] = null;
        }
        if (Session["ItemOnce"] != null)
        {
            dtOnce = (DataTable)Session["ItemOnce"];
        }
        if (Session["ItemPeriod"] != null)
        {
            dtPeriod = (DataTable)Session["ItemPeriod"];
        }
        
        LoadFormData();
        
        if (!IsPostBack)
        {
            //Step項目設定
            
            if (Session["InsertPeriod"] == "Y")
            {
                this.lblStep3.ForeColor = System.Drawing.Color.MediumBlue;
                this.lblStep5.ForeColor = System.Drawing.Color.Red;
                this.lblStep5.Font.Bold = true;
                this.msgPanel.Visible = true;
                this.msgOnce.Visible = true;
            }
            else
            {
                /*this.lblStep3.ForeColor = System.Drawing.Color.Red;
                this.lblStep3.Font.Bold = true;
                this.lblStep5.ForeColor = System.Drawing.Color.MediumBlue;*/
                this.lblStep3.ForeColor = System.Drawing.Color.MediumBlue;
                this.lblStep5.ForeColor = System.Drawing.Color.Red;
                this.lblStep5.Font.Bold = true;
                this.msgPanel.Visible = false;
                this.msgOnce.Visible = false;
                this.msgPeriod.Visible = false;
            }           

            //if (Session["DonorName"] != null)
            //{
            //    lblTitle.Text = "親愛的 " + Session["DonorName"].ToString() + " " + lblTitle.Text;
            //}
            LoadDropDownList();

            if (dtOnce.Rows.Count > 0)
            {
                this.msgOnce.Visible = true;
                ShowCartOnce(dtOnce);
            }

            if (dtPeriod.Rows.Count > 0)
            {
                this.msgPeriod.Visible = true;
                ShowCartPeriod(dtPeriod);
                //this.btnConfirm.Text = "確認, 開始填寫授權資料";
                this.btnConfirm.Text = "Next";
            }

        }
    }
    //-------------------------------------------------------------------------------------------------------------
    private void LoadDropDownList()
    {
    }
    //---------------------------------------------------------------------------
    private void LoadFormData()
    {
        //Session["DonorID"] = null;
        if (Session["DonorID"] == null)
        {
            Session["DonorName"] = null;
            Util.ShowMsg("Timeout! Please sign in again.");
            Response.Redirect(Util.RedirectByTime("CheckOut.aspx"));
        }

        strDonorID = Session["DonorID"].ToString();

        string strSql = @"select * from Donor where Donor_Id=@Donor_Id and DeleteDate is null";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", strDonorID);

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            DataRow dr = dt.Rows[0];
            lblDonorName.Text = "Donor Name：" + dr["Donor_Name"].ToString();
			
			if(!string.IsNullOrEmpty(dr["Tel_Office"].ToString()))
			{
				strPhone= "(" + dr["Tel_Office_Loc"].ToString() + ")" + dr["Tel_Office"].ToString() + " Extention: " + dr["Tel_Office_Ext"].ToString();
			}else{
                strPhone = dr["Cellular_Phone"].ToString();
			}
				
            lblPhoneNum.Text = "Contact Phone：" + strPhone;	
			
			lblEmail.Text = "E-mail：" + dr["Email"].ToString();
			
            //for 單筆奉獻金流使用
            //customer:　購買客戶的姓名 (可不填)
            customer.Value = dr["Donor_Name"].ToString();
            //cellphone:　購買客戶的行動電話 (可不填)
            cellphone.Value = dr["Cellular_Phone"].ToString();
            //param :　提供客戶自行運用(可不填)
            param.Value = dr["Donor_Id"].ToString();

            strDonorName = dr["Donor_Name"].ToString();
            //strDonorID = dr["Donor_Id"].ToString();
            strSex = dr["Sex"].ToString();
            strIDNo = dr["IDNo"].ToString();
            strBirthday = dr["Birthday"].ToString();
            strEducation = dr["Education"].ToString();
            strOccupation = dr["Occupation"].ToString();
            strMarriage = dr["Marriage"].ToString();
            strCellPhone = dr["Cellular_Phone"].ToString();
            strTelOfficeRegion = dr["Tel_Office_Loc"].ToString();
            strTelOffice = dr["Tel_Office"].ToString();
            strTelOfficeExt = dr["Tel_Office_Ext"].ToString();
            strEmail = dr["Email"].ToString();
            // 2014/6/9 信用卡授權確認通知信
            Session["Email"] = strEmail;
            Session["DonorName"] = dr["Donor_Name"].ToString();
            strTelHome = dr["Tel_Home"].ToString();
            strZipCode = dr["ZipCode"].ToString();
            strCityCode = dr["City"].ToString();
            strAreaCode = dr["Area"].ToString();
            strAddress = dr["Address"].ToString();

            if (dr["Invoice_City"].ToString() != "")
            {
                strInvoiceZipCode = dr["Invoice_ZipCode"].ToString();
                strInvoiceCityCode = dr["Invoice_City"].ToString();
                strInvoiceAreaCode = dr["Invoice_Area"].ToString();
                strInvoiceAddress = dr["Invoice_Address"].ToString();
            }
            else
            {
                strInvoiceZipCode = dr["ZipCode"].ToString();
                strInvoiceCityCode = dr["City"].ToString();
                strInvoiceAreaCode = dr["Area"].ToString();
                strInvoiceAddress = dr["Address"].ToString();
            }
            strCity = Util.GetCityName(dr["City"].ToString());  //聯絡地址-縣市
            strArea = Util.GetAreaName(dr["Area"].ToString());  //鄉鎮 Area

            // 2014/6/5 修正收據開立判斷
            strInvoiceType = dr["Invoice_Type"].ToString();

            // 2014/7/4 增加徵信錄欄位,台灣地址
            //strIsAnonymous = dr["IsAnonymous"].ToString();
            strIsAbroad = dr["IsAbroad"].ToString();
            strOverseasCountry = dr["OverseasCountry"].ToString();
            strOverseasAddress = dr["OverseasAddress"].ToString();

        }
    }
    //---------------------------------------------------------------------------
    protected void btnConfirm_Click(object sender, EventArgs e)
    {

        // 2014/11/6 增加金流網站回覆頁面(DonateReturnUrl.aspx)的網頁重新整理記號清除
        Session.Remove("IsRefresh");

        if (dtPeriod.Rows.Count > 0)
        {
            Session["DonateContent"] = lblDonateContent.Text;
            Response.Redirect(Util.RedirectByTime("DonateCreditCard.aspx"));
        }
        else
        {            
            if (dtOnce.Rows.Count > 0)
            {
                //Response.Redirect(Util.RedirectByTime("DonateSingle.aspx"));

                //for 單筆奉獻金流使用
                //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
                //奉獻金額總計
                int sum_account = 0;
                //orderid:　訂單編號(不得重複,勿超過15碼)
                orderid.Value = DateTime.Now.Ticks.ToString().Substring(0, 15);
                for (int i = 0; i < dtOnce.Rows.Count; i++)
                {
                    //InsertDonateWeb
                    //purpose: 項目                    
                    purpose.Value = dtOnce.Rows[i]["奉獻項目"].ToString();       
                    //account:　金額(正整數)
                    account.Value = dtOnce.Rows[i]["奉獻金額"].ToString();
                    
                    //'新增交易資料(DONATE_WEB)
                    InsertDonateWeb();
                    
                    sum_account += Convert.ToInt32(dtOnce.Rows[i]["奉獻金額"].ToString());
                }

                //總奉獻金額帶回account.Value
                account.Value = sum_account.ToString();

                // 2014/7/4 增加內部Email通知
                InternalSendMail();

                //介接思遠金流
                string script = "<script> document.forms[0].action='" + strIpayHttp + "'; document.forms[0].submit(); </script>";
                Page.ClientScript.RegisterStartupScript(this.GetType(), "postform", script);
            }
            else
            {                
                Response.Redirect(Util.RedirectByTime("DonateOnlineAll.aspx"));
            }
        }
    }

    private void InternalSendMail()
    {
        //發送內部通知mail*****************************************
        string strBody = strDonorName + " 奉獻資料如下：<BR/><BR/>";
        strBody += "捐款方式：單筆奉獻 <BR/>";
        strBody += "捐款金額：" + account.Value + " 元 <BR/>";
        strBody += "E-mail：" + strEmail + "  <BR/>";
        strBody += "電話：" + strPhone + "<BR/>";
        strBody += "收據地址：" + (strIsAbroad == "N" ? strCity + strArea + strAddress : strOverseasAddress + " " + strOverseasCountry) + "<BR/>";
        strBody += "收據寄發：" + strInvoiceType + " <BR/>";
        //strBody += "徵信錄：" + (strIsAnonymous == "Y" ? "不刊登" : "刊登") + " <BR/>";

        SendEMailObject MailObject = new SendEMailObject();
        string MailSubject = " GOODTV線上奉獻_新奉獻通知(" + strDonorName + "_單筆奉獻_等待授權)";
        string MailBody = strBody;

        string result = MailObject.SendEmail(StrEmailToDonations, MailFrom, MailSubject, strBody);

        if (MailObject.ErrorCode != 0)
        {
            this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        }
        //********************************************
    
    }

    //-------------------------------------------------------------------------------------------------------------
    private void InsertDonateWeb()
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();
        SetColumnDataDonateWeb(list);
        string strSql = "";
        strSql = Util.CreateInsertCommand("DONATE_WEB", list, dict);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    //-------------------------------------------------------------------------------------------------------------
    private void SetColumnDataDonateWeb(List<ColumnData> list)
    {
        //'新增交易資料(DONATE_WEB)
        //SQL1="DONATE_WEB"
        //Set RS1=Server.CreateObject("ADODB.RecordSet")
        //RS1.Open SQL1,Conn,1,3
        //RS1.Addnew
        //RS1("Donate_Type")=Request.Form("Donate_Type")
        // 2014/6/4 修正收據開立判斷
        list.Add(new ColumnData("Donate_Type", "單次捐款", true, false, false));
        //list.Add(new ColumnData("Donate_Type", strDonate_Type, true, false, false));
        //RS1("od_sob")=od_sob
        list.Add(new ColumnData("od_sob", orderid.Value, true, false, false));
        //RS1("Dept_Id")=Request.Form("Donate_DeptId")
        list.Add(new ColumnData("Dept_Id", "C001", true, false, false));
        //RS1("Donor_Id")=Donor_Id
        list.Add(new ColumnData("Donor_Id", strDonorID, true, false, false));
        //RS1("Donate_CreateDate")=Date()
        list.Add(new ColumnData("Donate_CreateDate", Util.GetDBDateTime().ToShortDateString(), true, false, false));
        //RS1("Donate_CreateDateTime")=Now()
        list.Add(new ColumnData("Donate_CreateDateTime", Util.GetDBDateTime(), true, false, false));
        //RS1("Donate_CreateIP")=Request.ServerVariables("REMOTE_ADDR")
        list.Add(new ColumnData("Donate_CreateIP", Request.ServerVariables["REMOTE_ADDR"], true, false, false));
        //RS1("Donate_Amount")=Request.Form("Donate_Amount")
        list.Add(new ColumnData("Donate_Amount", account.Value, true, false, false));
        //RS1("Donate_DonorName")=Request.Form("Donate_Name")
        list.Add(new ColumnData("Donate_DonorName", customer.Value, true, false, false));
        //RS1("Donate_Sex")=Request.Form("Donate_Sex")
        list.Add(new ColumnData("Donate_Sex", strSex, true, false, false));
        //RS1("Donate_IDNO")=Request.Form("Donate_IDNO")
        list.Add(new ColumnData("Donate_IDNO", strIDNo, true, false, false));
        //If Request.Form("Donate_Birthday")<>"" Then RS1("Donate_Birthday")=Request.Form("Donate_Birthday")
        list.Add(new ColumnData("Donate_Birthday", (String.IsNullOrEmpty(strBirthday) ? null : Util.FixDateTime(strBirthday)), true, false, false));
        //RS1("Donate_Education")=Request.Form("Donate_Education")
        list.Add(new ColumnData("Donate_Education", strEducation, true, false, false));
        //RS1("Donate_Occupation")=Request.Form("Donate_Occupation")
        list.Add(new ColumnData("Donate_Occupation", strOccupation, true, false, false));
        //RS1("Donate_Marriage")=Request.Form("Donate_Marriage")
        list.Add(new ColumnData("Donate_Marriage", strMarriage, true, false, false));
        //RS1("Donate_CellPhone")=Request.Form("Donate_CellPhone")
        list.Add(new ColumnData("Donate_CellPhone", strCellPhone, true, false, false));
        //RS1("Donate_TelOffice_Region")=Request.Form("Donate_TelOffice_Region")
        list.Add(new ColumnData("Donate_TelOffice_Region", strTelOfficeRegion, true, false, false));
        //RS1("Donate_TelOffice")=Request.Form("Donate_TelOffice")
        list.Add(new ColumnData("Donate_TelOffice", strTelOffice, true, false, false));
        //RS1("Donate_TelOffice_Ext")=Request.Form("Donate_TelOffice_Ext")
        list.Add(new ColumnData("Donate_TelOffice_Ext", strTelOfficeExt, true, false, false));
        //RS1("Donate_TelHome")=Request.Form("Donate_TelHome")
        list.Add(new ColumnData("Donate_TelHome", strTelHome, true, false, false));
        //RS1("Donate_ZipCode")=Request.Form("Donate_ZipCode")
        list.Add(new ColumnData("Donate_ZipCode", strZipCode, true, false, false));
        //RS1("Donate_CityCode")=Request.Form("Donate_CityCode")
        list.Add(new ColumnData("Donate_CityCode", strCityCode, true, false, false));
        //RS1("Donate_AreaCode")=Request.Form("Donate_AreaCode")
        list.Add(new ColumnData("Donate_AreaCode", strAreaCode, true, false, false));
        //RS1("Donate_Address")=Request.Form("Donate_Address")
        list.Add(new ColumnData("Donate_Address", strAddress, true, false, false));
        //RS1("Donate_Email")=Request.Form("Donate_Email")
        list.Add(new ColumnData("Donate_Email", strEmail, true, false, false));
        //Modify by GoodTV-Tanya:20130910:修改線上奉獻多筆金額項目紀錄邏輯
        //RS1("Donate_Purpose")=Request.Form("Donate_Purpose")
        //list.Add(new ColumnData("Donate_Purpose", purpose.Value, true, false, false));
        list.Add(new ColumnData("Donate_Purpose", "經常費", true, false, false));
        //RS1("Donate_Invoice_Type")=Request.Form("Donate_Invoice_Type")
        // 2014/6/4 修正收據開立判斷
        //list.Add(new ColumnData("Donate_Invoice_Type", "單次收據", true, false, false));
        list.Add(new ColumnData("Donate_Invoice_Type", strInvoiceType, true, false, false));
        //RS1("Donate_Invoice_Title")=Request.Form("Donate_DonorName")
        list.Add(new ColumnData("Donate_Invoice_Title", strDonorName, true, false, false));
        //RS1("Donate_Invoice_IDNo")=Request.Form("Donate_IDNO")
        list.Add(new ColumnData("Donate_Invoice_IDNo", strIDNo, true, false, false));
        //RS1("Donate_Invoice_ZipCode")=Request.Form("Donate_Invoice_ZipCode")
        list.Add(new ColumnData("Donate_Invoice_ZipCode", strInvoiceZipCode, true, false, false));
        //RS1("Donate_Invoice_CityCode")=Request.Form("Donate_Invoice_CityCode")
        list.Add(new ColumnData("Donate_Invoice_CityCode", strInvoiceCityCode, true, false, false));
        //RS1("Donate_Invoice_AreaCode")=Request.Form("Donate_Invoice_AreaCode")
        list.Add(new ColumnData("Donate_Invoice_AreaCode", strInvoiceAreaCode, true, false, false));
        //RS1("Donate_Invoice_Address")=Request.Form("Donate_Invoice_Address")
        list.Add(new ColumnData("Donate_Invoice_Address", strInvoiceAddress, true, false, false));
        //RS1("Donate_Update")="N"
        // 2014/6/4 先改為Donate_Update='T' 暫停此功能
        //list.Add(new ColumnData("Donate_Update", "N", true, false, false));
        list.Add(new ColumnData("Donate_Update", "T", true, false, false));
        //RS1("Donate_Export")="N"
        list.Add(new ColumnData("Donate_Export", "N", true, false, false));
        //If Request.Form("Donate_ActId")<>"" Then RS1("Donate_Act_Id")=Request.Form("Donate_ActId")
        //RS1.Update
        //RS1.Close
        //Set RS1=Nothing
    }
    //---------------------------------------------------------------------------
    protected void btnPrev_Click(object sender, EventArgs e)
    {
        //Response.Redirect("ShoppingCart.aspx?Mode=FromDefault");
        Response.Redirect("CheckOut.aspx");
    }
    //-------------------------------------------------------------------------
    public void ShowCartOnce(DataTable dtSrc)
    {
        if (dtSrc.Rows.Count > 0)
        {
            DataTable dtSum = dtSrc.Clone();
            //計算合計
            int iSum = 0;
            DataRow drSum = dtSum.NewRow();
            //DataColumn[] keys = new DataColumn[1];
            //keys[0] = dtSum.Columns[0];
            //dtSum.PrimaryKey = keys;
            for (int i = 0; i < dtSrc.Rows.Count; i++)
            {
                iSum += Convert.ToInt32(dtSrc.Rows[i]["奉獻金額"].ToString());
                drSum = dtSum.NewRow();
                drSum["Uid"] = dtSrc.Rows[i]["Uid"].ToString();
                drSum["奉獻項目"] = dtSrc.Rows[i]["奉獻項目"].ToString();
                drSum["奉獻金額"] = dtSrc.Rows[i]["奉獻金額"].ToString();
                dtSum.Rows.Add(drSum);
            }
            drSum = dtSum.NewRow();
            drSum["奉獻項目"] = "Total";
            drSum["奉獻金額"] = iSum.ToString();
            dtSum.Rows.Add(drSum);

            //Grid initial
            //沒有需要特別處理的欄位時
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dtSum;
            npoGridView.Keys.Add("Uid");
            npoGridView.DisableColumn.Add("Uid");
            npoGridView.ShowPage = false;
            npoGridView.TableID = "ItemOnce";

            //-------------------------------------------------------------------------
            NPOGridViewColumn col = new NPOGridViewColumn("Description");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("奉獻項目");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("Amount");
            col.ColumnType = NPOColumnType.CurrecyText;
            col.ColumnName.Add("奉獻金額");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------

            lblDonateContent.Text += @"                
                    <span align='center' style='color: blue; font-weight: bold;'>
                        ※ One time Donation
                    </span>
                " + npoGridView.Render();
        }
    }
    //-------------------------------------------------------------------------
    public void ShowCartPeriod(DataTable dtSrc)
    {
        if (dtSrc.Rows.Count > 0)
        {
            DataTable dtNew = new DataTable();

            dtNew.Columns.Add("Uid");
            dtNew.Columns.Add("奉獻項目");
            dtNew.Columns.Add("奉獻周期");
            dtNew.Columns.Add("奉獻金額");

            for (int i = 0; i < dtSrc.Rows.Count; i++)
            {
                DataRow drNew = dtNew.NewRow();
                drNew["Uid"] = dtSrc.Rows[i]["Uid"].ToString();
                drNew["奉獻項目"] = dtSrc.Rows[i]["奉獻項目"].ToString();
                string strContent = "";
                //月繳
                if (dtSrc.Rows[i]["繳費方式"].ToString() == "月繳")
                {
                    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年";
                    //strContent += dtSrc.Rows[i]["開始月"].ToString() + "月 開始；";
                    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年";
                    //strContent += dtSrc.Rows[i]["開始月"].ToString() + "月至";
                    //strContent += dtSrc.Rows[i]["結束年"].ToString() + "年";
                    //strContent += dtSrc.Rows[i]["結束月"].ToString() + "月止；";
                }
                //季繳
                if (dtSrc.Rows[i]["繳費方式"].ToString() == "季繳")
                {
                    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年第";
                    //strContent += dtSrc.Rows[i]["開始月"].ToString() + "季至";
                    //strContent += dtSrc.Rows[i]["結束年"].ToString() + "年第";
                    //strContent += dtSrc.Rows[i]["結束月"].ToString() + "季止；";
                }
                //年繳
                if (dtSrc.Rows[i]["繳費方式"].ToString() == "年繳")
                {
                    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年 開始；";
                    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年至";
                    //strContent += dtSrc.Rows[i]["結束年"].ToString() + "年止；";
                }
                strContent += dtSrc.Rows[i]["繳費方式"].ToString();
                drNew["奉獻周期"] = strContent;
                drNew["奉獻金額"] = dtSrc.Rows[i]["奉獻金額"].ToString();
                dtNew.Rows.Add(drNew);
            }

            //Grid initial
            //沒有需要特別處理的欄位時
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dtNew;
            npoGridView.Keys.Add("Uid");
            npoGridView.DisableColumn.Add("Uid");
            npoGridView.ShowPage = false;
            npoGridView.TableID = "ItemPeriod";

            //-------------------------------------------------------------------------
            NPOGridViewColumn col = new NPOGridViewColumn("Description");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("奉獻項目");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("Period");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("奉獻周期");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("Amount");
            col.ColumnType = NPOColumnType.CurrecyText;
            col.ColumnName.Add("奉獻金額");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------

            lblDonateContent.Text += @"                
                    <span align='center' style='color: Blue; font-weight: bold;'>
                        ※ Recurring Donation
                    </span>
                " + npoGridView.Render() /*+ @"<br><span align='center' style='color: red; font-size: small;'>
                        1.定期定額捐款扣款方式：捐款金額月繳是每月固定25日進行授權扣款，季繳為每三個月扣款一次，以此類推。<br><br>
                        2.若需終止定期定額奉獻，請電洽-奉獻服務專線(02)8024-3911 捐款服務組，謝謝您的支持。
                    </span>"*/;

            Session["DonatePeriod"] = @"                
                            <span align='center' style='color: Blue; font-weight: bold;'>
                                ※ Recurring Donation
                            </span>
                        " + npoGridView.Render();
        }
    }
}