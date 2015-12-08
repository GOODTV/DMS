using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class DonateMgr_Questionnaire_Add : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Donor_Id");
            Form_DataBind();
        }
    }
    //帶入資料
    public void Form_DataBind()
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;

        //****變數設定****//
        uid = HFD_Uid.Value;

        //****設定查詢****//
        strSql = @" Select * ,(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.Invoice_ZipCode+B.mValue+Invoice_Address Else A.mValue+DONOR.Invoice_ZipCode+Invoice_Address End End) as [地址]
                    From Donor 
                        Left Join CODECITY As A On Donor.Invoice_City=A.mCode Left Join CODECITY As B On Donor.Invoice_Area=B.mCode 
                    where Donor_id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("DonateInfo.aspx");

        DataRow dr = dt.Rows[0];

        //捐款人
        tbxDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
        //捐款人編號
        tbxDonor_Id.Text = dr["Donor_Id"].ToString().Trim();
        //身份別
        tbxDonor_Type.Text = dr["Donor_Type"].ToString().Trim();
        //收據地址
        tbxAddress.Text = dr["地址"].ToString();
        //收據開立
        tbxInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //最近捐款日：
        if (dr["Last_DonateDate"].ToString().Trim() != "")
        {
            tbxLast_DonateDate.Text = DateTime.Parse(dr["Last_DonateDate"].ToString().Trim()).ToShortDateString().ToString();
        }
    }
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = "insert into  Donate_OnlineQuestion\n";
        strSql += "( Donor_Id, DonateMotive1, DonateMotive2, DonateMotive3, DonateMotive4, DonateMotive5, \n";
        strSql += " WatchMode1, WatchMode2, WatchMode3, WatchMode4, WatchMode5, WatchMode6, WatchMode7, WatchMode8, WatchMode9, Device, ToGOODTV,Create_Date,Create_DateTime,DonateWay) values\n";
        strSql += "(@Donor_Id,@DonateMotive1,@DonateMotive2,@DonateMotive3,@DonateMotive4,@DonateMotive5, \n";
        strSql += "@WatchMode1,@WatchMode2,@WatchMode3,@WatchMode4,@WatchMode5,@WatchMode6,@WatchMode7,@WatchMode8,@WatchMode9,@Device,@ToGOODTV,@Create_Date,@Create_DateTime,@DonateWay)\n";

        dict.Add("Donor_Id", HFD_Uid.Value);
        //奉獻動機
        if (this.DonateMotive1.Checked)
        {
            dict.Add("DonateMotive1", "支持媒體宣教大平台，可廣傳福音");
        }
        else
        {
            dict.Add("DonateMotive1", "");
        }
        if (this.DonateMotive2.Checked)
        {
            dict.Add("DonateMotive2", "個人靈命得造就");
        }
        else
        {
            dict.Add("DonateMotive2", "");
        }
        if (this.DonateMotive3.Checked)
        {
            dict.Add("DonateMotive3", "支持優質節目製作");
        }
        else
        {
            dict.Add("DonateMotive3", "");
        }
        if (this.DonateMotive4.Checked)
        {
            dict.Add("DonateMotive4", "支持GOOD TV家庭事工");
        }
        else
        {
            dict.Add("DonateMotive4", "");
        }
        if (this.DonateMotive5.Checked)
        {
            dict.Add("DonateMotive5", "感恩奉獻");
        }
        else
        {
            dict.Add("DonateMotive5", "");
        }
        //收看管道
        if (this.WatchMode1.Checked)
        {
            dict.Add("WatchMode1", "GOOD TV電視頻道");
        }
        else
        {
            dict.Add("WatchMode1", "");
        }
        if (this.WatchMode2.Checked)
        {
            dict.Add("WatchMode2", "官網");
        }
        else
        {
            dict.Add("WatchMode2", "");
        }
        if (this.WatchMode3.Checked)
        {
            dict.Add("WatchMode3", "Facebook");
        }
        else
        {
            dict.Add("WatchMode3", "");
        }
        if (this.WatchMode4.Checked)
        {
            dict.Add("WatchMode4", "Youtube");
        }
        else
        {
            dict.Add("WatchMode4", "");
        }
        if (this.WatchMode5.Checked)
        {
            dict.Add("WatchMode5", "好消息月刊");
        }
        else
        {
            dict.Add("WatchMode5", "");
        }
        if (this.WatchMode6.Checked)
        {
            dict.Add("WatchMode6", "GOOD TV簡介刊物");
        }
        else
        {
            dict.Add("WatchMode6", "");
        }
        if (this.WatchMode7.Checked)
        {
            dict.Add("WatchMode7", "教會牧者");
        }
        else
        {
            dict.Add("WatchMode7", "");
        }
        if (this.WatchMode8.Checked)
        {
            dict.Add("WatchMode8", "親友");
        }
        else
        {
            dict.Add("WatchMode8", "");
        }
        if (this.WatchMode9.Checked)
        {
            dict.Add("WatchMode9", "報章雜誌");
        }
        else
        {
            dict.Add("WatchMode9", "");
        }
        dict.Add("Device", "一般網頁");
        dict.Add("ToGOODTV", txtToGoodTV.Text.Trim().ToString());
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        dict.Add("DonateWay", tbxDonateWay.Text);

        NpoDB.ExecuteSQLS(strSql, dict);
        //SetSysMsg("問卷資料新增成功！");
        Response.Redirect(Util.RedirectByTime("QuestionnaireDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("QuestionnaireDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
    }
}