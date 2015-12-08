using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateDone: System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {
            HFD_UID.Value = Request.Params["HFD_DonorID"];
            //捐款人名稱
            lblTitle.Text = Request.Params["HFD_DonorName"];
            //定期授權資料
            lblDonateContent.Text += "<span style='font-weight: bold;' align='center'>※　定　期　定　額　奉　獻</span>";
            lblDonateContent.Text += "<table width='100%' class='table_h'><tr><th>奉獻項目</th>";
            lblDonateContent.Text += "<th>奉獻周期</th><th>金額</th></tr>";
            lblDonateContent.Text += "<tr><td style='text-align: center;'>為GOODTV奉獻</td>";
            lblDonateContent.Text += "<td style='text-align: center;'>" + Request.Params["HFD_PayType"] + "</td>";
            lblDonateContent.Text += "<td style='text-align: right;'><span><font color='red'>NT$ " + String.Format("{0:0,0}", Convert.ToInt32(Request.Params["HFD_Amount"])) + "</font></span></td>";
            lblDonateContent.Text += "</tr></table>";
            HFD_Question.Value = "Y";
        }

    }

    //---------------------------------------------------------------------------
    protected void btnBackDefault_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("DonateOnlineAll.aspx"));
    }
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        //2015112 有填才新增問卷
        if (this.DonateMotive1.Checked || this.DonateMotive2.Checked || this.DonateMotive3.Checked || this.DonateMotive4.Checked || this.DonateMotive5.Checked || this.WatchMode1.Checked || this.WatchMode2.Checked || this.WatchMode3.Checked || this.WatchMode4.Checked || this.WatchMode5.Checked || this.WatchMode6.Checked || this.WatchMode7.Checked || this.WatchMode8.Checked || this.WatchMode9.Checked || txtToGoodTV.Text != "")
        {
            Dictionary<string, object> dict = new Dictionary<string, object>();
            string strSql = "insert into  Donate_OnlineQuestion\n";
            strSql += "(Donor_Id, DonateMotive1, DonateMotive2, DonateMotive3, DonateMotive4, DonateMotive5, \n";
            strSql += " WatchMode1, WatchMode2, WatchMode3, WatchMode4, WatchMode5, WatchMode6, WatchMode7, WatchMode8, WatchMode9, Device, ToGOODTV,Create_Date,Create_DateTime,DonateWay) values\n";
            strSql += "(@Donor_Id,@DonateMotive1,@DonateMotive2,@DonateMotive3,@DonateMotive4,@DonateMotive5, \n";
            strSql += "@WatchMode1,@WatchMode2,@WatchMode3,@WatchMode4,@WatchMode5,@WatchMode6,@WatchMode7,@WatchMode8,@WatchMode9,@Device,@ToGOODTV,@Create_Date,@Create_DateTime,@DonateWay)\n";

            dict.Add("Donor_Id", HFD_UID.Value);
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
            dict.Add("DonateWay", "線上定期定額");

            NpoDB.ExecuteSQLS(strSql, dict);
            ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"感謝您撥出寶貴的時間來完成問卷調查。願神祝福您！\");</script>");
        }
        HFD_Question.Value = "";
    }
}
