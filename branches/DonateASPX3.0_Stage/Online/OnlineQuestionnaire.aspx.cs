using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Online_OnlineQuestionnaire : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_UID.Value = Request.Params["Donor_Id"].ToString();
            HFD_OrderId.Value = Request.Params["OrderNumber"].ToString();
            HFD_Device.Value = Request.Params["Device"].ToString();
        }
    }
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = "insert into  Donate_OnlineQuestion\n";
        strSql += "( OrderNumber, Donor_Id, DonateMotive1, DonateMotive2, DonateMotive3, DonateMotive4, DonateMotive5, \n";
        strSql += " WatchMode1, WatchMode2, WatchMode3, WatchMode4, WatchMode5, WatchMode6, WatchMode7, WatchMode8, WatchMode9, Device, ToGOODTV) values\n";
        strSql += "(@OrderNumber,@Donor_Id,@DonateMotive1,@DonateMotive2,@DonateMotive3,@DonateMotive4,@DonateMotive5, \n";
        strSql += "@WatchMode1,@WatchMode2,@WatchMode3,@WatchMode4,@WatchMode5,@WatchMode6,@WatchMode7,@WatchMode8,@WatchMode9,@Device,@ToGOODTV)\n";

        dict.Add("OrderNumber", HFD_OrderId.Value);
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
        if (HFD_Device.Value == "1")
        {
            dict.Add("Device", "一般網頁");
        }
        else if (HFD_Device.Value == "2")
        {
            dict.Add("Device", "行動版");
        }
        else if (HFD_Device.Value == "3")
        {
            dict.Add("Device", "OTT");
        }
        else
        {
            dict.Add("Device", "不明");
        }
        dict.Add("ToGOODTV", txtToGoodTV.Text.Trim().ToString());
        NpoDB.ExecuteSQLS(strSql, dict);
        ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"感謝您撥出寶貴的時間來完成問卷調查。願神祝福您！\");window.close();</script>");
    }
}