using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Online_ResetPassword : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void Button1_Click(object sender, EventArgs e)
    {

        string strEmailTo = txtEmail.Text;
        string strDonorName = txtDonorName.Text;
        string strDonorId="";
        // 預設資料庫的捐款人的姓名、Email與有密碼是唯一的
        string strSql = "select Donor_Id from DONOR where DeleteDate is null and isnull(Donor_Pwd,'') <> '' and Donor_Name=@Donor_Name and Email=@Email ;";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Name", strDonorName);
        dict.Add("Email", strEmailTo);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count > 0)
        {
            strDonorId = dt.Rows[0][0].ToString();

            //string strEncryptDonorId = Util.desEncryptBase64(strDonorId);
            string stremailHead = "親愛的 " + txtDonorName.Text + " 平安！<BR><BR>";
            string strReport = "感謝您在GOODTV網站上的持續奉獻，請點選以下連結進入網站重新設定密碼。<BR>";
            strReport += "https://i-donate.goodtv.tv/Online/SetPassword.aspx?DonorId=" + HttpUtility.UrlEncode(Util.desEncryptBase64(strDonorId));
            //strReport += "http://localhost:88/Online/SetPassword.aspx?DonorId=" + Util.desEncryptBase64(strDonorId);
            strReport += "<BR><BR><BR>";
            strReport += "若您有任何操作上的問題敬請來電: 02-8024-3911 <BR>";
            strReport += "願神祝福您!<BR>";
            strReport += "GOODTV捐款服務組敬上<BR>";
            string SmtpServer = System.Configuration.ConfigurationManager.AppSettings["MailServer"];
            string strEmailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];

            SendEMailObject MailObject = new SendEMailObject();
            MailObject.SmtpServer = SmtpServer;
            string strEmailSubject = "GOODTV線上奉獻重設密碼確認信";
            string strEmailBody = stremailHead + strReport;

            try
            {
                string result = MailObject.SendEmail(strEmailTo, strEmailFrom, strEmailSubject, strEmailBody);
                Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('已送出重設密碼確認信');</script>");
            }
            catch (Exception ex)
            {
                Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert(" + ex.ToString() + ");</script>");
            }
        }
        else
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('找不到您所輸入的姓名和Email資料，謝謝！');</script>");
        }


    }


    protected void Button2_Click(object sender, EventArgs e)
    {
        Response.Redirect("CheckOut.aspx");
    }

}

