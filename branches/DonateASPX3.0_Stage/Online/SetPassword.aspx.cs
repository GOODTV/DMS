using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Online_SetPassword : System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!Page.IsPostBack)
        {
            string strEncryptDonorId = Util.GetQueryString("DonorId");
            if (strEncryptDonorId.ToString() == "")
            {
                setpwd.Visible = false;
                Response.Write("請勿直接輸入網址！謝謝！");
            }
            else
            {
                string strDonorId = Util.desDecryptBase64(Util.GetQueryString("DonorId"));
                TxtDonorId.Text = strDonorId;
                //Session["DonorId"] = strDonorId;

                // 回傳捐款人編號撈取捐款人的姓名、Email(捐款人的姓名、Email與有密碼是唯一的)
                string strSql = "select Donor_Name,Email from DONOR where DeleteDate is null and isnull(Donor_Pwd,'') <> '' and Donor_Id=@Donor_Id ;";
                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Donor_Id", int.Parse(strDonorId));
                DataTable dt = NpoDB.GetDataTableS(strSql, dict);
                if (dt.Rows.Count > 0)
                {
                    DataRow dr = dt.Rows[0];
                    txtDonorName.Text = dr["Donor_Name"].ToString();
                    txtEmail.Text = dr["Email"].ToString();

                }
            }
        }

        // 從取消退回再重設密碼送出時可正常顯示資料
        //Session["DonorId"] = TxtDonorId.Text;

    }

    protected void Button1_Click(object sender, EventArgs e)
    {


        if (txtDonorPwd1.Text.Trim() == "" || txtDonorPwd2.Text.Trim() == "")
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('密碼欄位或確認密碼欄位未輸入！');</script>");
        }
        else if (txtDonorPwd1.Text.Trim() == txtDonorPwd2.Text.Trim())
        {
            string strSql = "update DONOR set Donor_Pwd=@Donor_Pwd where isnull(Donor_Pwd,'') <> '' and Donor_Id=@Donor_Id ;";
            Dictionary<string, object> dict = new Dictionary<string, object>();
            dict.Add("Donor_Id", int.Parse(TxtDonorId.Text));
            dict.Add("Donor_Pwd", txtDonorPwd1.Text.Trim());
            int cut = NpoDB.ExecuteSQLS(strSql, dict);
            if (cut > 0)
            {
                string stremailHead = "親愛的 " + txtDonorName.Text + " 平安！<BR><BR>";
                string strReport = "感謝您在GOODTV網站上的持續奉獻，您所申請的重設密碼已成功完成。<BR><BR>";
                strReport += "若您有任何操作上的問題敬請來電: 02-8024-3911 <BR>";
                strReport += "願神祝福您!<BR>";
                strReport += "GOODTV捐款服務組敬上<BR>";
                string SmtpServer = System.Configuration.ConfigurationManager.AppSettings["MailServer"];
                string strEmailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];

                SendEMailObject MailObject = new SendEMailObject();
                MailObject.SmtpServer = SmtpServer;
                string strEmailSubject = "GOODTV線上奉獻重設密碼完成信";
                string strEmailBody = stremailHead + strReport;

                try
                {
                    string result = MailObject.SendEmail(txtEmail.Text, strEmailFrom, strEmailSubject, strEmailBody);
                    Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('重設密碼成功！');window.close();</script>");
                }
                catch (Exception ex)
                {
                    Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert(" + ex.ToString() + ");</script>");
                }

            }
            else
            {
                Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('重設密碼失敗！');</script>");

            }

        }
        else
        {

            Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('輸入的密碼與確認密碼不相同！');</script>");

        }

    }

}
