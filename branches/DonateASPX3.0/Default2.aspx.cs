using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net.Mail;
public partial class Default2 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
  
    protected void Button1_Click1(object sender, EventArgs e)
    {
        string Mailhead = "";
        string SmtpServer = System.Configuration.ConfigurationSettings.AppSettings["MailServer"];
        string MailSubject = "";
        string MailBody = "";
        string MailTo = "";
        string MailFrom = "";
            SendEMailObject MailObject = new SendEMailObject();
            MailObject.SmtpServer = SmtpServer;
      
            ////若無帳密則不需設定
            //MailObject.SmtpAccount = SmtpAccount;
            //MailObject.SmtpPassword = SmtpPassword;

          
                MailObject.SmtpServer = SmtpServer;
                 MailSubject = "GOODTV線上奉獻確認信";
                 MailBody = ""+ Util.GetDBDateTime().ToString()  +"  "+ SmtpServer +"";
                 MailTo= "nashlee@npois.com.tw";
                
                 MailFrom = "edm@mail.goodtv.tv";
                string result = MailObject.SendEmail(MailTo, MailFrom, MailSubject, MailBody);
                if (MailObject.ErrorCode != 0)
                {
                    this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
                }
             
       // }
    }
}