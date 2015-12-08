using System;
using System.Collections.Generic;
using System.Web;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Web.Configuration;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.IO;
using System.Text;
using System.Data;
using System.Web.SessionState;

/// <summary>
/// MailFun 的摘要描述
/// </summary>

public class MailFunc
{
   
	public MailFunc()
	{
	}
    private static System.Web.UI.Page the;
    public static bool CheckMail(string mail)
    {
        mail = mail.Replace(" ", "");
        if (mail.IndexOf(",") > 0 || mail.IndexOf(";") > 0 || mail.IndexOf("'") > 0 || mail.IndexOf("''") > 0)
        {
            return false;
        }
        if (Regex.IsMatch(mail, ".*@.{2,}\\..{2,}"))
            return true;
        else
            return false;
    }

    public static void SendMail(MailAddress _from, MailAddress _to, string _subject, string _message)
    {
        MailMessage message = new MailMessage();
        message.From = _from;
        message.To.Add(_to);
        message.Subject = _subject;
        message.Body = _message;
        message.IsBodyHtml = true;

        SmtpClient mailClient = new SmtpClient(WebConfigurationManager.AppSettings.Get("MailServer"));
        System.Net.NetworkCredential NC = new System.Net.NetworkCredential();
        NC.UserName = WebConfigurationManager.AppSettings.Get("MailUserID");
        NC.Password = WebConfigurationManager.AppSettings.Get("MailUserPW");
        mailClient.Credentials = NC;
        mailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
        try
        {
            mailClient.Send(message);
        }
        catch (System.Net.Mail.SmtpException ex)
        {
          //  DBFunction.SysLogBOX(ex.Source + "：" + ex.Message);
            throw;
        }
    }
    public static bool SendMailLog(string epaperid, MailAddress _from, string _to, string _subject, string _message)
    {
        MailMessage message = new MailMessage();
        message.From = _from;
        message.To.Add(new MailAddress(_to));
        message.Subject = _subject;
        message.Body = _message;
        message.IsBodyHtml = true;
        SmtpClient mailClient = new SmtpClient(WebConfigurationManager.AppSettings.Get("MailServer"));
        System.Net.NetworkCredential NC = new System.Net.NetworkCredential();
        NC.UserName = WebConfigurationManager.AppSettings.Get("MailUserID");
        NC.Password = WebConfigurationManager.AppSettings.Get("MailUserPW");

        mailClient.DeliveryMethod = SmtpDeliveryMethod.Network;

        try
        {
            //mailClient.Send(message);
        }
        catch (System.Net.Mail.SmtpException ex)
        {

              WriteMailLog(epaperid, ex.Message, ex.StatusCode.ToString(), _to);
            return false;
        }
        return true;
    }
    public static void WriteMailLog(string epaperid, string message, string code, string email)
    {
        string strSql = "insert into maillog (mlg_epa_num,mlg_message,mlg_code,mlg_email) values (@mlg_epa_num,@mlg_message,@mlg_code,@mlg_email)";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("mlg_epa_num", epaperid);
        dict.Add("mlg_message", message);
        dict.Add("mlg_code", code);
        dict.Add("mlg_email", email);
       // CmsADODB.ExecuteSQL(the, strSql, dict);
    }
  
 }