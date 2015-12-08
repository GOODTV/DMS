using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Net.Mail;
using System.Collections.Generic;
using log4net;

/// <summary>
/// 處理寄信功能程式
/// </summary>
public class SendEMailObject
{
    public string SmtpServer ="";
    public string SmtpAccount = "";
    public string SmtpPassword = "";
    public int ErrorCode = 0;
    public string ErrorMessage = "";
    public ILog logger = LogManager.GetLogger("RollingFileAppender");

	public SendEMailObject()
	{
    }
    public string SendEmail(string MailTo, string MailFrom, string Subject, string Body,string SmtpServer)
    {
        SmtpServer = System.Configuration.ConfigurationSettings.AppSettings["MailServer"];
        SmtpAccount = System.Configuration.ConfigurationSettings.AppSettings["MailUserID"];
        SmtpPassword = System.Configuration.ConfigurationSettings.AppSettings["MailUserPW"];
        this.SmtpServer = SmtpServer;
        return SendEmail(MailTo, MailFrom, Subject,  Body);
    }
    public string SendEmail(string MailTo, string MailFrom, string Subject, string Body)
    {
        string Result;
        SmtpServer = System.Configuration.ConfigurationSettings.AppSettings["MailServer"];
        SmtpAccount = System.Configuration.ConfigurationSettings.AppSettings["MailUserID"];
        SmtpPassword = System.Configuration.ConfigurationSettings.AppSettings["MailUserPW"];
        //Body += "<BR>SmtpServer:" + SmtpServer;
        //MailMessage Message = new MailMessage(MailTo, MailTo, Subject, Body);
        MailAddress from = new MailAddress(MailFrom, "GOODTV 捐款服務組");
        MailAddress to = new MailAddress(MailTo);        
        MailMessage Message = new MailMessage(from, to);
        Message.Subject = Subject;
        Message.Body = Body;

        Message.IsBodyHtml = true;

        SmtpClient client = new SmtpClient(SmtpServer, 25);
      
        //client.EnableSsl = true;
        //client.Credentials = new System.Net.NetworkCredential(this.SmtpAccount, this.SmtpPassword, SmtpServer);
        if (this.SmtpAccount != "")
        {
           //client.EnableSsl = true;           
            client.Credentials = new System.Net.NetworkCredential(SmtpAccount, SmtpPassword);
        }

        try
        {
            client.Send(Message);
            Result = "寄送信件成功";
            this.ErrorCode = 0;
            this.ErrorMessage = "";
            //this.logger.Info(Subject + " Mail To:" + MailTo + " Sucess!");
        }
        catch (ArgumentNullException ex)
        {
            this.ErrorCode = 1;//表示寄信失敗
            this.ErrorMessage = ex.Message;
            this.logger.Error(this.ErrorMessage);
            Result = ex.Message;            
        }
        catch (ArgumentOutOfRangeException ex)
        {
            this.ErrorCode = 2;//表示寄信失敗
            this.ErrorMessage = ex.Message;
            this.logger.Error(this.ErrorMessage);
            Result = ex.Message;
        }
        catch (InvalidOperationException ex)
        {
            this.ErrorCode = 3;//表示寄信失敗
            this.ErrorMessage = ex.Message;
            this.logger.Error(this.ErrorMessage);
            Result = ex.Message;
        }
        catch (SmtpFailedRecipientsException ex)
        {
            this.ErrorCode = 4;//表示寄信失敗
            this.ErrorMessage = ex.Message;
            this.logger.Error(this.ErrorMessage);
            Result = ex.Message;
        }
        catch (SmtpException ex)
        {
            this.ErrorCode = 5;//表示寄信失敗
            string strErrorMessage = "";
            
            SmtpStatusCode status = ex.StatusCode ;
            
            if (status == SmtpStatusCode.MailboxBusy)
            {
                strErrorMessage = "MailboxBusy";
            }
            else if (status == SmtpStatusCode.MailboxUnavailable)
            {
                strErrorMessage = "MailboxUnavailable";
            }
            else
            {
                strErrorMessage = "Failed to deliver message " + status.ToString();
            }

            this.ErrorMessage = strErrorMessage;
            this.logger.Error(this.ErrorMessage);
            Result = ex.Message;
        }
        catch (Exception ex)
        {
            this.ErrorCode = 6;//表示寄信失敗
            this.ErrorMessage = ex.Message;
            this.logger.Error(this.ErrorMessage);
            Result = ex.Message;
        }

        return Result;

    }
}



