using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Online_VerificationCode : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (HFD_chkItem.Value.Contains("單筆奉獻"))
            {
                lblStep3.Text = " 確認奉獻明細 ";
                lblStep4.Visible = false;
                lblStep5.Visible = false;
            }
            //SendCode.CssClass = "css_btn_class_login";
        }
    }

    protected void SendCode_Click(object sender, EventArgs e)
    {

        HFD_Email.Value = txtEmail.Text;
        // 預設資料庫的捐款人的姓名、Email與有密碼是唯一的
        string strSql = "select Donor_Id,Donor_Name from DONOR where DeleteDate is null and isnull(Donor_Pwd,'') <> '' and Email=@Email ;";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Email", HFD_Email.Value);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count > 0)
        {
            HFD_DonorID.Value = dt.Rows[0]["Donor_Id"].ToString();
            HFD_DonorName.Value = dt.Rows[0]["Donor_Name"].ToString();
            string strCode = RandomNumber(4);
            string strToday = String.Format("{0:yyyyMd}",DateTime.Today);
            HFD_code.Value = Convert.ToString((Convert.ToInt64(strCode)) * (Convert.ToInt64(strToday)));

            //string strEncryptDonorId = Util.desEncryptBase64(strDonorId);
            string stremailHead = "親愛的 " + HFD_DonorName.Value + " 平安！<BR><BR>";
            string strReport = "感謝您在GOODTV網站上的持續奉獻，請於網頁上輸入此驗證碼：" + strCode;
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

            //SendCode.CssClass = "css_btn_class";
            //UpdatePwd.CssClass = "css_btn_class_login";

            //帳號
            txtEmail.Enabled = false;
            rfvEmail.Enabled = false;
            SendCode.Visible = false;

            //驗證碼
            txtCode.Enabled = true;
            rfvCode.Enabled = true;
            cvCode.Enabled = true;
            txtDonorPwd1.Enabled = true;
            rfvDonorPwd1.Enabled = true;
            cvDonorPwd1.Enabled = true;
            txtDonorPwd2.Enabled = true;
            rfvDonorPwd2.Enabled = true;
            cvDonorPwd2.Enabled = true;
            UpdatePwd.Visible = true;

            try
            {
                string result = MailObject.SendEmail(HFD_Email.Value, strEmailFrom, strEmailSubject, strEmailBody);
                //Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('已送出驗證碼，請從Email取得驗證碼');</script>");
                SendOK.Text = "已送出驗證碼，請從Email取得驗證碼";
                SendOK.CssClass = "show";
            }
            catch (Exception ex)
            {
                //Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert(" + ex.ToString() + ");</script>");
                SendOK.Text = ex.ToString();
                SendOK.CssClass = "blink";
            }
        }
        else
        {
            //Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('找不到您所輸入的帳號(Email)資料，謝謝！');</script>");
            SendOK.Text = "找不到您所輸入的帳號(Email)資料";
            SendOK.CssClass = "blink";
        }


    }

    protected void UpdatePwd_Click(object sender, EventArgs e)
    {


        if (txtDonorPwd1.Text.Trim() == "" || txtDonorPwd2.Text.Trim() == "")
        {
            //Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('新密碼欄位或確認密碼欄位未輸入！');</script>");
            UpdateOK.Text = "新密碼欄位或確認密碼必填欄位";
            UpdateOK.CssClass = "blink";

        }
        else if (txtDonorPwd1.Text.Trim() == txtDonorPwd2.Text.Trim())
        {
            string strSql = "update DONOR set Donor_Pwd=@Donor_Pwd where isnull(Donor_Pwd,'') <> '' and Donor_Id=@Donor_Id and Email=@Email ;";
            Dictionary<string, object> dict = new Dictionary<string, object>();
            dict.Add("Donor_Id", int.Parse(HFD_DonorID.Value));
            dict.Add("Email", HFD_Email.Value);
            dict.Add("Donor_Pwd", txtDonorPwd1.Text.Trim());
            int cut = NpoDB.ExecuteSQLS(strSql, dict);
            if (cut > 0)
            {
                string stremailHead = "親愛的 " + HFD_DonorName.Value + " 平安！<BR><BR>";
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
                    //Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('重設密碼成功！'); GoBack();</script>");

                    UpdateOK.Text = "重設密碼成功！";
                    UpdateOK.CssClass = "show";

                    UpdatePwd.Visible = false;


                }
                catch (Exception ex)
                {
                    //Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert(" + ex.ToString() + ");</script>");
                    UpdateOK.Text = ex.ToString();
                    UpdateOK.CssClass = "blink";
                }

            }
            else
            {
                //Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('重設密碼失敗！');</script>");
                UpdateOK.Text = "重設密碼失敗";
                UpdateOK.CssClass = "blink";
            }

        }
        else
        {

            //Page.ClientScript.RegisterStartupScript(this.GetType(), "PopupScript", "<script type='text/javascript'>alert('輸入的新密碼與確認密碼不相同！');</script>");
            UpdateOK.Text = "輸入的新密碼與確認密碼不相符";
            UpdateOK.CssClass = "blink";
        }

    }

    /// <summary>
    /// 產生驗證碼
    /// </summary>
    /// <param name="Count">驗證碼的個數</param>
    /// <returns>返回生成的亂數</returns>
    public string RandomNumber(int Count)
    {
        //string strChar = "0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,a,b,c,d,e,f
        //,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z";
        string strChar = "0,1,2,3,4,5,6,7,8,9";
        string[] VcArray = strChar.Split(',');
        string VNum = "";
        int temp = -1;      //記錄上次亂數值，儘量避免產生幾個一樣的亂數

        //採用一個簡單的演算法以保證生成亂數的不同
        Random random = new Random();
        for (int i = 1; i < Count + 1; i++)
        {
            if (temp != -1)
                random = new Random(i * temp * unchecked((int)DateTime.Now.Ticks));
            //int t = random.Next(61);
            int t = random.Next(9);
            if (temp != -1 && temp == t)
                return RandomNumber(Count);
            temp = t;
            VNum += VcArray[t];
        }
        return VNum; //返回生成的亂數
    }

}
