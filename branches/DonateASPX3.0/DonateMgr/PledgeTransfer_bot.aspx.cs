using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;
using System.Diagnostics;

public partial class DonateMgr_Pledge_Transfer_bot : System.Web.UI.Page
{
    string Donate_Year_AD;
    string Donate_Year;
    string Donate_Month;
    string Donate_Day;
    string StrEmailToDonations = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations2"];
    string MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];

    protected void Page_Load(object sender, EventArgs e)
    {
        string strSql = "Select Pledge_Id From Pledge Where Donate_Payment='信用卡授權書(一般)' and Donate_Period='單筆' and Status = '授權中' and Create_Date = CONVERT(varchar(10) , getdate(), 111) ";
        //string strSql = "Select Pledge_Id From Pledge Where Donate_Payment='信用卡授權書(一般)' and Donate_Period='單筆' and Status = '授權中' and Create_Date = '2014/09/11' ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        DataRow dr;
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count != 0)
        {
            string[] Pledge_Ids = new string[dt.Rows.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dr = dt.Rows[i];
                Pledge_Ids[i] = dr["Pledge_Id"].ToString();
            }
            for (int i = 0; i < Pledge_Ids.Length; i++)
            {
                /*if (i == Pledge_Ids.Length - 1)
                {
                    Session["Pledge_Id"] +=  Pledge_Ids[i];
                }
                else
                {*/
                    Session["Pledge_Id"] += Pledge_Ids[i] + " ";
                //}
            }
        }
        if (Session["Pledge_Id"] != null)
        {
            string Uids = Session["Pledge_Id"].ToString();
            Print_Txt(Uids);
        }
        else
            return;
    }
    //---------------------------------------------------------------------------
    private void Print_Txt(string Uids)
    {
        if (Uids == "")
        {
            //Page.Controls.Add(new LiteralControl("<script>alert(\"查無資料！\");</script>"));
            Page.Controls.Add(new LiteralControl("<script>SelfClose();</script>"));
            return;
        }

        //抓出要轉出的txt格式資料
        string strSql = @"Select * From DONATE_TRANSFER Where Transfer_AspxName='Transfer_bot.aspx'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];  
                      
        GridList.Text = GetTable(Uids, dr);
        
        //輸出檔案名稱
        string Transfer_FileName = "";
        Transfer_FileName = dr["Transfer_FileName"].ToString();
        //匯出檔名加上交易扣款日期
        Transfer_FileName = Transfer_FileName + Donate_Month + Donate_Day + ".01";
        //string filepath = Request.PhysicalApplicationPath + "UpLoad\\BK\\" + Transfer_FileName;
        //System.IO.File.WriteAllText(filepath, GridList.Text);
        byte[] byteArray = Encoding.ASCII.GetBytes(GridList.Text);
        MemoryStream sm = new MemoryStream(byteArray);

        //20140911增加內部Email通知
        //發送內部通知mail*****************************************
        string strBody = "親愛的同工平安：<BR/><BR/>";
        strBody += "附件為 <font color = 'blue'>" + DateTime.Now.ToString("yyyy-MM-dd") + "</font> 輸入的單筆待授權匯出檔。<BR/><BR/>";
        strBody += "請審核確認後轉寄台銀授權處理！<BR/>";

        SendEMailObject MailObject = new SendEMailObject();
        string MailSubject = "信用卡紙本待授權提示(單筆) - 日期：" + DateTime.Now.ToString("yyyy-MM-dd");
        string MailBody = strBody;

        //string result = MailObject.SendEmailAttachment(StrEmailToDonations, MailFrom, MailSubject, strBody, filepath);
        string result = MailObject.SendEmailAttachment(StrEmailToDonations, Transfer_FileName, MailSubject, strBody, sm);
        if (MailObject.ErrorCode != 0)
        {
            this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        }
        //********************************************
        //Util.OutputTxt(GridList.Text, "5", Transfer_FileName);
        Session["Pledge_Id"]="";
        Page.Controls.Add(new LiteralControl("<script>SelfClose();</script>"));
    }
    //-------------------------------------------------------------------------------------------------------------
    private string GetTable(string Uids, DataRow dr_Txt)
    {
        //抓出txt中的資料
        string[] Pledge_Ids = Uids.Split(' ');
        string content = "";
        //資料長度
        int DataLen = 0;
        
        //授權筆數
        int row = 0;
        //授權金額
        int Donate_Total = 0;
        
        //交易日期:交易日期為畫面上的扣款日期       
        Donate_Year_AD = (Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/dd")).Year - 1911).ToString();
        Donate_Year = Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/dd")).Year.ToString();
        if (Donate_Year_AD.Length == 2)
        {
            Donate_Year_AD = "00" + Donate_Year_AD;
        }
        else if (Donate_Year_AD.Length == 3)
        {
            Donate_Year_AD = "0" + Donate_Year_AD;
        }
        
        Donate_Month = Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/dd")).Month.ToString();
        if (Donate_Month.Length == 1)
        {
            Donate_Month = "0" + Donate_Month;

        }
        
        Donate_Day = Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/dd")).Day.ToString();
        if (Donate_Day.Length == 1)
        {
            Donate_Day = "0" + Donate_Day;
        }
        string Donate_Hour = DateTime.Now.Hour.ToString();
        if (Donate_Hour.Length == 1)
        {
            Donate_Hour = "0" + Donate_Hour;
        }
        string Donate_Minute = DateTime.Now.Minute.ToString();
        if (Donate_Minute.Length == 1)
        {
            Donate_Minute = "0" + Donate_Minute;
        }
        string Donate_Second = DateTime.Now.Second.ToString();
        if (Donate_Second.Length == 1)
        {
            Donate_Second = "0" + Donate_Second;
        }

        //--------------------------------------------                
        for (int i = 0; i < (Pledge_Ids.Length) - 1; i++)
        {
            //必要條件(卡號不可為空白)
            //扣款金額一律抓Donate_Amt(因為沒有分期付款的首次扣款金額)
            string strSql = @" Select Pledge_Id, P.Donor_Id, Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, 
                                      Post_SavingsNo,Post_AccountNo,Post_IDNo,Donate_Amt= CAST((REPLACE((Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),'.00',''))as NVARCHAR),D.Donor_Name,Donate_ToDate 
                               From Pledge P Join Donor D On P.Donor_Id = D.Donor_Id 
                               Where Pledge_Id = '" + Pledge_Ids[i] + "' And Status='授權中' And Account_No<>''";

            Dictionary<string, object> dict = new Dictionary<string, object>();            
            DataTable dt;
            DataRow dr;
            dt = NpoDB.GetDataTableS(strSql, dict);
            if (dt.Rows.Count != 0)
            {
                dr = dt.Rows[0];
                row += 1;
                //--------------------------------------------
                //台銀明細資料
                //1.商店代號(銀行別004+統一編號81582054+分店別9001)
                DataLen = 15;
                if (dr_Txt["Transfer_StoreCode"].ToString().Length >= DataLen)
                    content += Util.Left(dr_Txt["Transfer_StoreCode"].ToString().Trim(), DataLen);
                else
                    content += dr_Txt["Transfer_StoreCode"].ToString().Trim();
                //2.商店行業代號
                DataLen = 4;
                content += dr_Txt["Transfer_ServerCode"].ToString().Trim();

                //3.卡號
                DataLen = 19;
                string Account_No = dr["Account_No"].ToString().Trim();
                if (Account_No.Length >= DataLen)
                    content += Util.Left(Account_No, DataLen);
                else
                    content += Account_No + Util.Space(DataLen - Account_No.Length);
                //4.信用卡CSV2
                DataLen = 3;
                string Authorize = dr["Authorize"].ToString().Trim();                
                if (Authorize.Length >= DataLen)
                    content += Util.Left(Authorize, DataLen);
                else
                    content += Authorize + Util.Space(DataLen - Authorize.Length);
                //5.有效年月(年YY月MM)
                DataLen = 4;
                content += dr["Valid_Date"].ToString().Trim();
                //6.交易金額
                DataLen = 12;
                string Donate_Amt = dr["Donate_Amt"].ToString().Trim();
                content += Util.Left("0000000000", DataLen - Donate_Amt.Length - 2) + Donate_Amt + "00";

                //7.回覆碼(空白)
                content += Util.Space(2);
                //8.授權碼(空白)
                content += Util.Space(6);
                //9.保留欄位(空白)
                content += Util.Space(3);

                //10.備註
                int RemarkLen = 30;
                string Donor_Id = dr["Donor_Id"].ToString().Trim();
                Donor_Id = Util.Left("00000000", 8 - Donor_Id.Length) + Donor_Id;
                string Pledge_Id = dr["Pledge_Id"].ToString().Trim();
                Pledge_Id = Util.Left("00000000", 8 - Pledge_Id.Length) + Pledge_Id;
                string Remark = Donor_Id + Pledge_Id;
                content += Remark + Util.Space(RemarkLen - Remark.Length);

                Donate_Total += int.Parse(dr["Donate_Amt"].ToString());
                content += "\r\n";
                //--------------------------------------------
            }
        }
        //--------------------------------------------
        //Tailer
        //1.檔尾
        content += "T";
        //2.資料筆數
        content += Util.Left("000000", 6 - (row.ToString().Length)) + row.ToString();
        //3.交易總額
        content += Util.Left("0000000000", 10 - (Donate_Total.ToString().Length)) + Donate_Total.ToString() + "00";
        //4.交易日期(YYYYMMDD)
        content += Donate_Year + Donate_Month + Donate_Day;
        //5.交易時間(HHMMSS)
        content += Donate_Hour + Donate_Minute + Donate_Second;

        //6.保留欄位
        content += Util.Space(35);
        //7.備註
        content += Util.Space(30);

        content += "\r\n";
        //--------------------------------------------
        return content;
    }

    public System.Security.AccessControl.FileSecurity ACL { get; set; }
}