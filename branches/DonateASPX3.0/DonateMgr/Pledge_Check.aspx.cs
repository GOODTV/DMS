using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Security;
using System.IO;
using System.Data;
using System.Text;

public partial class DonateMgr_Pledge_Check : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            HFD_DonateDate.Value = Util.GetQueryString("DonateDate");
            HFD_PledgeBatchFileName.Value = Util.GetQueryString("file");
        }
        //lblGridList.Text = Session["List"].ToString();  
        //2014/04/19 新增 by Ian_Kao
        LoadFormData();
    }
    protected void btnInput_Click(object sender, EventArgs e)
    {
        string Pledge_Id = "";
        string strId = "";
        foreach (string str in Request.Form.Keys)
        {
            if (str.StartsWith("PledgeId"))
            {

                strId = str.Split('_')[1];
                Pledge_Id += strId + " ";
            }
        }
        HFD_Pledge_Id.Value = Pledge_Id;
        string[] Pledge_Ids = Pledge_Id.Split(' ');//把Pledge_Id分開獨立出來
        int count = Pledge_Ids.Length - 1;//計算Pledge_Id共有幾個 
        HFD_Total_Row.Value = count.ToString();
        int amount = 0;
        string Total_Amount = "";
        int count_close = 0;//存放多少筆轉帳因關帳日期已過無法轉帳
        int amount_close = 0;//存放多少金額轉帳因關帳日期已過無法轉帳
        string Total_Amount_close = "";

        string strSql = "";
        DataTable dt;
        DataRow dr;

        //將扣款資料轉換成捐款資料
        for (int i = 0; i < count; i++)
        {
            strSql = "select * from Pledge where Pledge_Id = '" + Pledge_Ids[i] + "'";
            dt = NpoDB.QueryGetTable(strSql);
            dr = dt.Rows[0];
            //20140519 新增by Ian 確認關帳日期是否過期
            //新增捐款資料
            bool flag = true;
            try
            {
                //20140826關帳判斷依據改為「扣款日期」
                //flag = Util.Get_Close("1", SessionInfo.DeptID, dr["Next_DonateDate"].ToString(), SessionInfo.UserID);
                flag = Util.Get_Close("1", SessionInfo.DeptID, HFD_DonateDate.Value, SessionInfo.UserID);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (flag == false)//未關帳,執行轉帳作業,紀錄金額
            {
                Donate_AddNew(dr);
                if (Convert.ToInt32(Math.Floor(double.Parse(dr["Donate_Amt"].ToString()))) != 0)
                {
                    amount += Convert.ToInt32(Math.Floor(double.Parse(dr["Donate_Amt"].ToString())));
                }
                else
                {
                    amount += 0;
                }
            }
            else//已關帳,紀錄筆數和金額
            {
                if (Convert.ToInt32(Math.Floor(double.Parse(dr["Donate_Amt"].ToString()))) != 0)
                {
                    amount_close += Convert.ToInt32(Math.Floor(double.Parse(dr["Donate_Amt"].ToString())));
                }
                else
                {
                    amount_close += 0;
                }
                count_close += 1;
            }
        }
        //20140519 修改by Ian 若是當次轉帳有一筆以上確認轉入才變動Dept資料表中的最後轉帳日修改
        if (count - count_close > 0)
        {
            //Dept資料表中的最後轉帳日修改
            Dept_Edit();
            //扣款日期變動
            strSql = "Select * From DEPT Where DeptId='" + SessionInfo.DeptID + "'";
            dt = NpoDB.QueryGetTable(strSql);
            dr = dt.Rows[0];
            string Transfer_Date = dr["Transfer_Date"].ToString(); //DateTime.Parse(dr["Transfer_Date"].ToString()).ToString("yyyy/MM/dd");
            string LastTransfer_Date_S = DateTime.Parse(dr["LastTransfer_Date"].ToString()).ToString("yyyy/MM/dd");
            DateTime LastTransfer_Date = Convert.ToDateTime(LastTransfer_Date_S);
            if (HFD_DonateDate.Value != "")
            {
                if (LastTransfer_Date != null)
                {
                    HFD_DonateDate.Value = DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).ToString("yyyy/MM/dd");
                }
                else
                {
                    HFD_DonateDate.Value = DateTime.Now.Year + "/" + DateTime.Now.Month + "/" + Transfer_Date;
                }
            }
            else
            {
                HFD_DonateDate.Value = DateTime.Now.ToString("yyyy/MM/dd");
            }
        }
        
        Total_Amount = string.Format("{0:##,##0}", Convert.ToDecimal(amount));
        Total_Amount_close = string.Format("{0:##,##0}", Convert.ToDecimal(amount_close));
        //將PLEDGE_SEND_RETURN中資料刪除
        strSql = "Delete from Pledge_Send_Return where Pledge_Id>0";
        NpoDB.ExecuteSQLS(strSql, null);
        //2014/04/19 修改 by Ian_Kao
        //Page.Controls.Add(new LiteralControl("<script>alert(\"" + "授權轉入捐款成功！\\n轉入筆數：" + count + "\\n轉入金額：" + Total_Amount + "\");</script>"));
        //this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript'>ReturnOpener();</script>"); 
        SetSysMsg("授權轉入捐款成功！\\n轉入筆數：" + (count - count_close) + "\\n轉入金額：" + Total_Amount + "\\n授權轉入捐款失敗筆數：" + count_close + "\\n失敗未轉入金額：" + Total_Amount_close);
        Response.Redirect(Util.RedirectByTime("Pledge_Transfer.aspx"));
    }
    public void Donate_AddNew(DataRow dr)
    {
        //20140519 修改by Ian 調整確認關帳時間的判斷 
        //bool check_close = Util.Get_Close("1", SessionInfo.DeptID, dr["Next_DonateDate"].ToString(), SessionInfo.UserID);
        //if (check_close == false)
        //{
            //bool flag = false;
            //try
            //{
                string strSql = "insert into  Donate\n";
                strSql += "( Donor_Id, Pledge_Id, Donate_Date, Donate_Payment, Donate_Purpose, Donate_Purpose_Type, Invoice_Type, Donate_Amt, Donate_Fee, Donate_Accou,\n";
                strSql += " Dept_Id, Invoice_Title, Issue_Type, Issue_Type_Keep, Invoice_Pre, Invoice_No, Request_Date, Accoun_Bank, Accoun_Date,\n";
                strSql += " Donate_Type, Donation_NumberNo, Donation_SubPoenaNo, Accounting_Title, InvoceSend_Date, Act_Id,\n";
                strSql += " Comment, Invoice_PrintComment, Export, Create_Date, Create_DateTime, Create_User, Create_IP,\n";
                strSql += " Check_No, Check_ExpireDate, Card_Bank, Card_Type, Account_No, Valid_Date, Card_Owner, Owner_IDNo, Relation, \n";
                strSql += " Authorize, Post_Name, Post_IDNo, Post_SavingsNo, Post_AccountNo,P_BANK, P_RCLNO, P_PID, PledgeBatchFileName) Values\n";
                strSql += "( @Donor_Id,@Pledge_Id,@Donate_Date,@Donate_Payment,@Donate_Purpose,@Donate_Purpose_Type,@Invoice_Type,@Donate_Amt,@Donate_Fee,@Donate_Accou,\n";
                strSql += "@Dept_Id,@Invoice_Title,@Issue_Type,@Issue_Type_Keep,@Invoice_Pre,@Invoice_No,@Request_Date,@Accoun_Bank,@Accoun_Date,\n";
                strSql += "@Donate_Type,@Donation_NumberNo,@Donation_SubPoenaNo,@Accounting_Title,@InvoceSend_Date,@Act_Id,\n";
                strSql += "@Comment,@Invoice_PrintComment,@Export,@Create_Date,@Create_DateTime,@Create_User,@Create_IP,\n";
                strSql += "@Check_No,@Check_ExpireDate,@Card_Bank,@Card_Type,@Account_No,@Valid_Date,@Card_Owner,@Owner_IDNo,@Relation, \n";
                strSql += "@Authorize,@Post_Name,@Post_IDNo,@Post_SavingsNo,@Post_AccountNo,@P_BANK,@P_RCLNO,@P_PID,@PledgeBatchFileName) ";
                strSql += "\n";
                strSql += "select @@IDENTITY";


                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Donor_Id", dr["Donor_Id"].ToString());
                dict.Add("Pledge_Id", dr["Pledge_Id"].ToString());
                dict.Add("Donate_Date", HFD_DonateDate.Value);
                dict.Add("Donate_Payment", dr["Donate_Payment"].ToString());
                dict.Add("Donate_Purpose", dr["Donate_Purpose"].ToString());
                dict.Add("Donate_Purpose_Type", dr["Donate_Purpose_Type"].ToString());
                dict.Add("Invoice_Type", dr["Invoice_Type"].ToString());
                if (Convert.ToInt32(Math.Floor(double.Parse(dr["Donate_Amt"].ToString()))) != 0)
                {
                    dict.Add("Donate_Amt", Decimal.Parse(dr["Donate_Amt"].ToString()));
                    dict.Add("Donate_Fee", "0");
                    dict.Add("Donate_Accou", Decimal.Parse(dr["Donate_Amt"].ToString()));
                }
                else
                {
                    dict.Add("Donate_Amt", "0");
                    dict.Add("Donate_Fee", "0");
                    dict.Add("Donate_Accou", "0");
                }
                //dict.Add("Donate_Forign", 0);
                if (dr["Donate_Payment"].ToString() == "信用卡授權書(一般)" || dr["Donate_Payment"].ToString() == "信用卡授權書(聯信)")
                {
                    dict.Add("Check_No", "");
                    dict.Add("Check_ExpireDate", "");

                    dict.Add("Card_Bank", dr["Card_Bank"].ToString());
                    dict.Add("Card_Type", dr["Card_Type"].ToString());
                    dict.Add("Account_No", dr["Account_No"].ToString());
                    dict.Add("Valid_Date", dr["Valid_Date"].ToString());
                    dict.Add("Card_Owner", dr["Card_Owner"].ToString());
                    dict.Add("Owner_IDNo", dr["Owner_IDNo"].ToString());
                    dict.Add("Relation", dr["Relation"].ToString());
                    dict.Add("Authorize", dr["Authorize"].ToString());

                    dict.Add("Post_Name", "");
                    dict.Add("Post_IDNo", "");
                    dict.Add("Post_SavingsNo", "");
                    dict.Add("Post_AccountNo", "");

                    dict.Add("P_BANK", "");
                    dict.Add("P_RCLNO", "");
                    dict.Add("P_PID", "");
                }
                else if (dr["Donate_Payment"].ToString() == "郵局帳戶轉帳授權書")
                {
                    dict.Add("Check_No", "");
                    dict.Add("Check_ExpireDate", "");

                    dict.Add("Card_Bank", "");
                    dict.Add("Card_Type", "");
                    dict.Add("Account_No", "");
                    dict.Add("Valid_Date", "");
                    dict.Add("Card_Owner", "");
                    dict.Add("Owner_IDNo", "");
                    dict.Add("Relation", "");
                    dict.Add("Authorize", "");

                    dict.Add("Post_Name", dr["Post_Name"].ToString());
                    dict.Add("Post_IDNo", dr["Post_IDNo"].ToString());
                    dict.Add("Post_SavingsNo", dr["Post_SavingsNo"].ToString());
                    dict.Add("Post_AccountNo", dr["Post_AccountNo"].ToString());

                    dict.Add("P_BANK", "");
                    dict.Add("P_RCLNO", "");
                    dict.Add("P_PID", "");
                }
                //2014/5/15新增ACH轉帳授權欄位
                else if (dr["Donate_Payment"].ToString() == "ACH轉帳授權書")
                {
                    dict.Add("Check_No", "");
                    dict.Add("Check_ExpireDate", "");

                    dict.Add("Card_Bank", "");
                    dict.Add("Card_Type", "");
                    dict.Add("Account_No", "");
                    dict.Add("Valid_Date", "");
                    dict.Add("Card_Owner", "");
                    dict.Add("Owner_IDNo", "");
                    dict.Add("Relation", "");
                    dict.Add("Authorize", "");

                    dict.Add("Post_Name", "");
                    dict.Add("Post_IDNo", "");
                    dict.Add("Post_SavingsNo", "");
                    dict.Add("Post_AccountNo", "");
                    //收受行代號
                    dict.Add("P_BANK", dr["P_BANK"].ToString());
                    //收受者帳號
                    dict.Add("P_RCLNO", dr["P_RCLNO"].ToString());
                    //收受者身分證/統編
                    dict.Add("P_PID", dr["P_PID"].ToString());
                }
                else
                {
                    dict.Add("Check_No", "");
                    dict.Add("Check_ExpireDate", "");

                    dict.Add("Card_Bank", "");
                    dict.Add("Card_Type", "");
                    dict.Add("Account_No", "");
                    dict.Add("Valid_Date", "");
                    dict.Add("Card_Owner", "");
                    dict.Add("Owner_IDNo", "");
                    dict.Add("Relation", "");
                    dict.Add("Authorize", "");

                    dict.Add("Post_Name", "");
                    dict.Add("Post_IDNo", "");
                    dict.Add("Post_SavingsNo", "");
                    dict.Add("Post_AccountNo", "");

                    dict.Add("P_BANK", "");
                    dict.Add("P_RCLNO", "");
                    dict.Add("P_PID", "");
                }
                dict.Add("Dept_Id", SessionInfo.DeptID);
                dict.Add("Invoice_Title", dr["Invoice_Title"].ToString());
                //收據編號
                dict.Add("Invoice_Pre", "A");
                //******設定收據編號******//
                string strSql2 = @"select isnull(MAX(Invoice_No),'') as Invoice_No from Donate
                    where Invoice_No like '%" + HFD_DonateDate.Value.Replace("/", "") + "%'";
                //****執行查詢流水號語法****//
                DataTable dt2 = NpoDB.QueryGetTable(strSql2);
                string Invoice_No = "";
                if (dt2.Rows.Count > 0)
                {
                    string Invoice_No_Value = dt2.Rows[0]["Invoice_No"].ToString();

                    if (Invoice_No_Value == "")
                    {
                        Invoice_No = HFD_DonateDate.Value.Replace("/", "") + "0001";
                    }
                    else
                    {
                        Invoice_No = (Convert.ToInt64(Invoice_No_Value) + 1).ToString();
                    }
                }
                //************************//
                dict.Add("Invoice_No", Invoice_No);
                dict.Add("Issue_Type", "");
                dict.Add("Issue_Type_Keep", "");
                dict.Add("Request_Date", null);
                dict.Add("Accoun_Bank", dr["Accoun_Bank"].ToString());
                dict.Add("Accoun_Date", null);
                dict.Add("Donate_Type", dr["Donate_Type"].ToString());
                dict.Add("Donation_NumberNo", "");
                dict.Add("Donation_SubPoenaNo", "");
                dict.Add("Accounting_Title", dr["Accounting_Title"].ToString());
                dict.Add("InvoceSend_Date", null);
                if (dr["Accounting_Title"].ToString() != "")
                {
                    dict.Add("Act_Id", dr["Act_Id"].ToString());
                }
                else
                {
                    dict.Add("Act_Id", "");
                }
                dict.Add("Comment", dr["Comment"].ToString());
                dict.Add("Invoice_PrintComment", "");
                dict.Add("Export", "N");
                dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
                dict.Add("Create_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
                dict.Add("Create_User", SessionInfo.UserName);
                dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
                //2015/1/8 增加授權轉收據的批次檔案
                dict.Add("PledgeBatchFileName", HFD_PledgeBatchFileName.Value);
                NpoDB.ExecuteSQLS(strSql, dict);
                //flag = true;
            //}
            //catch (Exception ex)
            //{
            //    throw ex;
            //}
            //if (flag == true)
            //{
                //Donor資料表中的捐款資料增加
                Donor_Edit(dr);
                //Pledge資料表中的扣款資料修改
                Pledge_Edit(dr);
            //}
        //}//
        //else
        //{
        //    ShowSysMsg("您輸入的捐款日期『" + dr["Next_DonateDate"].ToString() + "』 已關帳無法新增 ！");
        //}
    }
    public void Dept_Edit()
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = " update Dept set ";
        strSql += " LastTransfer_Date = @LastTransfer_Date";
        strSql += " where DeptId = @DeptId";

        dict.Add("LastTransfer_Date", HFD_DonateDate.Value);
        dict.Add("DeptId", SessionInfo.DeptID);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void Donor_Edit(DataRow dr_Data)
    {
        //先抓出第一次/最近一次的捐款日期、捐款次數和捐款總額
        string Begin_DonateDate, Last_DonateDate, Donate_Total_S;
        int Donate_No;
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;
        //****變數設定****//
        uid = dr_Data["Donor_Id"].ToString();
        //****設定查詢****//
        strSql = " select *  from Donor where Donor_Id='" + uid + "'";
        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];

        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        string strSql2 = "";
        if (dr["Begin_DonateDate"].ToString() == "" || DateTime.Parse(dr["Begin_DonateDate"].ToString()).ToString("yyyy/MM/dd") == "1900/01/01")
        {
            //初次捐款
            strSql2 = " update Donor set ";
            strSql2 += "  Begin_DonateDate = @Begin_DonateDate";
            strSql2 += ", Last_DonateDate = @Last_DonateDate";
            strSql2 += ", Donate_No = @Donate_No";
            strSql2 += ", Donate_Total = @Donate_Total";
            strSql2 += " where Donor_Id = @Donor_Id";

            dict2.Add("Begin_DonateDate", HFD_DonateDate.Value);
            dict2.Add("Last_DonateDate", HFD_DonateDate.Value);
            dict2.Add("Donate_No", "1");
            //之前寫成Donate_Amt 應該是Donate_Total
            if (dr["Donate_Total"].ToString() != "0")
            {
                dict2.Add("Donate_Total", Decimal.Parse(dr_Data["Donate_Amt"].ToString()));
            }
            else
            {
                dict2.Add("Donate_Total", "0");
            }
            dict2.Add("Donor_Id", dr_Data["Donor_Id"].ToString());
        }
        else
        {
            Begin_DonateDate = dr["Begin_DonateDate"].ToString();
            Last_DonateDate = dr["Last_DonateDate"].ToString();
            Donate_No = Int32.Parse(dr["Donate_No"].ToString());
            Donate_Total_S = (Convert.ToInt64(dr["Donate_Total"])).ToString();
            int Donate_Total = Int32.Parse(Donate_Total_S);
            //更新以上欄位

            strSql2 = " update Donor set ";
            strSql2 += " Donate_No = @Donate_No";
            strSql2 += ", Donate_Total = @Donate_Total";
            //if (DateTime.Parse(DateTime.Parse(Last_DonateDate).ToString("yyyy/MM/dd")) < DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")))
            if (DateTime.Parse(DateTime.Parse(Last_DonateDate).ToString("yyyy/MM/dd")) < DateTime.Parse(HFD_DonateDate.Value))
            {
                strSql2 += ", Last_DonateDate = @Last_DonateDate";
                //dict2.Add("Last_DonateDate", DateTime.Now.ToString("yyyy-MM-dd"));
                dict2.Add("Last_DonateDate", HFD_DonateDate.Value);
            }
            strSql2 += " where Donor_Id = @Donor_Id";

            dict2.Add("Donate_No", Donate_No + 1);
            if (dr_Data["Donate_Amt"].ToString() != "0")
            {
                dict2.Add("Donate_Total", Donate_Total + Decimal.Parse(dr_Data["Donate_Amt"].ToString()));
            }
            else
            {
                dict2.Add("Donate_Total", Donate_Total);
            }

            dict2.Add("Donor_Id", dr_Data["Donor_Id"].ToString());
        }
        NpoDB.ExecuteSQLS(strSql2, dict2);
    }
    public void Pledge_Edit(DataRow dr_Data)
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;
        //****變數設定****//
        uid = dr_Data["Pledge_Id"].ToString();
        //****設定查詢****//
        strSql = " select * from Pledge where Pledge_Id='" + uid + "'";
        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];

        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        string strSql2 = " update Pledge set ";
        if (DateTime.Parse(dr_Data["Donate_ToDate"].ToString()).Year < DateTime.Parse(HFD_DonateDate.Value).Year || (DateTime.Parse(dr_Data["Donate_ToDate"].ToString()).Year == DateTime.Parse(HFD_DonateDate.Value).Year && DateTime.Parse(dr_Data["Donate_ToDate"].ToString()).Month <= DateTime.Parse(HFD_DonateDate.Value).Month))
        {
            strSql2 += " Status = @Status";
            dict2.Add("Status", "停止");
            //20140515 新增 by Ian_Kao 紀錄最後扣款日期
            strSql2 += ",transfer_Date = @transfer_Date";
            dict2.Add("transfer_Date", HFD_DonateDate.Value);
        }
        else
        {
            if (dr_Data["Donate_Period"].ToString() == "單筆")
            {
                strSql2 += " Status = @Status";
                dict2.Add("Status", "停止");
                //20140515 新增 by Ian_Kao 紀錄最後扣款日期
                strSql2 += ",transfer_Date = @transfer_Date";
                dict2.Add("transfer_Date", HFD_DonateDate.Value);
            }
            else if (dr_Data["Donate_Period"].ToString() == "月繳")
            {
                strSql2 += "transfer_Date = @transfer_Date";
                //---------------------------------------------------------------------------------------
                //20140515 修改 by Ian_Kao 避免Next_DonateDate可能會是6/31這類狀況發生所做特別處理
                //if (DateTime.Parse(dr_Data["Donate_ToDate"].ToString()) < DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day))
                string Next_DonateDate_string = Util.DateTime2String(DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day, DateType.yyyyMMdd, EmptyType.ReturnEmpty);
                DateTime Next_DonateDate;
                if (Next_DonateDate_string == "")
                {
                    Next_DonateDate = DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).AddMonths(1).Day);
                }
                else
                {
                    Next_DonateDate = DateTime.Parse(Next_DonateDate_string);
                }
                if (DateTime.Parse(dr_Data["Donate_ToDate"].ToString()) < Next_DonateDate)
                //---------------------------------------------------------------------------------------
                {
                    strSql2 += " ,Status = @Status";
                    dict2.Add("Status", "停止");
                }
                //20140515 修改 by Ian_Kao Donate_Payment有可能是「信用卡授權書(一般)」或是「信用卡授權書(聯信)」
                //else if (dr_Data["Donate_Payment"].ToString() == "信用卡授權書(一般)" && (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Year || (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) == DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Year && (int.Parse(dr_Data["Valid_Date"].ToString().Substring(0, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Month))))
                else if ((dr_Data["Donate_Payment"].ToString() == "信用卡授權書(一般)" || dr_Data["Donate_Payment"].ToString() == "信用卡授權書(聯信)") && (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Year || (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) == DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Year && (int.Parse(dr_Data["Valid_Date"].ToString().Substring(0, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Month))))
                {
                    strSql2 += " ,Status = @Status";
                    dict2.Add("Status", "停止");
                }
                else
                {
                    strSql2 += ",Next_DonateDate = @Next_DonateDate";
                    //dict2.Add("Next_DonateDate", DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(1).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day));
                    dict2.Add("Next_DonateDate", Next_DonateDate);
                }
                dict2.Add("transfer_Date", HFD_DonateDate.Value);
            }
            else if (dr_Data["Donate_Period"].ToString() == "季繳")
            {
                strSql2 += "transfer_Date = @transfer_Date";
                //---------------------------------------------------------------------------------------
                //20140515 修改 by Ian_Kao 避免Next_DonateDate可能會是6/31這類狀況發生所做特別處理
                //if (DateTime.Parse(dr_Data["Donate_ToDate"].ToString()) < DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day))
                string Next_DonateDate_string = Util.DateTime2String(DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day, DateType.yyyyMMdd, EmptyType.ReturnEmpty);
                DateTime Next_DonateDate;
                if (Next_DonateDate_string == "")
                {
                    Next_DonateDate = DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).AddMonths(3).Day);
                }
                else
                {
                    Next_DonateDate = DateTime.Parse(Next_DonateDate_string);
                }
                if (DateTime.Parse(dr_Data["Donate_ToDate"].ToString()) < Next_DonateDate)
                //---------------------------------------------------------------------------------------
                {
                    strSql2 += " ,Status = @Status";
                    dict2.Add("Status", "停止");
                }
                //20140515 修改 by Ian_Kao Donate_Payment有可能是「信用卡授權書(一般)」或是「信用卡授權書(聯信)」
                //else if (dr_Data["Donate_Payment"].ToString() == "信用卡授權書(一般)" && (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Year || (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) == DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Year && (int.Parse(dr_Data["Valid_Date"].ToString().Substring(0, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Month))))
                else if ((dr_Data["Donate_Payment"].ToString() == "信用卡授權書(一般)" || dr_Data["Donate_Payment"].ToString() == "信用卡授權書(聯信)") && (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Year || (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) == DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Year && (int.Parse(dr_Data["Valid_Date"].ToString().Substring(0, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Month))))
                {
                    strSql2 += " ,Status = @Status";
                    dict2.Add("Status", "停止");
                }
                else
                {
                    strSql2 += ",Next_DonateDate = @Next_DonateDate";
                    //dict2.Add("Next_DonateDate", DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(3).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day));
                    dict2.Add("Next_DonateDate", Next_DonateDate);
                }
                dict2.Add("transfer_Date", HFD_DonateDate.Value);
            }
            else if (dr_Data["Donate_Period"].ToString() == "半年繳")
            {
                strSql2 += "transfer_Date = @transfer_Date";
                //---------------------------------------------------------------------------------------
                //20140515 修改 by Ian_Kao 避免Next_DonateDate可能會是6/31這類狀況發生所做特別處理
                //if (DateTime.Parse(dr_Data["Donate_ToDate"].ToString()) < DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day))
                string Next_DonateDate_string = Util.DateTime2String(DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day, DateType.yyyyMMdd, EmptyType.ReturnEmpty);
                DateTime Next_DonateDate;
                if (Next_DonateDate_string == "")
                {
                    Next_DonateDate = DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).AddMonths(6).Day);
                }
                else
                {
                    Next_DonateDate = DateTime.Parse(Next_DonateDate_string);
                }
                if (DateTime.Parse(dr_Data["Donate_ToDate"].ToString()) < Next_DonateDate)
                //---------------------------------------------------------------------------------------
                {
                    strSql2 += " ,Status = @Status";
                    dict2.Add("Status", "停止");
                }
                //20140515 修改 by Ian_Kao Donate_Payment有可能是「信用卡授權書(一般)」或是「信用卡授權書(聯信)」
                //else if (dr_Data["Donate_Payment"].ToString() == "信用卡授權書(一般)" && (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Year || (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) == DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Year && (int.Parse(dr_Data["Valid_Date"].ToString().Substring(0, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Month))))
                else if ((dr_Data["Donate_Payment"].ToString() == "信用卡授權書(一般)" || dr_Data["Donate_Payment"].ToString() == "信用卡授權書(聯信)") && (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Year || (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) == DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Year && (int.Parse(dr_Data["Valid_Date"].ToString().Substring(0, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Month))))
                {
                    strSql2 += " ,Status = @Status";
                    dict2.Add("Status", "停止");
                }
                else
                {
                    strSql2 += ",Next_DonateDate = @Next_DonateDate";
                    //dict2.Add("Next_DonateDate", DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(6).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day));
                    dict2.Add("Next_DonateDate", Next_DonateDate);
                }
                dict2.Add("transfer_Date", HFD_DonateDate.Value);
            }
            else if (dr_Data["Donate_Period"].ToString() == "年繳")
            {
                strSql2 += "transfer_Date = @transfer_Date";
                //---------------------------------------------------------------------------------------
                //20140515 修改 by Ian_Kao 避免Next_DonateDate可能會是6/31這類狀況發生所做特別處理
                //if (DateTime.Parse(dr_Data["Donate_ToDate"].ToString()) < DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day))
                string Next_DonateDate_string = Util.DateTime2String(DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day, DateType.yyyyMMdd, EmptyType.ReturnEmpty);
                DateTime Next_DonateDate;
                if (Next_DonateDate_string == "")
                {
                    Next_DonateDate = DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).AddMonths(12).Day);
                }
                else
                {
                    Next_DonateDate = DateTime.Parse(Next_DonateDate_string);
                }
                if (DateTime.Parse(dr_Data["Donate_ToDate"].ToString()) < Next_DonateDate)
                //---------------------------------------------------------------------------------------
                {
                    strSql2 += " ,Status = @Status";
                    dict2.Add("Status", "停止");
                }
                //20140515 修改 by Ian_Kao Donate_Payment有可能是「信用卡授權書(一般)」或是「信用卡授權書(聯信)」
                //else if (dr_Data["Donate_Payment"].ToString() == "信用卡授權書(一般)" && (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Year || (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) == DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Year && (int.Parse(dr_Data["Valid_Date"].ToString().Substring(0, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Month))))
                else if ((dr_Data["Donate_Payment"].ToString() == "信用卡授權書(一般)" || dr_Data["Donate_Payment"].ToString() == "信用卡授權書(聯信)") && (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Year || (int.Parse("20" + dr_Data["Valid_Date"].ToString().Substring(2, 2)) == DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Year && (int.Parse(dr_Data["Valid_Date"].ToString().Substring(0, 2)) < DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Month))))
                {
                    strSql2 += " ,Status = @Status";
                    dict2.Add("Status", "停止");
                }
                else
                {
                    strSql2 += ",Next_DonateDate = @Next_DonateDate";
                    //dict2.Add("Next_DonateDate", DateTime.Parse(DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Year + "/" + DateTime.Parse(HFD_DonateDate.Value).AddMonths(12).Month + "/" + DateTime.Parse(dr_Data["Next_DonateDate"].ToString()).Day));
                    dict2.Add("Next_DonateDate", Next_DonateDate);
                }
                dict2.Add("transfer_Date", HFD_DonateDate.Value);
            }
        }
        strSql2 += " where Pledge_Id = @Pledge_Id";
        dict2.Add("Pledge_Id", dr_Data["Pledge_Id"].ToString());
        NpoDB.ExecuteSQLS(strSql2, dict2);
    }
    //-------------------------------------------------------------------------
    //2014/04/19 新增 by Ian_Kao
    public void LoadFormData()
    {
        string Pledge_Id = Session["Pledge_Id"].ToString();
        int count = Pledge_Id.Split(' ').Count() - 1;//最後有一個空格會造成判斷錯誤
        string[] Pledge_Ids = new string[count];
        for (int i = 0; i < count; i++)
        {
            Pledge_Ids[i] = Pledge_Id.Split(' ')[i];
        }
        bool first = true;
        lblGridList.Text = "";//先清空再填入
        load_list(true, first, Pledge_Ids);//有勾選
        load_list(false, first, Pledge_Ids); //沒有勾選
    }
    private void load_list(bool check, bool first, string[] pledge_id)
    {
        
        string strSql = Session["strSql"].ToString();
        if (check == true)
        {
            strSql += " AND Pledge_Id IN ( ";
            //2014/06/04 新增 by Ian_Kao 計算總筆數和總金額
            int count = pledge_id.Length;//總筆數
            string strSql2 = @"
                               select REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum(Donate_Amt)),1),'.00','')  as 總金額
                               From PLEDGE P 
                               where 1=1";
            strSql2 += " AND Pledge_Id IN ( ";
            for (int i = 0; i < pledge_id.Length; i++)
            {
                if (i == pledge_id.Length - 1)
                {
                    strSql2 += "'" + pledge_id[i] + "'";
                }
                else
                {
                    strSql2 += "'" + pledge_id[i] + "',";
                }
            }
            strSql2 += " ) ";
            DataTable dt2 = NpoDB.GetDataTableS(strSql2, null);
            DataRow dr = dt2.Rows[0];
            lblCount_Amount.Text = "共勾選了 " + count + " 筆資料，總金額為 " + dr["總金額"].ToString() + " 元"
                + "<br/><font color='blue'>匯入台銀回覆檔檔名：" + HFD_PledgeBatchFileName.Value + "</font>";
        }
        else
        {
            strSql += " AND Pledge_Id NOT IN ( ";
        }
        for (int i = 0; i < pledge_id.Length; i++)
        {
            if (i == pledge_id.Length - 1)
            {
                strSql += "'" + pledge_id[i] + "'";
            }
            else
            {
                strSql += "'" + pledge_id[i] + "',";
            }
        }
        strSql += " ) ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (check == true && dt.Rows.Count == 0)
        {
            lblGridList.Text = "";
            return;
        }
        else if (check == false && dt.Rows.Count == 0)
        {
            return;
        }
        //Grid initial
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("Pledge_Id");
        npoGridView.DisableColumn.Add("Pledge_Id");
        npoGridView.DisableColumn.Add("disabled");
        npoGridView.ShowPage = false;
        //-------------------------------------------------------------------------
        // 2014/4/11 修正標題換成為checkbox
        //NPOGridViewColumn col = new NPOGridViewColumn("選擇");
        NPOGridViewColumn col = new NPOGridViewColumn();
        col.ColumnType = NPOColumnType.Checkbox;
        col.ControlKeyColumn.Add("Pledge_Id");
        if (check == true)//決定是否check
        {
            col.ColumnName.Add("Pledge_Id");
        }
        else
        {
            col.ColumnName.Add("disabled");
        }
        col.ControlText.Add("");
        col.ControlName.Add("PledgeId");
        col.ControlValue.Add("");
        col.ControlId.Add("checkbox");
        if (first == false)//第一次顯示標頭,之後不顯示,但是width會亂掉
        {
            // col.ShowTitle = false;
        }
        // 2014/4/10 修正可點選checkbox
        //col.DisableValue = "0";
        col.DisableValue = "";
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權編號");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權編號");
        if (first == false)
        {
            // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("捐款人");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("捐款人");
        if (first == false)
        {
            // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權方式");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權方式");
        if (first == false)
        {
            // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("扣款金額");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("扣款金額");
        if (first == false)
        {
            // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權起日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權起日");
        if (first == false)
        {
            // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權迄日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權迄日");
        if (first == false)
        {
            // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("轉帳週期");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("轉帳週期");
        if (first == false)
        {
            // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("下次扣款日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("下次扣款日");
        if (first == false)
        {
            // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("有效月年");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("有效月年");
        if (first == false)
        {
            // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("最後扣款日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("最後扣款日");
        if (first == false)
        {
            // col.ShowTitle = false;
        }
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        lblGridList.Text += npoGridView.Render();
        Session["List"] = lblGridList.Text;
    }
}