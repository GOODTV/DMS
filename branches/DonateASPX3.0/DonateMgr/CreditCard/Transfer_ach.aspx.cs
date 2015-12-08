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

public partial class DonateMgr_CreditCard_Transfer_ach : BasePage
{
    string Donate_Year_AD;
    string Donate_Year;
    string Donate_Month;
    string Donate_Day;

    protected void Page_Load(object sender, EventArgs e)
    {
        string Uids = Session["Pledge_Id"].ToString();
        Print_Txt(Uids);
    }
    //---------------------------------------------------------------------------
    private void Print_Txt(string Uids)
    {
        if (Uids == "")
        {
            ShowSysMsg("查無資料!!!");
            return;
        }

        //抓出要轉出的txt格式資料
        string strSql = @"Select * From DONATE_TRANSFER Where Transfer_AspxName='" + Session["Transfer_AspxName"].ToString() + "'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        //輸出檔案名稱
        string  Transfer_FileName = "";
        Transfer_FileName = dr["Transfer_FileName"].ToString();
        Transfer_FileName = Transfer_FileName.Replace(".txt", "");                
        GridList.Text = GetTable(Uids, dr);
        //Add by GoodTV Tanya:匯出檔名加上交易扣款日期
        Transfer_FileName = Transfer_FileName + Donate_Year_AD + Donate_Month + Donate_Day + "01";
        Util.OutputTxt(GridList.Text, "4", Transfer_FileName);
    }
    //---------------------------------------------------------------------------
    private string GetTable(string Uids, DataRow dr_Txt)
    {
        //抓出txt中的資料
        string[] Pledge_Ids = Uids.Split(' ');
        string content = "";
        string P_PCLNO = "00066001039995";
        string P_CID = "81582054";
        //授權筆數
        int row = 0;
        //授權金額
        int Donate_Total = 0;
        //Modify by GoodTV Tanya:交易日期為畫面上的扣款日期
        //交易日期
        //string Donate_Year_AD = DateTime.Now.Year.ToString();
        Donate_Year_AD = (Convert.ToDateTime(Session["Donate_Date"].ToString()).Year - 1911).ToString();
        Donate_Year = Convert.ToDateTime(Session["Donate_Date"].ToString()).Year.ToString();
        if (Donate_Year_AD.Length == 2)
        {
            Donate_Year_AD = "00" + Donate_Year_AD;
        }
        else if (Donate_Year_AD.Length == 3)
        {
            Donate_Year_AD = "0" + Donate_Year_AD;
        }
        //string Donate_Month = DateTime.Now.Month.ToString();
        Donate_Month = Convert.ToDateTime(Session["Donate_Date"].ToString()).Month.ToString();
        if (Donate_Month.Length == 1)
        {
            Donate_Month = "0" + Donate_Month;

        }
        //string Donate_Day = DateTime.Now.Day.ToString();
        Donate_Day = Convert.ToDateTime(Session["Donate_Date"].ToString()).Day.ToString();
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
        //ACH控制首錄
        //1.首錄別
        content += "BOF";
        //2.資料代號
        content += "ACHP01";
        //3.處理日期
        content += Donate_Year_AD + Donate_Month + Donate_Day;
        //4.處理時間
        content += Donate_Hour + Donate_Minute + Donate_Second;
        //5.發送單位代號
        content += dr_Txt["Transfer_StoreCode"];
        //6.接收單位代號
        content += "9990250";
        //7.備用
        content += Util.Space(123);
        content += "\r\n";
        //--------------------------------------------
        for (int i = 0; i < (Pledge_Ids.Length) - 1; i++)
        {
            //必要條件(收受行代號、收受者帳號、收受者身分證/統編不可為空白)
            //扣款金額一律抓Donate_Amt(因為沒有分期付款的首次扣款金額)
            string strSql = @" Select Pledge_Id,Isnull(P.Donor_Id_Old,P.Donor_Id) as Donor_Id, Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, 
                                      Post_SavingsNo,Post_AccountNo,Post_IDNo,P_BANK,P_RCLNO,P_PID,Member_No,Donate_Amt= CAST((REPLACE((Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),'.00',''))as NVARCHAR),D.Donor_Name,Donate_ToDate 
                               From Pledge P Join Donor D On P.Donor_Id = D.Donor_Id 
                               Where Pledge_Id = '" + Pledge_Ids[i] + "' And Status='授權中' And P_BANK<>'' And P_RCLNO<>'' And P_PID<>''";

            Dictionary<string, object> dict = new Dictionary<string, object>();
            //dict.Add(" Pledge_Id", Pledge_Ids[i]);
            DataTable dt;
            DataRow dr;
            dt = NpoDB.GetDataTableS(strSql, dict);
            if (dt.Rows.Count != 0)
            {
                dr = dt.Rows[0];
                row += 1;
                //--------------------------------------------
                //ACH明細資料
                //1.交易型態
                content += "N";
                //2.交易類別
                content += "SD";
                //3.交易代號(529公益捐款)
                content += "529";
                //4.交易序號(6位流水號)
                content += Util.Left("000000",6-(row).ToString().Length) + (row).ToString();
                //5.提出行代號
                content += dr_Txt["Transfer_ServerCode"];
                //6.發動者帳號
                content += P_PCLNO;
                //7.提回行代號
                int P_BANKLen = 7;
                string P_BANK = dr["P_BANK"].ToString();
                if(P_BANK.Length >= P_BANKLen)
                {
                    content += Util.Left(P_BANK, P_BANKLen);
                }
                else
                {
                    content += P_BANK + Util.Space(P_BANKLen - P_BANK.Length);
                }
                //8收受者帳號
                content += Util.Left("00000000000000", 14 - (dr["P_RCLNO"].ToString().Length)) + dr["P_RCLNO"].ToString();
                //9金額
                content += Util.Left("0000000000", 10 - (dr["Donate_Amt"].ToString().Length)) + dr["Donate_Amt"].ToString();
                //10退件理由代號
                content += "00";
                //11提示交換次序
                content += "B";
                //12發動者統一編號
                int P_CIDLen = 10;
                if (P_CID.Length >= P_CIDLen)
                {
                    content += Util.Left(P_CID, P_CIDLen);
                }
                else
                {
                    content += P_CID + Util.Space(P_CIDLen - P_CID.Length);
                }
                //13收受者統一編號
                int P_PIDLen = 10;
                string P_PID = dr["P_PID"].ToString();
                if (P_PID.Length >= P_PIDLen)
                {
                    content += Util.Left(P_PID, P_PIDLen);
                }
                else
                {
                    content += P_PID + Util.Space(P_PIDLen - P_PID.Length);
                }
                //14上市上櫃公司代號
                content += Util.Space(6);
                //15原提示交易日期
                content += "00000000";
                //16原提示交易序號
                content += "000000";
                //17原提示交換次序
                content += Util.Space(1);
                //18用戶號碼
                int Member_NoLen = 20;
                string Member_No = dr["Donor_Id"].ToString();
                if (Member_No.Length >= Member_NoLen)
                {
                    content += Util.Left(Member_No, Member_NoLen);
                }
                else
                {
                    content += Member_No + Util.Space(Member_NoLen - Member_No.Length);
                }
                //19發動者專用區
                content += Util.Space(20);
                //20存摺摘要
                content += Util.Space(10);
                //21備用
                content += Util.Space(2);

                Donate_Total += int.Parse(dr["Donate_Amt"].ToString());
                content += "\r\n";
                //--------------------------------------------
            }
        }
        //--------------------------------------------
        //ACHACH控制尾錄
        //1.尾錄別
        content += "EOF";
        //2.資料代號
        content += "ACHP01";
        //3.處理日期
        content += Donate_Year_AD + Donate_Month + Donate_Day;
        //4.發送單位代號
        content += dr_Txt["Transfer_StoreCode"];
        //5.接收單位代號
        content += "9990250";
        //6.總筆數
        content += Util.Left("00000000",8-(row.ToString().Length)) + row.ToString();
        //7.總金額
        content += Util.Left("0000000000000000", 16 - (Donate_Total.ToString().Length)) + Donate_Total.ToString();
        //8.備用
        content += Util.Space(105);
        content += "\r\n";
        //--------------------------------------------
        return content;
    }
}