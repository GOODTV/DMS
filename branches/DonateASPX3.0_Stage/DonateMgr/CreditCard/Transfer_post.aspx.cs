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

public partial class DonateMgr_CreditCard_Transfer_post : BasePage
{
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
        string OrderNo, Transfer_FileName = "";
        if (dr["Transfer_StoreCode"].ToString() != "")
        {
            //if(dr["Transfer_StoreCode"].ToString().IndexOf(DateTime.Now.Year.ToString()+DateTime.Now.Month.ToString()+DateTime.Now.Day.ToString())>-1)
            if (dr["Transfer_StoreCode"].ToString().IndexOf(DateTime.Now.ToString("yyyyMMdd")) > -1)
            {
                //這段照asp的code自己翻成.net打的
                OrderNo = DateTime.Now.ToString("yyyyMMdd") + Util.Left("00", 2 - (((Util.Right(dr["Transfer_OrderNo"].ToString(), 2) + 1).Length))) + (int.Parse(Util.Right(dr["Transfer_OrderNo"].ToString(), 2)) + 1).ToString();
            }
            else
            {
                OrderNo = DateTime.Now.ToString("yyyyMMdd") + "01";
            }
        }
        else
        {
            OrderNo = DateTime.Now.ToString("yyyyMMdd") + "01";
        }
        Transfer_FileName = dr["Transfer_FileName"].ToString();
        Transfer_FileName = Transfer_FileName.Replace(".txt", "");
        GridList.Text = GetTable(Uids, dr);
        Util.OutputTxt(GridList.Text, "4", Transfer_FileName);
    }
    //---------------------------------------------------------------------------
    private string GetTable(string Uids, DataRow dr_Txt)
    {
        //抓出txt中的資料
        string[] Pledge_Ids = Uids.Split(' ');
        string content = "";
        //授權筆數
        int row = 0;
        //授權金額
        int Donate_Total = 0;
        //資料長度
        int SpaceLen=0;
        //交易日期:交易日期為畫面上的扣款日期
        string Donate_Year_AD = (Convert.ToDateTime(Session["Donate_Date"].ToString()).Year - 1911).ToString();
        string Donate_Year = Convert.ToDateTime(Session["Donate_Date"].ToString()).Year.ToString();
        //string Donate_Year_AD = DateTime.Now.Year.ToString();
        //string Donate_Year = (DateTime.Now.Year - 1911).ToString();
        if (Donate_Year_AD.Length == 2)
        {
            Donate_Year_AD = "0" + Donate_Year_AD;
        }        
        string Donate_Month = Convert.ToDateTime(Session["Donate_Date"].ToString()).Month.ToString();
        if (Donate_Month.Length == 1)
        {
            Donate_Month = "0" + Donate_Month;

        }
        string Donate_Day = Convert.ToDateTime(Session["Donate_Date"].ToString()).Day.ToString();
        if (Donate_Day.Length == 1)
        {
            Donate_Day = "0" + Donate_Day;
        }
        
        for (int i = 0; i < (Pledge_Ids.Length) - 1; i++)
        {
            string strSql = @" Select Pledge_Id,Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, 
                                      Post_SavingsNo,Post_AccountNo,Post_IDNo,Donate_Amt= CAST((REPLACE((Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),'.00',''))as NVARCHAR),D.Donor_Name,Donate_ToDate 
                               From Pledge P Join Donor D On P.Donor_Id = D.Donor_Id 
                               Where Pledge_Id = '" + Pledge_Ids[i] + "' And Status='授權中' And Post_SavingsNo<>'' And Post_AccountNo<>''";

            Dictionary<string, object> dict = new Dictionary<string, object>();
            //dict.Add(" Pledge_Id", Pledge_Ids[i]);
            DataTable dt;
            DataRow dr;
            dt = NpoDB.GetDataTableS(strSql, dict);
            if (dt.Rows.Count != 0)
            {
                dr = dt.Rows[0];
                //--------------------------------------------
                //郵局明細資料
                //1.資料別(明細為1) + 2.存款別(存簿為P,劃撥為G) + 3.事業單位代號 
                content += "1" + "P" + dr_Txt["Transfer_StoreCode"].ToString() ;
                //4.區處站所
                SpaceLen = 4;
                content += Util.Space(SpaceLen);
                //5.繳費日期(民國年月日)
                content += Donate_Year_AD + Donate_Month + Donate_Day;
                //6.保留欄
                SpaceLen = 3;
                content += Util.Space(SpaceLen);
                //7.1局帳
                int SavingsNoLen = 7;
                string Post_SavingsNo = dr["Post_SavingsNo"].ToString();
                if (Post_SavingsNo.ToString().Length >= SavingsNoLen)
                {
                    content += Util.Left(Post_SavingsNo, SavingsNoLen);
                }
                else
                {
                    content += Post_SavingsNo + Util.Space(SavingsNoLen - Post_SavingsNo.Length);
                }
                //7.2帳帳
                int AccountNoLen = 7;
                string Post_AccountNo = dr["Post_AccountNo"].ToString();
                if (Post_AccountNo.ToString().Length >= AccountNoLen)
                {
                    content += Util.Left(Post_AccountNo, AccountNoLen);
                }
                else
                {
                    content += Post_AccountNo + Util.Space(AccountNoLen - Post_AccountNo.Length);
                }
                //7.3劃撥固定值 + 7.4劃撥帳號
                //content += "000000" + Util.Space(8);
                //8.身分證號
                int PostIDNoLen = 10;
                string Post_IDNo = dr["Post_IDNo"].ToString();
                if (Post_IDNo.Length >= PostIDNoLen)
                {
                    content += Util.Left(Post_IDNo, PostIDNoLen);
                }
                else
                {
                    content += Post_IDNo + Util.Space(PostIDNoLen - Post_IDNo.Length);
                }
                //9.繳費金額
                int DonateAmtLen = 9;
                string Donate_Amt = dr["Donate_Amt"].ToString();
                if (Donate_Amt.Length >= DonateAmtLen)
                {
                    content += Util.Left(Donate_Amt, DonateAmtLen) + "00";
                }
                else
                {
                    string StrDonateAmt = "";
                    for (int j = 0; j < DonateAmtLen; j++)
                    {
                        StrDonateAmt += "0";
                    }
                    content += Util.Left(StrDonateAmt, DonateAmtLen - (Donate_Amt.Length)) + Donate_Amt + "00";
                }
                //10用戶編號欄(商家自用)
                int RemarkLen = 10;
                string donor_remark = "";
                string donorremark = dr["Pledge_Id"].ToString();
                int iLen = 0;
                for (int j = 0; j < donorremark.Length; j++)
                {
                    if (Util.Asc(Util.Mid(donorremark, j, 1)) < 0)
                    {
                        iLen += 2;
                    }
                    else
                    {
                        iLen += 1;
                    }
                    if (iLen <= RemarkLen)
                    {
                        donor_remark = donor_remark + Util.Mid(donorremark, j, 1);
                    }
                    else
                    {
                        break;
                    }
                }
                content += donor_remark;
                if (iLen < RemarkLen)
                {
                    content += Util.Space(RemarkLen - iLen);
                }
                //11.事業單位使用欄 + 12.用戶編號列印記號
                content += Util.Left("0000000000", 10 - dr_Txt["Transfer_StoreCode"].ToString().Length) + dr_Txt["Transfer_StoreCode"].ToString() + "1";
                //13.保留欄
                SpaceLen = 1;
                content += Util.Space(SpaceLen);
                //14.非連線記號
                SpaceLen = 1;
                content += Util.Space(SpaceLen);
                //15.變更存簿局號記號
                SpaceLen = 1;
                content += Util.Space(SpaceLen);
                //16.狀況代號 
                SpaceLen = 2;
                content += Util.Space(SpaceLen);
                //17.繳費月份(民國年月)
                content += Donate_Year_AD + Donate_Month;
                // + 18.保留欄
                SpaceLen = 5;
                content += Util.Space(SpaceLen);
                content += "\r\n";

                row += 1;
                Donate_Total += int.Parse(dr["Donate_Amt"].ToString());                
            }            
        }
        //郵局總數資料 
        //1.資料別(總數)
        content += "2";
        //2.存款別
        SpaceLen = 1;
        content += Util.Space(SpaceLen);
        //3.事業單位代號
        content += dr_Txt["Transfer_StoreCode"].ToString();
        //4.區處站所
        SpaceLen = 4;
        content += Util.Space(SpaceLen);
        //5.繳費日期
        content += Donate_Year_AD + Donate_Month + Donate_Day;
        //6.保留欄
        SpaceLen = 3;
        content += Util.Space(SpaceLen);
        //7.總件數
        int DonateRowLen = 7;
        if (row.ToString().Length >= DonateRowLen)
        {
            content += Util.Left(row.ToString(), DonateRowLen);
        }
        else
        {
            string StrDonateRow = "";
            for (int j = 0; j < DonateRowLen; j++)
            {
                StrDonateRow += "0";
            }
            content += Util.Left(StrDonateRow, DonateRowLen - (row.ToString().Length)) + row.ToString();
        }
        //8.總金額
        int DonateTotalLen = 13;
        if (Donate_Total.ToString().Length >= DonateTotalLen)
        {
            content += Util.Left(Donate_Total.ToString(), DonateTotalLen);
        }
        else
        {
            string StrDonateTotal = "";
            for (int j = 0; j < DonateTotalLen; j++)
            {
                StrDonateTotal += "0";
            }
            // 2014/8/13 修正格式
            //content += Util.Left(StrDonateTotal, DonateTotalLen - (Donate_Total.ToString().Length)) + Donate_Total.ToString();
            content += Util.Right(StrDonateTotal + Donate_Total.ToString() + "00", DonateTotalLen);
        }
        //9.保留欄
        SpaceLen = 16;
        content += Util.Space(SpaceLen);
        //10.成功件數
        DonateRowLen = 7;
        for (int j = 0; j < DonateRowLen; j++)
        {
            content += "0";
        }
        //11.成功金額
        DonateTotalLen = 13;
        for (int j = 0; j < DonateTotalLen; j++)
        {
            content += "0";
        }
        //12.保留欄
        SpaceLen = 15;
        content += Util.Space(SpaceLen);
        //content += "\r\n";
        //--------------------------------------------
        return content;
    }
}