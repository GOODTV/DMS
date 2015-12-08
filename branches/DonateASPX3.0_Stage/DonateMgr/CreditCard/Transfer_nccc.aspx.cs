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

public partial class DonateMgr_CreditCard_Transfer_nccc : BasePage
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
        Transfer_FileName = dr["Transfer_FileName"].ToString() ;
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
        //交易日期
        string Donate_Year_AD = DateTime.Now.Year.ToString();
        string Donate_Year = (DateTime.Now.Year - 1911).ToString();
        if (Donate_Year.Length == 2)
        {
            Donate_Year = "0" + Donate_Year;
        }
        string Donate_Month = DateTime.Now.Month.ToString();
        if (Donate_Month.Length == 1)
        {
            Donate_Month = "0" + Donate_Month;

        }
        string Donate_Day = DateTime.Now.Day.ToString();
        if (Donate_Day.Length == 1)
        {
            Donate_Day = "0" + Donate_Day;
        }

        for (int i = 0; i < (Pledge_Ids.Length) - 1; i++)
        {
            string strSql = @" Select Pledge_Id,D.Donor_Id, Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, 
                                      Post_SavingsNo,Post_AccountNo,Post_IDNo,Donate_Amt= CAST((REPLACE((Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),'.00',''))as NVARCHAR),D.Donor_Name,Donate_ToDate 
                               From Pledge P Join Donor D On P.Donor_Id = D.Donor_Id 
                               Where Pledge_Id = '" + Pledge_Ids[i] + "' And Status='授權中' And Account_No<>''";

            Dictionary<string, object> dict = new Dictionary<string, object>();
            //dict.Add(" Pledge_Id", Pledge_Ids[i]);
            DataTable dt;
            DataRow dr;
            dt = NpoDB.GetDataTableS(strSql, dict);
            if (dt.Rows.Count != 0)
            {
                dr = dt.Rows[0];
                //--------------------------------------------
                //data
                //1.商店代號
                int StoreLen = 10;
                if (dr_Txt["Transfer_StoreCode"].ToString().Length >= StoreLen)
                {
                    content += Util.Left(dr_Txt["Transfer_StoreCode"].ToString(), StoreLen);
                }
                else
                {
                    content += dr_Txt["Transfer_StoreCode"].ToString() + Util.Space(StoreLen - dr_Txt["Transfer_StoreCode"].ToString().Length);
                }
                //2.端末機代號
                int ServereLen = 8;
                if (dr_Txt["Transfer_ServerCode"].ToString().Length >= ServereLen)
                {
                    content += Util.Left(dr_Txt["Transfer_ServerCode"].ToString(), ServereLen);
                }
                else
                {
                    content += dr_Txt["Transfer_ServerCode"].ToString() + Util.Space(ServereLen - dr_Txt["Transfer_ServerCode"].ToString().Length);
                }
                //3.卡號
                int CardLen = 16;
                string account_no = dr["account_no"].ToString();
                if (dr["account_no"].ToString().Length >= CardLen)
                {
                    content += Util.Left(dr["account_no"].ToString(), CardLen);
                }
                else
                {
                    content += dr["account_no"].ToString() + Util.Space(CardLen - dr["account_no"].ToString().Length);
                }
                //4.金額
                int AmtLen = 11;
                string Donate_Amt = dr["Donate_Amt"].ToString();
                for (int j = Donate_Amt.Length; j < AmtLen; j++)
                {
                    Donate_Amt = "0" + Donate_Amt;
                }
                content += Donate_Amt;
                //5.空白(交易碼8)
                int SpaceLen = 8;
                content += Util.Space(SpaceLen);
                //6.交易日期
                content += "00" + Util.Right(Donate_Year_AD,2) + Donate_Month + Donate_Day ;
                //7.信用卡有效年月
                int ValidDateLen = 4;
                string Valid_Date = dr["Valid_Date"].ToString();
                if (Valid_Date.Length >= ValidDateLen)
                {
                    content += Util.Left(Valid_Date, ValidDateLen);
                }
                else
                {
                    content += Valid_Date + Util.Space(ValidDateLen - (Valid_Date.Length));
                }
                //8.空白(回應碼3,回應訊息6,處理日期6,處理時間6)
                SpaceLen = 31;
                content += Util.Space(SpaceLen);
                //9.備註(商家自用)
                int RemarkLen = 22;
                string Donor_Remark = "";
                string Donor_Name = "";
                if (dr["Donor_Name"].ToString().IndexOf("&#") > -1)
                {
                    for (int j = 0; j < dr["Donor_Name"].ToString().Length; j++)
                    {
                        if (Util.Mid(dr["Donor_Name"].ToString(), j, 1) == "&" && j < dr["Donor_Name"].ToString().Length)
                        {
                            if (Util.Mid(dr["Donor_Name"].ToString(), j + 1, 1) == "#")
                            {
                                Donor_Name = Donor_Name + "*";
                                j = j + 7;
                            }
                            else
                            {
                                Donor_Name = Util.Mid(dr["Donor_Name"].ToString(), j, 1);
                            }
                        }
                        else
                        {
                            Donor_Name = Util.Mid(dr["Donor_Name"].ToString(), j, 1);
                        }
                    }
                }
                else
                {
                    Donor_Name = dr["Donor_Name"].ToString();
                }
                Donor_Name = "";
                string DonorRemark = Util.Left("00000", 5 - (dr["Pledge_Id"].ToString()).Length) + dr["Pledge_Id"].ToString() + Donor_Name + dr["Donor_Id"].ToString();
                int iLen = 0;
                int sLen = 0;
                for (int j = 0; j < DonorRemark.Length; j++)
                {
                    if (Util.Asc(Util.Mid(DonorRemark, j, 1)) < 0)
                    {
                        sLen = 2;
                    }
                    else
                    {
                        sLen = 1;
                    }
                    if (sLen + iLen <= RemarkLen)
                    {
                        Donor_Remark = Util.Mid(DonorRemark, j, 1);
                        iLen = iLen + sLen;
                    }
                    else
                    {
                        break;
                    }
                }
                content += DonorRemark;
                if (iLen < RemarkLen)
                {
                    content += Util.Space(RemarkLen - iLen);
                }
                //10.取消旗號
                content += "0";
                //11.交易旗號
                content += "2";

                Donate_Total += int.Parse(dr["Donate_Amt"].ToString());
                row += 1;
                content += "\r\n";
                //--------------------------------------------
            }
        }
        return content.Trim();
    }
}