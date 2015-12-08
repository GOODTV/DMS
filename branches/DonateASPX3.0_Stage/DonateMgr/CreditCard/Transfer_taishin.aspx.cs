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

public partial class DonateMgr_CreditCard_Transfer_taishin : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string Uids = Session["Pledge_Id"].ToString();
        Print_Txt(Uids);
    }
    //---------------------------------------------------------------------------
    private void Print_Txt(string Uids)
    {
        if (Uids=="")
        {
            ShowSysMsg("查無資料!!!");
            return;
        }

        //抓出要轉出的txt格式資料
        string strSql = @"Select * From DONATE_TRANSFER Where Transfer_AspxName='" + Session["Transfer_AspxName"].ToString() + "'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        //輸出檔案名稱
        string OrderNo,Transfer_FileName ="";
        if (dr["Transfer_StoreCode"].ToString() != "")
        {
           //if(dr["Transfer_StoreCode"].ToString().IndexOf(DateTime.Now.Year.ToString()+DateTime.Now.Month.ToString()+DateTime.Now.Day.ToString())>-1)
           if(dr["Transfer_StoreCode"].ToString().IndexOf(DateTime.Now.ToString("yyyyMMdd"))>-1)
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
        Transfer_FileName = dr["Transfer_StoreCode"].ToString() + OrderNo;

        GridList.Text = GetTable(Uids, dr);
        Util.OutputTxt(GridList.Text, "4", Transfer_FileName);
    }
    //---------------------------------------------------------------------------
    private string GetTable(string Uids, DataRow dr_Txt)
    {
        //抓出txt中的資料
        string[] Pledge_Ids = Uids.Split(' ');
        string content="";
        //授權筆數
        int row = 0;
        //授權金額
        int Donate_Total = 0;
        for (int i = 0; i < (Pledge_Ids.Length)-1; i++)
        {
            string strSql = @" Select Pledge_Id,P.Donor_Id,Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, 
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
                //Data Record
                //1.Record type
                content += "D";
                //2.Transaction type (Space(2))
                int SpaceLen = 2;
                content += Util.Space(SpaceLen);
                //3.Merchant Number + 4.MCC Code
                content += dr_Txt["Transfer_StoreCode"].ToString() + dr_Txt["Transfer_ServerCode"].ToString() ;
                //5.卡號
                int CardLen = 19;
                content += dr["Account_No"].ToString() + Util.Space(CardLen - dr["Account_No"].ToString().Length);
                //6.信用卡有效年月
                int ValidDateLen = 4;
                string Valid_Date = Util.Left(dr["Account_No"].ToString(), 2) + Util.Right(dr["Account_No"].ToString(), 2);
                content += Util.Left(Valid_Date + Util.Space(ValidDateLen - Valid_Date.Length), ValidDateLen);
                //7.金額
                content += Util.Left("0000000000", 10 - (dr["Donate_Amt"].ToString().Length)) + dr["Donate_Amt"].ToString() + "00";
                //8.Response Code (Space(2))
                SpaceLen = 2;
                content += Util.Space(SpaceLen);
                //9.Approval Code (Space(6))
                SpaceLen = 6;
                content += Util.Space(SpaceLen);
                //10.交易日期 
                content += DateTime.Now.ToString("yyyyMMdd");
                //11.SSL / SET flag (Space(1)) 
                SpaceLen = 1;
                content += Util.Space(SpaceLen);
                //12.CVC2 (Space(3)) 
                SpaceLen = 3;
                content += Util.Space(SpaceLen);
                //13.SET XID (Space(40))
                SpaceLen = 40;
                content += Util.Space(SpaceLen);
                //14.備註(商家自用)
                int RemarkLen = 60;
                string Donor_Name= "";
                string Donor_Remark = "";
                string DonorRemark = "";
                if (dr["Donor_Name"].ToString().IndexOf("&#") > -1)
                {
                    for(int j=0;j<dr["Donor_Name"].ToString().Length;j++)
                    {
                        string strName = Util.Mid(dr["Donor_Name"].ToString(), j, 1);
                        if (strName == "&" && j < dr["Donor_Name"].ToString().Length)
                        {
                            string strName2 = Util.Mid(dr["Donor_Name"].ToString(), j + 1, 1);
                            if (strName2 == "#")
                            {
                                Donor_Name = dr["Donor_Name"].ToString() + "*";
                                j = j + 7;
                            }
                            else
                            {
                                Donor_Name = dr["Donor_Name"].ToString() + strName;
                            }
                        }
                        else
                        {
                            Donor_Name = dr["Donor_Name"].ToString() + strName;
                        }
                    }
                }
                else
                {
                    Donor_Name = dr["Donor_Name"].ToString();
                }
                DonorRemark = Util.Left("00000", 5 - (dr["Pledge_Id"].ToString()).Length) + dr["Pledge_Id"].ToString() + Donor_Name + dr["Donor_Id"].ToString();
                int iLen = 0;
                int sLen = 0;
                for(int j=1;j< DonorRemark.Length;j++)
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
                        iLen = iLen +sLen;
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
                //15.交易類別Message Type + 16.分期期數Installment Period
                content += "00" + "00";
                //7.訂單編號Installment OrderNo
                SpaceLen = 12;
                content += Util.Space(SpaceLen);
                //18.首期應繳金額First Period Amount + 19.每期應繳金額Period Amount + 20.手續費Process Fee
                content += "000000000000" + "000000000000" + "000000000000";
                //21.首期手續費First Period Process Fee + 22.每期手續費Period Process Fee
                content += "000000000000" + "000000000000" + "1";
                //23.訊息來源Message Source 
                SpaceLen = 1;
                content += Util.Space(SpaceLen);
                //24.強制授權(郵購用)Strong Authorization 
                SpaceLen = 1;
                content += Util.Space(SpaceLen);
                //25.郵購商品代碼MO PRODUCTNO + 26.郵購商品數量MO_Quantity
                content += "000000000000000" + "000";
                //27.System Trace Audit No. 
                SpaceLen = 6;
                content += Util.Space(SpaceLen);
                //28.Transmission Date 
                SpaceLen = 8;
                content += Util.Space(SpaceLen);
                //29.Transmission Time
                SpaceLen = 6;
                content += Util.Space(SpaceLen);
                //30.Retrieval Reference Number 
                SpaceLen = 12;
                content += Util.Space(SpaceLen);
                //31.Filler
                SpaceLen = 24;
                content += Util.Space(SpaceLen);
                content += "\r\n";

                row += 1;
                Donate_Total += int.Parse(dr["Donate_Amt"].ToString());
                //Trailer Record
                //1.Record type + 2.Merchant Number + 3.Total 05 items扣款
                content += "T" + dr_Txt["Transfer_StoreCode"].ToString() + Util.Left("00000", 6 - (row.ToString()).Length) + row.ToString();
                //4.Total 05 amount扣款 + 5.Total 06 items刷退 + 6.Total 06 amount刷退
                content += Util.Left("0000000000", 10 - (Donate_Total.ToString()).Length) + Donate_Total.ToString() + "00" + "000000" + "000000000000";
                content += "\r\n";
                //--------------------------------------------
            }
        }
        return content;
    }
}