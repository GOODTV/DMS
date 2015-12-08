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

public partial class DonateMgr_CreditCard_Transfer_cathay : BasePage
{

    string Donate_Year;
    string Donate_Month;
    string Donate_Day;
    string Donate_Hour;
    string Donate_Minute;

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

        //產出格式
        GridList.Text = GetTable(Uids, dr);

        //輸出檔案名稱
        string Transfer_FileName = dr["Transfer_FileName"].ToString();
        //匯出檔名加上交易扣款日期
        Transfer_FileName = Transfer_FileName.Replace(".txt", Donate_Year + Donate_Month + Donate_Day + Donate_Hour + Donate_Minute + ".txt");
        Util.OutputTxt(GridList.Text, "5", Transfer_FileName);
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
        //交易日期:交易日期為畫面上的扣款日期
        Donate_Year = Convert.ToDateTime(Session["Donate_Date"].ToString()).Year.ToString();     //DateTime.Now.Year.ToString();
        Donate_Month = Convert.ToDateTime(Session["Donate_Date"].ToString()).Month.ToString();   //DateTime.Now.Month.ToString();
        if (Donate_Month.Length == 1)
        {
            Donate_Month = "0" + Donate_Month;

        }
        Donate_Day = Convert.ToDateTime(Session["Donate_Date"].ToString()).Day.ToString();       //DateTime.Now.Day.ToString();
        if (Donate_Day.Length == 1)
        {
            Donate_Day = "0" + Donate_Day;
        }
        Donate_Hour = DateTime.Now.Hour.ToString();
        if (Donate_Hour.Length == 1)
        {
            Donate_Hour = "0" + Donate_Hour;
        }
        Donate_Minute = DateTime.Now.Minute.ToString();
        if (Donate_Minute.Length == 1)
        {
            Donate_Minute = "0" + Donate_Minute;
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
                //1.商店代號,15位,靠左，右補白
                int StoreLen = 15;
                string StoreCode = dr_Txt["Transfer_StoreCode"].ToString();
                content += StoreCode + Util.Space(StoreLen - StoreCode.Length);
                //2.訂單編號,20位,靠右，左補白
                int OrderLen = 20;
                string Pledge_Id = dr["Pledge_Id"].ToString().Trim();
                Pledge_Id = Util.Left("00000000", 8 - Pledge_Id.Length) + Pledge_Id;
                string Donate_Order = Donate_Year + Donate_Month + Donate_Day + Pledge_Id;
                content += Util.Space(OrderLen - (Donate_Order.Length)) + Donate_Order;
                //3.信用卡卡號,16位,靠右，左補白
                int CardLen = 16;
                string account_no = dr["account_no"].ToString();
                content += Util.Space(CardLen - account_no.Length) + account_no;
                //4.信用卡CSV2,3 位,Visa卡可不填寫此欄位，請補3位空白，Master卡則一定要填寫此欄位
                int AuthorizeLen = 3;
                string Authorize = dr["Authorize"].ToString();
                if (Authorize.Length > 0)
                {
                    content += Authorize + Util.Space(AuthorizeLen - Authorize.Length);
                }
                else
                {
                    content += Util.Space(AuthorizeLen);
                }
                //5.有效年月(西元年月),6 位,為YYYYMM
                int ValidDateLen = 6;
                //sql 撈取時,dr["Valid_Date"] 已轉換成YYMM格式
                string Valid_Date = "20" + dr["Valid_Date"].ToString();
                if (Valid_Date.Length > ValidDateLen)
                {
                    content += Util.Left(Valid_Date, ValidDateLen);
                }
                else
                {
                    content += Valid_Date + Util.Space(ValidDateLen - (Valid_Date.Length));
                }
                //6.金額,8 位,靠右，左補0
                int AmtLen = 8;
                string Donate_Amt = dr["Donate_Amt"].ToString();
                for (int j = Donate_Amt.Length; j < AmtLen ; j++)
                {
                    Donate_Amt = "0" + Donate_Amt;
                }
                content += Donate_Amt;

                Donate_Total += int.Parse(dr["Donate_Amt"].ToString());
                row += 1;
                content += "\r\n";
                //長度共68位
                //--------------------------------------------
            }
        }
        return content;
    }

}
