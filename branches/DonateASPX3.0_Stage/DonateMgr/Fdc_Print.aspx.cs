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

public partial class DonateMgr_Fdc_Print : BasePage
{
    List<ControlData> list = new List<ControlData>();
    string TableWidth = "800px";

    protected void Page_Load(object sender, EventArgs e)
    {
        string Uids = Session["Donor_Id"].ToString();
        Print_Excel(Uids);
    }
    //---------------------------------------------------------------------------
    private void Print_Excel(string Uids)
    {
        if (Uids=="")
        {
            ShowSysMsg("查無資料!!!");
            return;
        }
        GridList.Text = GetTable(Uids);
        //GetTableFooter();
        //20140422Modify by GoodTV Tanya:匯出檔為*.csv
        Util.OutputTxt(GridList.Text, "5", (int.Parse(Session["Donate_Date_Year"].ToString()) - 1911).ToString() + ".31." + Session["Uniform_No"] + ".csv");
    }
    //---------------------------------------------------------------------------
    private string GetTable(string Uids)
    {
        string[] Donate_Ids = Uids.Split(' ');
        string content="";

        //20140422Modify by GoodTV Tanya:匯出資料加上標題，並以逗號區隔
        //output(標題列)
        content += "捐贈年度,捐贈者身分證統一編號,捐贈者姓名,捐款金額,受捐贈單位統一編號,捐贈別,受捐贈者姓名,專案核准文號" + "\r\n";

        for (int i = 0; i < Donate_Ids.Length; i++)
        {
            string strSql = @" select dr.Invoice_IDNo,dr.Donor_Name,dr.Invoice_Title,
                                      CAST((REPLACE(sum(do.Donate_Amt),'.00',''))as NVARCHAR) as [Donate_Total]
                               from Donor dr left join Donate do on dr.Donor_Id = do.Donor_Id
                               Where dr.DeleteDate is null and Year(do.Donate_Date)=@Year and do.Donate_Amt>0 whereclause and dr.Donor_Id = @Donor_Id
                               group by Invoice_IDNo,Donor_Name,dr.Invoice_Title";
            

            Dictionary<string, object> dict = new Dictionary<string, object>();
            dict.Add("Year", Session["Donate_Date_Year"].ToString());
            dict.Add("Donor_Id", Donate_Ids[i]);
            
            if (Session["Act_Id"].ToString() == "")
            {
                strSql = strSql.Replace("whereclause", "");
            }
            else
            {
                strSql = strSql.Replace("whereclause", "and Act_Id=@Act_Id");
                dict.Add("Act_Id", Session["Act_Id"].ToString());
            }
            DataTable dt;
            DataRow dr;
            dt = NpoDB.GetDataTableS(strSql, dict);
            if (dt.Rows.Count != 0)
            {
                
                dr = dt.Rows[0];
                //data
                //--------------------------------------------
                //捐贈年度
                string Donate_Year = (int.Parse(Session["Donate_Date_Year"].ToString()) - 1911).ToString().PadLeft(3, '0');
                //捐贈者身分證統一編號(收據身分證/統編)
                string IDNo = dr["Invoice_IDNo"].ToString().ToUpper();
                //string IDNo = "";
                //int IDNoLen = 10;
                //if (IDNo_temp.Length >= IDNoLen)
                //{
                //    IDNo = Util.Left(IDNo_temp, IDNoLen);

                //}
                //else
                //{
                //    IDNo = IDNo_temp + Util.Space(IDNoLen - IDNo_temp.Length);
                //}
                //捐贈者姓名(收據抬頭)
                string Invoice_Title = dr["Invoice_Title"].ToString();
                //int NameLen = 12;
                //string InvoiceTitle = "";
                //string strInvoice_Title = dr["Invoice_Title"].ToString();
                //if (strInvoice_Title.IndexOf("&#", 0) > 0)
                //{
                //    for (int a = 0; a < strInvoice_Title.Length; a++)
                //    {
                //        if (Util.Mid(strInvoice_Title, a, 1) == "&" && a < strInvoice_Title.Length)
                //        {
                //            if (Util.Mid(strInvoice_Title, a + 1, 1) == "#")
                //            {
                //                InvoiceTitle = InvoiceTitle + "*";
                //                a = a + 7;
                //            }
                //            else
                //            {
                //                InvoiceTitle = InvoiceTitle + Util.Mid(strInvoice_Title, a, 1);
                //            }
                //        }
                //        else
                //        {
                //            InvoiceTitle = InvoiceTitle + Util.Mid(strInvoice_Title, a, 1);
                //        }
                //    }
                //}
                //else
                //{
                //    InvoiceTitle = strInvoice_Title;
                //}
                //InvoiceTitle = InvoiceTitle.Replace(" ", "~");//空格會使function出錯，只能先暫時用其他符號代替之後再替換回空白
                //string Invoice_Title = "";
                //int TitleLen = 0;
                //for (int b = 0; b < InvoiceTitle.Length; b++)
                //{
                //    if (Util.Asc(Util.Mid(InvoiceTitle, b, 1)) < 0)
                //    {
                //        TitleLen = TitleLen + 2;
                //    }
                //    else
                //    {
                //        TitleLen = TitleLen + 1;
                //    }
                //    if (TitleLen <= 12)
                //    {
                //        Invoice_Title = Invoice_Title + Util.Mid(InvoiceTitle, b, 1);
                //    }
                //    else
                //    {
                //        if (Util.Asc(Util.Mid(InvoiceTitle, b, 1).Trim()) < 0)
                //        {
                //            TitleLen = TitleLen - 2;
                //        }
                //        else
                //        {
                //            TitleLen = TitleLen - 1;
                //        }
                //        break;
                //    }
                //}
                //Invoice_Title = Invoice_Title.Replace("~", " ");//替換回空白
                //捐款金額
                //int TotalLen = 10;
                string Donate_Total = dr["Donate_Total"].ToString();
                //if( Donate_Total.Length >= TotalLen )
                //{
                //    Donate_Total = Util.Right(Donate_Total,TotalLen);
                //}
                //else
                //{
                //    Donate_Total = Util.Left("0000000000",10-Donate_Total.Length) + Donate_Total;
                //}
                //受捐贈單位統一編號
                string Uniform_No = Session["Uniform_No"].ToString();
                //Uniform_No = Uniform_No + Util.Space(10-Uniform_No.Length);                
                //受捐贈者姓名
                string Comp_Name = Session["OrgName"].ToString();
                //int CompNameLen = 0;
                //for (int c = 0; c < Comp_Name.Length; c++)
                //{
                //    if (Util.Asc(Util.Mid(Comp_Name, c, 1)) < 0)
                //    {
                //        CompNameLen = CompNameLen + 2;
                //    }
                //    else
                //    {
                //        CompNameLen = CompNameLen + 1;
                //    }
                //}
                //Comp_Name = Comp_Name + Util.Space(40 - CompNameLen);
                //許可文號
                string Licence = Session["Licence"].ToString();
                //int LicenceLen = 0;
                //for (int d = 0; d < Licence.Length; d++)
                //{
                //    if (Util.Asc(Util.Mid(Licence, d, 1)) < 0)
                //    {
                //        LicenceLen = LicenceLen + 2;
                //    }
                //    else
                //    {
                //        LicenceLen = LicenceLen + 1;
                //    }
                //}
                //Licence = Licence + Util.Space(60 - LicenceLen);



                //output(明細)              
                content += Donate_Year + ",";
                content += IDNo + ",";
                content += Invoice_Title + ",";
                //if (TitleLen < 12)
                //{
                //    content += Util.Space(12 - TitleLen);
                //}
                content += Donate_Total + ",";
                content += Uniform_No + ",";
                content += "00,";//捐贈別+ 
                content += Comp_Name + ",";
                content += Licence + "\r\n";
                //--------------------------------------------
            }
        }
        return content;
    }
}