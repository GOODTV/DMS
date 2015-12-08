using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;

public partial class DonorMgr_Gift_Merge : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "Gift_Merge";
        //權控處理
        AuthrityControl();
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            //資料合併按鈕
            btntransfer.Visible = false;
        }
    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_Query", btnQuery);
        Authrity.CheckButtonRight("_Update", btntransfer);
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        //轉出人
        string strSql_From;
        DataTable dt_From;
        strSql_From = @"Select * ,(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) as [通訊地址]
                        From Donor 
                            Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode
                        where Donor_Id = @Donor_Id and DeleteDate is null";

        Dictionary<string, object> dict_From = new Dictionary<string, object>();
        dict_From.Add("Donor_Id", tbxFrom_Donor_Id.Text.Trim());
        dt_From = NpoDB.GetDataTableS(strSql_From, dict_From);
        DataRow dr_From;

        if (dt_From.Rows.Count == 0)
        {
            SetSysMsg("您輸入的『轉出』捐款人編號不存在，請您重新確認！");
            Response.Redirect("Donate_Merge.aspx");
        }
        else
        {
            //轉出人基本資料
            dr_From = dt_From.Rows[0];
            lblDonor_From.Text = "轉出捐款人：" + dr_From["Donor_Name"].ToString() + "<br>";
            lblDonor_From.Text += "　通訊地址：" + dr_From["通訊地址"].ToString() + "<br>";
            lblDonor_From.Text += "欲轉出資料：<span class=\"necessary\">( 請先勾選然後按下『 資料合併 』按鈕) </span>";

            //轉出人的公關贈品資料
            strSql_From = @"select CD.Ser_No as [Ser_No], CONVERT(nvarchar,C.Gift_Date,111) as [贈送日期], CD.Goods_Name as [物品名稱], CD.Goods_Qty as [物品數量],
                                     CONVERT(VARCHAR(10) , CD.Create_Date, 111 ) as [建立日期], C.Create_User as [經手人] 
                            from GiftData CD
                                    left join Donor Dr on CD.Donor_Id = Dr.Donor_Id 
                                    left join Gift C on CD.Gift_Id = C.Gift_Id
                            where CD.Donor_Id = @Donor_Id and Dr.DeleteDate is null";
            dict_From = new Dictionary<string, object>();
            dict_From.Add("Donor_Id", tbxFrom_Donor_Id.Text.Trim());
            dt_From = NpoDB.GetDataTableS(strSql_From, dict_From);
            if (dt_From.Rows.Count == 0)
            {
                ShowSysMsg("您輸入的『轉出』捐款人編號無公關贈品紀錄，請您重新確認！");
                btntransfer.Visible = false;
            }
            else
            {
                btntransfer.Visible = true;
            }
            lblData_From.Text = CheckBoxData(strSql_From, "Ser_No", 0, dt_From);
        }

        //轉入人
        string strSql_To;
        DataTable dt_To;
        strSql_To = @"Select * ,(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) as [通訊地址]
                      From Donor 
                          Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode
                      where Donor_Id = @Donor_Id and DeleteDate is null";

        Dictionary<string, object> dict_To = new Dictionary<string, object>();
        dict_To.Add("Donor_Id", tbxTo_Donor_Id.Text.Trim());
        dt_To = NpoDB.GetDataTableS(strSql_To, dict_To);
        DataRow dr_To;

        if (dt_To.Rows.Count == 0)
        {

            SetSysMsg("您輸入的『轉入』捐款人編號不存在，請您重新確認！");
            Response.Redirect("Pledge_Merge.aspx");
        }
        else
        {
            dr_To = dt_To.Rows[0];
            //轉入人基本資料
            dr_To = dt_To.Rows[0];
            lblDonor_To.Text = "轉入捐款人：" + dr_To["Donor_Name"].ToString() + "<br>";
            lblDonor_To.Text += "　通訊地址：" + dr_To["通訊地址"].ToString() + "<br>";
            lblDonor_To.Text += "欲轉入資料：";

            //轉入人的捐款資料
            strSql_To = @"select CD.Ser_No as [Ser_No], CONVERT(nvarchar,C.Gift_Date,111) as [贈送日期], CD.Goods_Name as [物品名稱], CD.Goods_Qty as [物品數量],
                                     CONVERT(VARCHAR(10) , CD.Create_Date, 111 ) as [建立日期], C.Create_User as [經手人] 
                            from GiftData CD
                                    left join Donor Dr on CD.Donor_Id = Dr.Donor_Id 
                                    left join Gift C on CD.Gift_Id = C.Gift_Id
                            where CD.Donor_Id = @Donor_Id and Dr.DeleteDate is null";
            dict_To = new Dictionary<string, object>();
            dict_To.Add("Donor_Id", tbxTo_Donor_Id.Text.Trim());
            dt_To = NpoDB.GetDataTableS(strSql_To, dict_To);
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt_To;
            npoGridView.ShowPage = false;
            npoGridView.DisableColumn.Add("Ser_No");
            lblData_To.Text = npoGridView.Render();
        }

    }
    //---------------------------------------------------------------------------
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    protected void btntransfer_Click(object sender, EventArgs e)
    {
        string Ser_No = CatchSer_No();
        string SerNo = Ser_No.Substring(0, Ser_No.Length - 1);//去掉最後一個,

        if (Ser_No == "")
        {
            ShowSysMsg("您尚未勾選任何轉出資料");
            return;
        }
        else
        {
            //有幾筆資料
            string[] Ser_Nos = Ser_No.Split(',');
            string Gif_Id,GiftId = "";
            
            string strSql = @" update GiftData 
                                    set Donor_Id = @Donor_Id
                                    where Ser_No in (" + SerNo + ")";
            Dictionary<string, object> dict = new Dictionary<string, object>();
            dict.Add("Donor_Id", tbxTo_Donor_Id.Text);

            NpoDB.ExecuteSQLS(strSql, dict);
            for (int i = 0; i < Ser_Nos.Length-1; i++)
            {
                //抓取該物品的Gift_Id，再update
                Dictionary<string, object> dict2 = new Dictionary<string, object>();
                string strSql2 = @"select Gift_Id from GiftData
                                        where Ser_No = @Ser_No ";
                dict2.Add("Ser_No", Ser_Nos[i]);
                Gif_Id = NpoDB.GetScalarS(strSql2, dict2);
                GiftId += Gif_Id + ",";
            }
            string strGiftId = GiftId.Substring(0, GiftId.Length - 1);//去掉最後一個,
            //公關贈品需同時update兩個table
            Dictionary<string, object> dt = new Dictionary<string, object>();
            string sql = @" update Gift 
                                    set Donor_Id = @Donor_Id
                                    where Gift_Id in (" + strGiftId + ")";
            dt.Add("Donor_Id", tbxTo_Donor_Id.Text);
            NpoDB.ExecuteSQLS(sql, dt);


            int Count = Ser_Nos.Length - 1;
            ShowSysMsg("公關資料已合併成功，共計" + Count + " 筆");
            LoadFormData();
        }
    }
    public string CheckBoxData(string strSql, string FName, int DataNum, DataTable dt)
    {

        StringBuilder sb = new StringBuilder();

        //k用來決定第幾欄以後的內容是特殊內容(例如文字框(textbox))
        int k;

        k = dt.Columns.Count - DataNum;
        sb.AppendLine(@"<table width='100%' border='0' cellpadding='0' cellspacing='1' class='table_h'>");
        sb.AppendLine(@"<tr>");
        //20140430 Modify by GoodTV Tanya:增加「全選」的功能
        //sb.AppendLine(@"<th nowrap><span style='font-size: 9pt; font-family: 新細明體'>選擇</span></th>");
        sb.AppendLine(@"<th nowrap><span style='font-size: 9pt; font-family: 新細明體'><Input Type='checkbox' Name='checkboxAll' Id='checkboxAll' checked></span></th>");


        int c = 0;
        foreach (DataColumn dc in dt.Columns)
        {
            if (c == 0)
            {
                c++;
            }
            else
            {
                sb.AppendLine(@"<th nowrap><span style='font-size: 9pt; font-family: 新細明體'>" + dc.ColumnName + "</span></th>");
            }
        }
        sb.AppendLine("</tr>");
        int j = 1;
        int i = 1;
        foreach (DataRow dr in dt.Rows)
        {
            sb.AppendLine(@"<tr>");
            //20140430 Modify by GoodTV Tanya:增加「全選」的功能
            sb.AppendLine(@"<td nowrap><span style='font-size: 9pt; font-family: 新細明體'><Input Type='checkbox' Id='checkbox' Name='" + FName + dr[FName].ToString() + @"' value='" + j + "' checked></span></td>");
            i = 1;
            c = 0;
            foreach (DataColumn dc in dt.Columns)
            {
                if (c != 0)
                {
                    if (i > k)
                        sb.AppendLine(@"<td nowrap><span style='font-size: 9pt; ; font-family: 新細明體'><input type='text' name='" + i + "_" + j + @"' size='4' style='font-size: 9pt; font-family: 新細明體' value='" + dr[dc.ColumnName] + @"'></span></td>");
                    else
                        sb.AppendLine(@"<td nowrap><span style='font-size: 9pt; font-family: 新細明體'>" + dr[dc.ColumnName] + "</span></td>");
                }

                c++;
                i++;
            }
            j++;
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</table>");
        return sb.ToString();
    }
    //抓公關贈品序號
    private string CatchSer_No()
    {
        string Ser_No = "";
        string strSer_No = "";

        foreach (string str in Request.Form.Keys)
        {
            if (str.StartsWith("Ser_No"))
            {
                strSer_No = str.Split('o')[1];
                Ser_No += strSer_No + ",";
            }
        }
        return Ser_No;
    }
}