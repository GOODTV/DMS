using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;
using System.Web.UI.HtmlControls;

public partial class ContributeMgr_Contribute_Detail : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            HFD_Uid.Value = Util.GetQueryString("Contribute_Id");
            Form_DataBind();
        }
        Export();
    }
    //----------------------------------------------------------------------
    //帶入資料
    public void Form_DataBind()
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;

        //****變數設定****//
        uid = HFD_Uid.Value;

        //****設定查詢****//
        strSql = @" select *, (Case When Dr.Invoice_City='' Then Dr.Invoice_Address Else Case When B.mValue<>E.mValue Then B.mValue+Dr.Invoice_ZipCode+E.mValue+Dr.Invoice_Address Else B.mValue+Dr.Invoice_ZipCode+Dr.Invoice_Address End End) as [通訊地址]  
                    from Contribute C left join Donor Dr on C.Donor_Id = Dr.Donor_ID  
                        Left Join Act A on C.Act_Id = A.Act_Id 
                        Left Join Dept D on C.Dept_Id = D.DeptID 
                        Left Join CODECITY As B On Dr.Invoice_City=B.mCode 
                        Left Join CODECITY As E On Dr.Invoice_Area=E.mCode 
                    where Dr.DeleteDate is null and C.Contribute_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("ContributeList.aspx");

        DataRow dr = dt.Rows[0];

        HFD_Donor_Id.Value = dr["Donor_Id"].ToString().Trim();
        //捐款人
        tbxDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
        //類別
        tbxCategory.Text = dr["Category"].ToString().Trim();
        //身份別
        tbxDonor_Type.Text = dr["Donor_Type"].ToString().Trim();
        //收據地址
        tbxAddress.Text = dr["通訊地址"].ToString().Trim();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //捐贈日期
        tbxContribute_Date.Text = DateTime.Parse(dr["Contribute_Date"].ToString()).ToString("yyyy/MM/dd");
        //捐款方式
        tbxContribute_Payment.Text = dr["Contribute_Payment"].ToString().Trim();
        //捐款用途
        tbxContribute_Purpose.Text = dr["Contribute_Purpose"].ToString().Trim();
        //收據開立
        tbxInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();
        //機構
        tbxDept.Text = dr["DeptShortName"].ToString().Trim();
        //收據抬頭
        tbxInvoice_Title.Text = dr["Invoice_Title"].ToString().Trim();
        //收據號碼
        tbxInvoice_No.Text = dr["Invoice_Pre"].ToString().Trim() + dr["Invoice_No"].ToString().Trim();

        //沖帳日期
        if (dr["Accoun_Date"].ToString() != "" && DateTime.Parse(dr["Accoun_Date"].ToString()).ToString("yyyy/MM/dd") != "1900/01/01")
        {
            tbxAccoun_Date.Text = DateTime.Parse(dr["Accoun_Date"].ToString()).ToString("yyyy/MM/dd");
        }
        //會計科目
        tbxAccounting_Title.Text = dr["Accounting_Title"].ToString().Trim();
        //專案活動
        tbxAct_Id.Text = dr["Act_ShortName"].ToString().Trim();
        //捐款備註
        tbxComment.Text = dr["Comment"].ToString().Trim();
        //收據備註
        tbxInvoice_PrintComment.Text = dr["Invoice_PrintComment"].ToString().Trim();
        //總捐贈金額
        HFD_Contribute_Amt.Value = (Convert.ToInt64(dr["Contribute_Amt"])).ToString();
        //載入捐贈內容
        string strSql2, uid2;
        DataTable dt2;
        uid2 = HFD_Uid.Value;
        strSql2 = " select ROW_NUMBER() OVER(ORDER BY Ser_No) as [ROWID], ROW_NUMBER() OVER(ORDER BY Ser_No) as [序號], \n";
        strSql2 += " G.Goods_Name as [物品名稱], CD.Goods_Qty as [數量] , G.Goods_Unit as [單位], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Goods_Amt),1),'.00','')  as [折合現金], \n";
        strSql2 +=" (case when Goods_DueDate!='' or CONVERT(VARCHAR(10) , Goods_DueDate, 111 )!='1900/01/01' then CONVERT(VARCHAR(10), \n";
        strSql2 +=" Goods_DueDate, 111 ) else '' end)  as [保存期限], Goods_Comment as [備註], \n";
        strSql2 +=" (case when Contribute_IsStock='Y' then 'V' else '' end) as 寫入庫存\n";
        strSql2 += " from ContributeData CD right join Goods G on CD.Goods_Id = G.Goods_Id\n";
        strSql2 +=" where Contribute_Id='" + uid2 + "'\n";
        strSql2 +=" union \n";
        strSql2 += " select '100',null,'',null,'折合現金合計：',REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Goods_Amt!='0' then Goods_Amt else '0' end)),1),'.00','') ,'','',''\n";
        strSql2 += " from ContributeData where Contribute_Id='" + uid2 + "'";
        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        dt2 = NpoDB.GetDataTableS(strSql2, dict2);
        //Grid initial
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt2;
        npoGridView.DisableColumn.Add("ROWID");
        npoGridView.ShowPage = false;
        lblGridList.Text = npoGridView.Render();
    }
    public void Export()
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;

        //****變數設定****//
        uid = HFD_Uid.Value;

        //****設定查詢****//
        strSql = @" select Export , Invoice_Print
                    from Contribute C
                    where C.Contribute_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);

        DataRow dr = dt.Rows[0];
        //判斷是否作廢
        if (dr["Export"].ToString() == "N")
        {
            if (dr["Invoice_Print"].ToString() == "0")
            {
                Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Export_N_Invoice_Print_N();", true);
            }
            if (dr["Invoice_Print"].ToString() == "1")
            {
                Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Export_N_Invoice_Print_Y();", true);
            }
        }
        if (dr["Export"].ToString() == "Y")
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Export_Y();", true);
        }
    }
    //----------------------------------------------------------------------
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("Contribute_Edit.aspx", "Contribute_Id=" + HFD_Uid.Value));
    }
    protected void btnExport_Click(object sender, EventArgs e)
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;
        //****變數設定****//
        uid = HFD_Uid.Value;
        //****設定查詢****//
        strSql = @" select Invoice_Print 
                    from Contribute C
                    where C.Contribute_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Invoice_Print = dr["Invoice_Print"].ToString().Trim();
        //作廢單據前要先列印
        if (Invoice_Print == "0")
        {
            ShowSysMsg("作廢收據前請先列印");
            return;
        }
        else
        {
            //******修改Contribute的資料******//
            //****變數宣告****//
            Dictionary<string, object> dict = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql2 = " update Contribute set ";
            strSql2 += " Export= @Export";
            strSql2 += " where Contribute_Id = @Contribute_Id";
            dict.Add("Export", "Y");
            dict.Add("Contribute_Id", HFD_Uid.Value);
            NpoDB.ExecuteSQLS(strSql2, dict);

            //找出ContributeData有幾筆資料
            //****設定查詢****//
            string strSql3 = " select *  from ContributeData where Contribute_Id='" + HFD_Uid.Value + "'";
            //****執行語法****//
            DataTable dt2;
            dt2 = NpoDB.QueryGetTable(strSql3);
            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                DataRow dr2 = dt2.Rows[i];
                string Goods_Id = dr2["Goods_Id"].ToString();
                string Goods_Qty = dr2["Goods_Qty"].ToString();

                //******修改Goods的資料******//
                Dictionary<string, object> dict_Goods = new Dictionary<string, object>();
                //****設定SQL指令****//
                string strSql_Goods = " update Goods set ";
                strSql_Goods += "  Goods_Qty = Goods_Qty -'" + Goods_Qty + "'";
                strSql_Goods += " where Goods_Id = @Goods_Id";
                dict_Goods.Add("Goods_Id", Goods_Id);
                NpoDB.ExecuteSQLS(strSql_Goods, dict_Goods);
            }

            //******修改Donor的資料******//
            Dictionary<string, object> dict_Donor = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Donor = " update Donor set ";
            strSql_Donor += "  Donate_NoC = Donate_NoC -1";

            //找尋同個人是否有其他筆捐款紀錄
            string strSql4 = @"select isnull(MAX(CONVERT(NVARCHAR, Create_Date, 111)),'') as Create_Date from Contribute ";
            strSql4 += " where Donor_Id='" + HFD_Donor_Id.Value + "' and Export != 'Y'";
            DataTable dt4 = NpoDB.GetDataTableS(strSql4, null);
            DataRow dr4 = dt4.Rows[0];
            //找不到其他筆
            if (dr4["Create_Date"].ToString() == "")
            {
                strSql_Donor += "  ,Begin_DonateDateC = ''";
                strSql_Donor += "  ,Last_DonateDateC = ''";
            }
            else//有其他筆
            {
                strSql_Donor += "  ,Last_DonateDateC = '" + dr4["Create_Date"].ToString() + "'";
            }
            strSql_Donor += "  ,Donate_TotalC = Donate_TotalC -'" + int.Parse(HFD_Contribute_Amt.Value) + "'";
            strSql_Donor += " where Donor_Id = @Donor_Id";
            dict_Donor.Add("Donor_Id", HFD_Donor_Id.Value);
            NpoDB.ExecuteSQLS(strSql_Donor, dict_Donor);

            SetSysMsg("收據已作廢！");
            Response.Redirect(Util.RedirectByTime("Contribute_Detail.aspx", "Contribute_Id=" + HFD_Uid.Value));
        }
    }
    protected void btn_ReExport_Click(object sender, EventArgs e)
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Contribute set ";
        strSql += " Export= @Export";
        strSql += " where Contribute_Id = @Contribute_Id";
        dict.Add("Export", "N");
        dict.Add("Contribute_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);

        //找出ContributeData有幾筆資料
        //****設定查詢****//
        string strSql3 = " select *  from ContributeData where Contribute_Id='" + HFD_Uid.Value + "'";
        //****執行語法****//
        DataTable dt2;
        dt2 = NpoDB.QueryGetTable(strSql3);
        for (int i = 0; i < dt2.Rows.Count; i++)
        {
            DataRow dr2 = dt2.Rows[i];
            string Goods_Id = dr2["Goods_Id"].ToString();
            string Goods_Qty = dr2["Goods_Qty"].ToString();

            //******修改Goods的資料******//
            Dictionary<string, object> dict_Goods = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Goods = " update Goods set ";
            strSql_Goods += "  Goods_Qty = Goods_Qty +'" + Goods_Qty + "'";
            strSql_Goods += " where Goods_Id = @Goods_Id";
            dict_Goods.Add("Goods_Id", Goods_Id);
            NpoDB.ExecuteSQLS(strSql_Goods, dict_Goods);
        }

        //******修改Donor的資料******//
        Dictionary<string, object> dict_Donor = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql_Donor = " update Donor set ";
        strSql_Donor += "  Donate_NoC = Donate_NoC + 1";

        //找尋同個人是否有其他筆捐款紀錄
        string strSql4 = @"select isnull(MAX(CONVERT(NVARCHAR, Create_Date, 111)),'') as Create_Date from Contribute ";
        strSql4 += " where Donor_Id='" + HFD_Donor_Id.Value + "' and Export != 'N'";
        DataTable dt4 = NpoDB.GetDataTableS(strSql4, null);
        DataRow dr4 = dt4.Rows[0];
        //找不到其他筆
        if (dr4["Create_Date"].ToString() == "")
        {
            strSql_Donor += "  ,Begin_DonateDateC = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";
            strSql_Donor += "  ,Last_DonateDateC = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";
        }
        else//有其他筆
        {
            strSql_Donor += "  ,Last_DonateDateC = '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";
        }
        strSql_Donor += "  ,Donate_TotalC = Donate_TotalC +'" + int.Parse(HFD_Contribute_Amt.Value) + "'";
        strSql_Donor += " where Donor_Id = @Donor_Id";
        dict_Donor.Add("Donor_Id", HFD_Donor_Id.Value);
        NpoDB.ExecuteSQLS(strSql_Donor, dict_Donor);

        SetSysMsg("收據已還原！");
        Response.Redirect(Util.RedirectByTime("Contribute_Detail.aspx", "Contribute_Id=" + HFD_Uid.Value));
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        if (Session["cType"] == "ContributeList")
        {
            Response.Redirect(Util.RedirectByTime("ContributeList.aspx"));
        }
        else if (Session["cType"] == "ContributeDataList")
        {
            Response.Redirect(Util.RedirectByTime("ContributeDataList.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
        }
        else
        {
            Response.Redirect(Util.RedirectByTime("ContributeList.aspx"));
        }
    }
    protected void btn_Add_self_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("Contribute_Add.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
    }
    protected void btn_Add_other_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("DonorQry.aspx"));
    }
}