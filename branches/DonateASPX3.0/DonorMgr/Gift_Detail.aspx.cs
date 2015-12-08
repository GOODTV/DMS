using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;

public partial class DonorMgr_Gift_Detail : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            HFD_Uid.Value = Util.GetQueryString("Gift_Id");
            Form_DataBind();
        }
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
        strSql = @" select *   
                    from Gift C left join Donor Dr on C.Donor_Id = Dr.Donor_ID  
                        Left Join Act A on C.Act_Id = A.Act_Id 
                        Left Join Dept D on C.Dept_Id = D.DeptID
                    where Dr.DeleteDate is null and C.Gift_Id='" + uid + "'";

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
        //tbxCategory.Text = dr["Category"].ToString().Trim();
        //身份別
        tbxDonor_Type.Text = dr["Donor_Type"].ToString().Trim();
        //收據地址
        //tbxAddress.Text = dr["通訊地址"].ToString().Trim();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //贈送日期
        tbxContribute_Date.Text = DateTime.Parse(dr["Gift_Date"].ToString()).ToString("yyyy/MM/dd");
        //捐款備註
        tbxComment.Text = dr["Comment"].ToString().Trim();
        //總捐贈金額
        //HFD_Contribute_Amt.Value = (Convert.ToInt64(dr["Contribute_Amt"])).ToString();
        //載入捐贈內容
        string strSql2, uid2;
        DataTable dt2;
        uid2 = HFD_Uid.Value;
        strSql2 = " select ROW_NUMBER() OVER(ORDER BY CD.Ser_No) as [ROWID], ROW_NUMBER() OVER(ORDER BY CD.Ser_No) as [序號], \n";
        strSql2 += " L.Linked2_Name as [物品名稱], CD.Goods_Qty as [數量] from GiftData CD right join Linked2 L on CD.Goods_Id = L.Ser_No\n";
        strSql2 += " where Gift_Id='" + uid2 + "'\n";
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
    //----------------------------------------------------------------------
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("Gift_Edit.aspx", "Gift_Id=" + HFD_Uid.Value));
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        if (Session["cType"].ToString() == "ContributeDataList")
        {
            Response.Redirect(Util.RedirectByTime("../ContributeMgr/ContributeDataList.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
        }
        else
        {
            Response.Redirect(Util.RedirectByTime("../ContributeMgr/ContributeList.aspx"));
        }
    }
}