using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class DonorMgr_Gift_Add : BasePage
{
    string strOrderIds = "";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Donor_Id");
            Form_DataBind();
            tbxGift_Date.Text = DateTime.Now.ToString("yyyy/MM/dd");
            //20140827新增
            LoadFormData();
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
        strSql = @" select * from Donor 
                    where Donor_id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("ContributeDataList.aspx");

        DataRow dr = dt.Rows[0];

        //捐款人
        tbxDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
    }
    //----------------------------------------------------------------------
    protected void btnAdd_Click(object sender, EventArgs e)
    {

        strOrderIds = "";
        string strId = "";

        int count = 0; // 合計筆數
        foreach (string str in Request.Form.Keys)
        {
            if (str.StartsWith("orderid") && str != "orderidAll")
            {
                count++;
                strId = str.Split('_')[1];//將勾選的Linked2.Ser_No串成字串
                if (count == 1) strOrderIds = strId;
                else strOrderIds += "," + strId;
            }
        }
        string[] strOrderId = strOrderIds.Split(',');//把Linked2.Ser_No分開獨立出來
        int count2 = strOrderId.Length;//計算Linked2.Ser_No共有幾個 

        string strSql = "";
        DataTable dt;
        DataRow dr;
        //******新增公關贈品資料的資料******//

        bool flag = false;
        try
            {
                Gift_AddNew();//Gift table

                for (int i = 0; i < count2; i++)
                {
                    strSql = "select * from Linked2 where Ser_No = '" + strOrderId[i] + "'";
                    dt = NpoDB.QueryGetTable(strSql);
                    dr = dt.Rows[0];
                    String GiftsQty = "GiftsQty_" + strOrderId[i];
                    String GiftsComment = "GiftsComment_" + strOrderId[i];
                    GiftData_AddNew(strOrderId[i], dr["Linked2_Name"].ToString(), Request.Form[GiftsQty].Trim(), Request.Form[GiftsComment].Trim());
                }
                flag = true;
            }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("資料新增成功！");

            Response.Redirect(Util.RedirectByTime("../ContributeMgr/ContributeDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
        }
        
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("../ContributeMgr/ContributeDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
    }
    public void Gift_AddNew()
    {
        string strSql = "insert into  Gift\n";
        strSql += "( Donor_Id, Gift_Date, Gift_Payment, Dept_Id \n";
        strSql += " , Act_id, Comment \n";
        strSql += " , Create_Date, Create_User, Create_IP) values\n";
        strSql += "( @Donor_Id,@Gift_Date,@Gift_Payment,@Dept_Id \n";
        strSql += " ,@Act_id,@Comment \n";
        strSql += " ,@Create_Date,@Create_User,@Create_IP)";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", HFD_Uid.Value);
        dict.Add("Gift_Date", tbxGift_Date.Text.Trim());
        dict.Add("Gift_Payment", "公關贈品");
        dict.Add("Dept_Id", SessionInfo.DeptID);
        dict.Add("Act_id", "");
        dict.Add("Comment", tbxComment.Text);
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        NpoDB.ExecuteSQLS(strSql, dict);

        strSql = "Select Max(Gift_Id) Gift_Id From GIFT Where Donor_Id = '" + HFD_Uid.Value + "' And Gift_Date = '" + tbxGift_Date.Text.Trim() + "'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        HFD_Gift_Id.Value = dt.Rows[0]["Gift_Id"].ToString();

    }
    public void GiftData_AddNew(string Goods_Id, string Goods_Name, string Goods_Qty, string Goods_Comment)
    {
        string strSql = "insert into GiftData\n";
        strSql += "( Gift_Id, Donor_Id, Goods_Id, Goods_Name, Goods_Qty, \n";
        strSql += " Goods_Comment, Create_Date, Create_User, \n";
        strSql += " Create_IP) values\n";

        strSql += "( @Gift_Id,@Donor_Id,@Goods_Id,@Goods_Name,@Goods_Qty, \n";
        strSql += " @Goods_Comment,@Create_Date, @Create_User, \n";
        strSql += " @Create_IP)";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Gift_Id", HFD_Gift_Id.Value);
        dict.Add("Donor_Id", HFD_Uid.Value);
        dict.Add("Goods_Id", Goods_Id);
        dict.Add("Goods_Name", Goods_Name);
        if (string.IsNullOrEmpty(Goods_Qty) || Goods_Qty == "")
            dict.Add("Goods_Qty", "1");
        else
            dict.Add("Goods_Qty", Goods_Qty);
        dict.Add("Goods_Comment", Goods_Comment);
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    
    //20140827 改為大批匯入 by 詩儀
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = " Select Linked2.Ser_No as [物品編號] ,Linked.Linked_Name as [公關贈品品項] ,Linked2_Name as [物品名稱] ,Linked.Linked_Id as [物品代號] ";
        strSql += " From LINKED2 Join Linked On Linked.Linked_Id=LINKED2.Linked_Id Where Linked.Linked_Type='gift' ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        strSql += " Order By Linked2.Linked_Id,Linked2_Seq ";

        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
            int Donate_cnt = dt.Rows.Count;
            string strBody = "<table class='table_h' width='100%'>";
            strBody += @"<TR><TH noWrap><SPAN><INPUT id=checkboxAll CHECKED type=checkbox name=orderidAll></SPAN></TH>
                            <TH noWrap><SPAN>公關贈品品項</SPAN></TH>
                            <TH noWrap><SPAN>公關贈品名稱</SPAN></TH>
                            <TH noWrap><SPAN>數量</SPAN></TH>
                            <TH noWrap><SPAN>備註</SPAN></TH>
                        </TR>";

            foreach (DataRow dr in dt.Rows)
            {

                //int Donate_Amt = Convert.ToInt32(dr["捐款金額"].ToString());
                //Donate_Amt_Sum += Donate_Amt;

                strBody += String.Format("<TR align='left'><TD noWrap align='center'><SPAN><INPUT id=checkbox CHECKED type=checkbox value={1} name=orderid_{0}></SPAN></TD>", dr["物品編號"].ToString().Trim(), Convert.ToInt32(dr["物品編號"].ToString()));
                strBody += "<TD noWrap><SPAN>" + dr["公關贈品品項"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["物品名稱"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><input type='text' name='GiftsQty_" + dr["物品編號"].ToString() + "' id='GiftsQty_" + dr["物品編號"].ToString() + "' value='1'></TD>";
                strBody += "<TD noWrap><input type='text' name='GiftsComment_" + dr["物品編號"].ToString() + "' id='GiftsComment_" + dr["物品編號"].ToString() + "'></TD>";
            }

            strBody += "</table>";
            lblGridList.Text = strBody;
            Session["strSql"] = strSql;
            
        }
    }
}