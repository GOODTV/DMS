using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class ContributeMgr_Contribute_Add : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Donor_Id");
            LoadDropDownListData();
            Form_DataBind();
        }
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " ReadOnly();", true);
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //捐款方式
        Util.FillDropDownList(ddlContribute_Payment, Util.GetDataTable("CaseCode", "GroupName", "物品捐贈方式", "", ""), "CaseName", "CaseName", false);
        ddlContribute_Payment.SelectedIndex = 0;

        //捐贈用途
        Util.FillDropDownList(ddlContribute_Purpose, Util.GetDataTable("CaseCode", "GroupName", "物品捐贈用途", "", ""), "CaseName", "CaseName", false);
        ddlContribute_Purpose.SelectedIndex = 0;

        //收據開立
        ddlInvoice_Type.Items.Insert(0, new ListItem("", ""));
        ddlInvoice_Type.Items.Insert(1, new ListItem("不寄", "不寄"));
        ddlInvoice_Type.Items.Insert(2, new ListItem("單次收據", "單次收據"));
        ddlInvoice_Type.SelectedIndex = 2;

        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.SelectedIndex = 0;

        //會計科目
        Util.FillDropDownList(ddlAccounting_Title, Util.GetDataTable("CaseCode", "GroupName", "款項會計科目", "", ""), "CaseName", "CaseName", false);
        ddlAccounting_Title.Items.Insert(0, new ListItem("", ""));
        ddlAccounting_Title.SelectedIndex = 0;

        //募款活動
        Util.FillDropDownList(ddlAct_Id, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        ddlAct_Id.Items.Insert(0, new ListItem("", ""));
        ddlAct_Id.SelectedIndex = 0;
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
        strSql = @" select *, (Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.Invoice_ZipCode+B.mValue+Invoice_Address Else A.mValue+DONOR.Invoice_ZipCode+Invoice_Address End End) as [通訊地址]  
                    from Donor Left Join CODECITY As A On DONOR.Invoice_City=A.mCode Left Join CODECITY As B On DONOR.Invoice_Area=B.mCode 
                    where Donor_id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("ContributeList.aspx");

        DataRow dr = dt.Rows[0];

        //捐款人
        tbxDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
        ////類別
        //tbxCategory.Text = dr["Category"].ToString().Trim();
        //身份別
        tbxDonor_Type.Text = dr["Donor_Type"].ToString().Trim();
        //收據地址
        tbxAddress.Text = dr["通訊地址"].ToString().Trim();
        //收據開立
        tbxInvoice_Type.Text = dr["Invoice_Type"].ToString().Trim();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //最近捐款日：
        if (dr["Last_DonateDateC"].ToString().Trim() != "")
        {
            tbxLast_DonateDateC.Text = DateTime.Parse(dr["Last_DonateDateC"].ToString().Trim()).ToShortDateString().ToString();
        }
        //收據抬頭
        tbxInvoice_Title.Text = dr["Invoice_Title"].ToString().Trim();

        // 2014/4/8 修改可帶入公關贈品輸入的收據開立
        ddlInvoice_Type.SelectedItem.Text = tbxInvoice_Type.Text;

    }
    //----------------------------------------------------------------------
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        bool flag = false;
        string Contribute_Id = "";
        try
        {
            //******新增Contribute的資料******//
            Contribute_AddNew();
            
            //******取得Contribute_Id******//
            string strSql = @"select Contribute_Id from Contribute
                        where Invoice_No ='" + HFD_Invoice_No.Value + "'";
            //****執行查詢流水號語法****//
            DataTable dt = NpoDB.QueryGetTable(strSql);
            //--------
            //資料異常
            if (dt.Rows.Count <= 0)
                //todo : add Default.aspx page
                Response.Redirect("ContributeList.aspx");
            //--------
            DataRow dr = dt.Rows[0];
            Contribute_Id = dr["Contribute_Id"].ToString().Trim();

            bool Contribute_IsStock;
            //******新增ContributeData的資料******//
            //--------------------------------------
            //Goods_1
            if (tbxGoods_Name_1.Text != "")
            {
                if (cbxContribute_IsStock_1.Checked){
                    Contribute_IsStock = true;
                }else{
                    Contribute_IsStock = false;
                }
                ContributeData_AddNew(Contribute_Id, HFD_Goods_Id_1.Value, tbxGoods_Name_1.Text.Trim(), tbxGoods_Qty_1.Text.Trim(), tbxGoods_Unit_1.Text.Trim(), tbxGoods_Amt_1.Text.Trim(), tbxGoods_DueDate_1.Text.Trim(), tbxGoods_Comment_1.Text.Trim(), Contribute_IsStock);
                //******修改Goods的資料******//
                Goods_Edit(tbxGoods_Qty_1.Text.Trim(), HFD_Goods_Id_1.Value);
            }
            //--------------------------------------
            //Goods_2
            if (tbxGoods_Name_2.Text != "")
            {
                if (cbxContribute_IsStock_2.Checked)
                {
                    Contribute_IsStock = true;
                }
                else
                {
                    Contribute_IsStock = false;
                }
                ContributeData_AddNew(Contribute_Id, HFD_Goods_Id_2.Value, tbxGoods_Name_2.Text.Trim(), tbxGoods_Qty_2.Text.Trim(), tbxGoods_Unit_2.Text.Trim(), tbxGoods_Amt_2.Text.Trim(), tbxGoods_DueDate_2.Text.Trim(), tbxGoods_Comment_2.Text.Trim(), Contribute_IsStock);
                //******修改Goods的資料******//
                Goods_Edit(tbxGoods_Qty_2.Text.Trim(), HFD_Goods_Id_2.Value);
            }
            //--------------------------------------
            //Goods_3
            if (tbxGoods_Name_3.Text != "")
            {
                if (cbxContribute_IsStock_3.Checked)
                {
                    Contribute_IsStock = true;
                }
                else
                {
                    Contribute_IsStock = false;
                }
                ContributeData_AddNew(Contribute_Id, HFD_Goods_Id_3.Value, tbxGoods_Name_3.Text.Trim(), tbxGoods_Qty_3.Text.Trim(), tbxGoods_Unit_3.Text.Trim(), tbxGoods_Amt_3.Text.Trim(), tbxGoods_DueDate_3.Text.Trim(), tbxGoods_Comment_3.Text.Trim(), Contribute_IsStock);
                //******修改Goods的資料******//
                Goods_Edit(tbxGoods_Qty_3.Text.Trim(), HFD_Goods_Id_3.Value);
            }
            //--------------------------------------
            //Goods_4
            if (tbxGoods_Name_4.Text != "")
            {
                if (cbxContribute_IsStock_4.Checked)
                {
                    Contribute_IsStock = true;
                }
                else
                {
                    Contribute_IsStock = false;
                }
                ContributeData_AddNew(Contribute_Id, HFD_Goods_Id_4.Value, tbxGoods_Name_4.Text.Trim(), tbxGoods_Qty_4.Text.Trim(), tbxGoods_Unit_4.Text.Trim(), tbxGoods_Amt_4.Text.Trim(), tbxGoods_DueDate_4.Text.Trim(), tbxGoods_Comment_4.Text.Trim(), Contribute_IsStock);
                //******修改Goods的資料******//
                Goods_Edit(tbxGoods_Qty_4.Text.Trim(), HFD_Goods_Id_4.Value);
            }
            //--------------------------------------
            //Goods_5
            if (tbxGoods_Name_5.Text != "")
            {
                if (cbxContribute_IsStock_5.Checked)
                {
                    Contribute_IsStock = true;
                }
                else
                {
                    Contribute_IsStock = false;
                }
                ContributeData_AddNew(Contribute_Id, HFD_Goods_Id_5.Value, tbxGoods_Name_5.Text.Trim(), tbxGoods_Qty_5.Text.Trim(), tbxGoods_Unit_5.Text.Trim(), tbxGoods_Amt_5.Text.Trim(), tbxGoods_DueDate_5.Text.Trim(), tbxGoods_Comment_5.Text.Trim(), Contribute_IsStock);
                //******修改Goods的資料******//
                Goods_Edit(tbxGoods_Qty_5.Text.Trim(), HFD_Goods_Id_5.Value);
            }
            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            //Donor 資料表中的捐贈資料增加
            Donor_Edit();
            SetSysMsg("捐物資料新增成功！");

            if (Session["cType"] == "ContributeList")
            {
                Response.Redirect(Util.RedirectByTime("Contribute_Detail.aspx", "Contribute_Id=" + Contribute_Id));
            }
            else if (Session["cType"] == "ContributeDataList")
            {
                Response.Redirect(Util.RedirectByTime("ContributeDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
            }
            else
            {
                Response.Redirect(Util.RedirectByTime("ContributeList.aspx"));
            }
            
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        if (Session["cType"] == "ContributeList")
        {
            Response.Redirect(Util.RedirectByTime("ContributeList.aspx"));
        }
        else if (Session["cType"] == "ContributeDataList")
        {
            Response.Redirect(Util.RedirectByTime("ContributeDataList.aspx", "Donor_Id=" + HFD_Uid.Value));
        }
        else
        {
            Response.Redirect(Util.RedirectByTime("ContributeList.aspx"));
        }
    }
    public void Contribute_AddNew()
    {
        string strSql = "insert into  Contribute\n";
        strSql += "( Donor_Id, Contribute_Date, Contribute_Payment, Contribute_Purpose, Invoice_Type, Contribute_Amt,Dept_Id, \n";
        strSql += " Invoice_Title, Invoice_Pre, Invoice_No, Invoice_Print, Invoice_Print_Add, Invoice_Print_Yearly, \n";
        strSql += " Invoice_Print_Yearly_Add, Accoun_Date, Accounting_Title, Act_id, Comment, Invoice_PrintComment, Export, \n";
        strSql += " Create_Date, Create_User, Create_IP) values\n";
        strSql += "( @Donor_Id,@Contribute_Date,@Contribute_Payment,@Contribute_Purpose,@Invoice_Type,@Contribute_Amt,@Dept_Id, \n";
        strSql += " @Invoice_Title,@Invoice_Pre,@Invoice_No,@Invoice_Print,@Invoice_Print_Add,@Invoice_Print_Yearly, \n";
        strSql += " @Invoice_Print_Yearly_Add,@Accoun_Date,@Accounting_Title,@Act_id,@Comment,@Invoice_PrintComment,@Export, \n";
        strSql += " @Create_Date,@Create_User,@Create_IP)";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", HFD_Uid.Value);
        dict.Add("Contribute_Date", tbxContribute_Date.Text.Trim());
        dict.Add("Contribute_Payment", ddlContribute_Payment.SelectedItem.Text);
        dict.Add("Contribute_Purpose", ddlContribute_Purpose.SelectedItem.Text);
        dict.Add("Invoice_Type", ddlInvoice_Type.SelectedItem.Text);
        //Contribute_Amt = total Goods_Amt
        int Contribute_Amt = int.Parse(tbxGoods_Amt_1.Text.Trim()) + int.Parse(tbxGoods_Amt_2.Text.Trim()) + int.Parse(tbxGoods_Amt_3.Text.Trim()) + int.Parse(tbxGoods_Amt_4.Text.Trim()) + int.Parse(tbxGoods_Amt_5.Text.Trim());
        dict.Add("Contribute_Amt", Contribute_Amt);
        dict.Add("Dept_Id", ddlDept.SelectedValue);
        dict.Add("Invoice_Title", tbxInvoice_Title.Text.Trim());
        //收據編號
        if (cbxInvoice_Pre.Checked == false)
        {
            dict.Add("Invoice_Pre", "A");
            //******設定收據編號******//
            string strSql2 = @"select isnull(MAX(Invoice_No),'') as Invoice_No from Contribute
                          where Invoice_No like '%" + DateTime.Now.ToString("yyyyMMdd") + "%'";
            //****執行查詢流水號語法****//
            DataTable dt2 = NpoDB.QueryGetTable(strSql2);
            string Invoice_No = "";
            if (dt2.Rows.Count > 0)
            {
                string Invoice_No_value = dt2.Rows[0]["Invoice_No"].ToString();

                if (Invoice_No_value == "")
                {
                    Invoice_No = DateTime.Now.ToString("yyyyMMdd") + "0001";
                }
                else
                {
                    Invoice_No = (Convert.ToInt64(Invoice_No_value) + 1).ToString();
                }
            }
            //************************//
            dict.Add("Invoice_No", Invoice_No);
            HFD_Invoice_No.Value = Invoice_No;
        }
        if (cbxInvoice_Pre.Checked == true)
        {
            dict.Add("Invoice_Pre", "B");
            dict.Add("Invoice_No", tbxInvoice_No.Text.Trim());
            HFD_Invoice_No.Value = tbxInvoice_No.Text.Trim();
        }

        dict.Add("Invoice_Print", "0");
        dict.Add("Invoice_Print_Add", "0");
        dict.Add("Invoice_Print_Yearly", "0");
        dict.Add("Invoice_Print_Yearly_Add", "0");
        dict.Add("Accoun_Date", tbxAccoun_Date.Text.Trim());
        dict.Add("Accounting_Title", ddlAccounting_Title.SelectedItem.Text);
        dict.Add("Act_id", ddlAct_Id.SelectedValue);
        dict.Add("Comment", tbxComment.Text.Trim());
        dict.Add("Invoice_PrintComment", tbxInvoice_PrintComment.Text.Trim());
        dict.Add("Export", "N");
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void ContributeData_AddNew(string Contribute_Id, string Goods_Id, string Goods_Name, string Goods_Qty, string Goods_Unit, string Goods_Amt, string Goods_DueDate, string Goods_Comment, bool Contribute_IsStock)
    {
        string strSql = "insert into ContributeData\n";
        strSql += "( Contribute_Id, Donor_Id, Goods_Id, Goods_Name, Goods_Qty, Goods_Unit,Goods_Amt, \n";
        strSql += " Goods_DueDate, Goods_Comment, Contribute_IsStock, Export, Create_Date, Create_User, \n";
        strSql += " Create_IP) values\n";

        strSql += "( @Contribute_Id,@Donor_Id,@Goods_Id,@Goods_Name,@Goods_Qty,@Goods_Unit,@Goods_Amt, \n";
        strSql += " @Goods_DueDate,@Goods_Comment,@Contribute_IsStock,@Export,@Create_Date, @Create_User, \n";
        strSql += " @Create_IP)";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Contribute_Id", Contribute_Id);
        dict.Add("Donor_Id", HFD_Uid.Value);
        dict.Add("Goods_Id", Goods_Id);
        dict.Add("Goods_Name", Goods_Name);
        dict.Add("Goods_Qty", Goods_Qty);
        dict.Add("Goods_Unit", Goods_Unit);
        dict.Add("Goods_Amt", Goods_Amt);
        dict.Add("Goods_DueDate", Goods_DueDate);
        dict.Add("Goods_Comment", Goods_Comment);
        if (Contribute_IsStock == true)
        {
            dict.Add("Contribute_IsStock", "Y");
        }
        else
        {
            dict.Add("Contribute_IsStock", "N");
        }
        dict.Add("Export", "N");
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void Goods_Edit(string Goods_Qty, string Goods_Id)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Goods set ";
        strSql += "  Goods_Qty = Goods_Qty +'" + Goods_Qty + "'";
        strSql += " where Goods_Id = @Goods_Id";
        dict.Add("Goods_Id", Goods_Id);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void Donor_Edit()
    {
        //先抓出第一次/最近一次的捐贈日期、捐贈次數和捐贈總額
        string Donate_Total_S, Last_DonateDateC;
        int Donate_NoC;
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;
        //****變數設定****//
        uid = HFD_Uid.Value;
        //****設定查詢****//
        strSql = " select *  from Donor  where Donor_Id='" + uid + "'";
        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];

        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        string strSql2 = "";
        if (dr["Begin_DonateDateC"].ToString() == "" || DateTime.Parse(dr["Begin_DonateDateC"].ToString()).ToString("yyyy/MM/dd") == "1900/01/01")
        {
            //總捐贈金額
            int Donate_TotalC = int.Parse(tbxGoods_Amt_1.Text.Trim()) + int.Parse(tbxGoods_Amt_2.Text.Trim()) + int.Parse(tbxGoods_Amt_3.Text.Trim()) + int.Parse(tbxGoods_Amt_4.Text.Trim()) + int.Parse(tbxGoods_Amt_5.Text.Trim());
            //初次捐款
            strSql2 = " update Donor set ";
            strSql2 += "  Begin_DonateDateC = @Begin_DonateDateC";
            strSql2 += ", Last_DonateDateC = @Last_DonateDateC";
            strSql2 += ", Donate_NoC = @Donate_NoC";
            strSql2 += ", Donate_TotalC = @Donate_TotalC";
            strSql2 += " where Donor_Id = @Donor_Id";

            dict2.Add("Begin_DonateDateC", DateTime.Now.ToString("yyyy-MM-dd"));
            dict2.Add("Last_DonateDateC", DateTime.Now.ToString("yyyy-MM-dd"));
            dict2.Add("Donate_NoC", "1");
            dict2.Add("Donate_TotalC", Donate_TotalC);
            dict2.Add("Donor_Id", HFD_Uid.Value);
        }
        else
        {
            Last_DonateDateC = DateTime.Parse(dr["Last_DonateDateC"].ToString()).ToString("yyyy/MM/dd");
            Donate_NoC = Int16.Parse(dr["Donate_NoC"].ToString());
            Donate_Total_S = (Convert.ToInt64(dr["Donate_TotalC"])).ToString();
            int Donate_Total = Int32.Parse(Donate_Total_S);
            int Donate_TotalC = int.Parse(tbxGoods_Amt_1.Text.Trim()) + int.Parse(tbxGoods_Amt_2.Text.Trim()) + int.Parse(tbxGoods_Amt_3.Text.Trim()) + int.Parse(tbxGoods_Amt_4.Text.Trim()) + int.Parse(tbxGoods_Amt_5.Text.Trim());
            //更新以上欄位

            strSql2 = " update Donor set ";
            strSql2 += " Donate_NoC = @Donate_NoC";
            strSql2 += ", Donate_TotalC = @Donate_TotalC";
            if (DateTime.Parse(DateTime.Parse(Last_DonateDateC).ToString("yyyy/MM/dd")) < DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")))
            {
                strSql2 += ", Last_DonateDateC = @Last_DonateDateC";
                dict2.Add("Last_DonateDateC", DateTime.Now.ToString("yyyy-MM-dd"));
            }
            strSql2 += " where Donor_Id = @Donor_Id";

            dict2.Add("Donate_NoC", Donate_NoC + 1);
            dict2.Add("Donate_TotalC", Donate_Total + Donate_TotalC);
            dict2.Add("Donor_Id", HFD_Uid.Value);
        }
        NpoDB.ExecuteSQLS(strSql2, dict2);
    }
}