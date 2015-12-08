using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;

public partial class ContributeMgr_ContributeIssue_Add : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            LoadDropDownListData();
        }
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " ReadOnly();", true);
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.SelectedIndex = 0;

        //領取用途
        Util.FillDropDownList(ddlIssue_Purpose, Util.GetDataTable("CaseCode", "GroupName", "物品領取用途", "", ""), "CaseName", "CaseName", false);
        ddlIssue_Purpose.SelectedIndex = 0;
    }
    //----------------------------------------------------------------------
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        string strRet = CheckStock();
        if (strRet != "")
        {
            ShowSysMsg(strRet);
            return;
        }
        else
        {
            bool flag = false;
            string Issue_Id = "";
            try
            {
                //******新增Contribute的資料******//
                Contribute_AddNew();
                
                //******取得Issue_No******//
                string strSql = @"select Issue_Id from Contribute_Issue
                        where Issue_No ='" + HFD_Issue_No.Value + "'";
                //****執行查詢流水號語法****//
                DataTable dt = NpoDB.QueryGetTable(strSql);
                //--------
                //資料異常
                if (dt.Rows.Count <= 0)
                    //todo : add Default.aspx page
                    Response.Redirect("ContributeIssueList.aspx");
                //--------
                DataRow dr = dt.Rows[0];
                Issue_Id = dr["Issue_Id"].ToString().Trim();

                //******新增ContributeData的資料******//
                bool Contribute_IsStock;
                //--------------------------------------
                //Goods_1
                if (tbxGoods_Name_1.Text != "")
                {
                    if (cbxContribute_IsStock_1.Checked)
                    {
                        Contribute_IsStock = true;
                    }
                    else
                    {
                        Contribute_IsStock = false;
                    }
                    Contribute_IssueData_AddNew(Issue_Id, HFD_Goods_Id_1.Value, tbxGoods_Name_1.Text.Trim(), tbxGoods_Qty_1.Text.Trim(), tbxGoods_Unit_1.Text.Trim(), tbxGoods_Comment_1.Text.Trim(), Contribute_IsStock);
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
                    Contribute_IssueData_AddNew(Issue_Id, HFD_Goods_Id_2.Value, tbxGoods_Name_2.Text.Trim(), tbxGoods_Qty_2.Text.Trim(), tbxGoods_Unit_2.Text.Trim(), tbxGoods_Comment_2.Text.Trim(), Contribute_IsStock);
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
                    Contribute_IssueData_AddNew(Issue_Id, HFD_Goods_Id_3.Value, tbxGoods_Name_3.Text.Trim(), tbxGoods_Qty_3.Text.Trim(), tbxGoods_Unit_3.Text.Trim(), tbxGoods_Comment_3.Text.Trim(), Contribute_IsStock);
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
                    Contribute_IssueData_AddNew(Issue_Id, HFD_Goods_Id_4.Value, tbxGoods_Name_4.Text.Trim(), tbxGoods_Qty_4.Text.Trim(), tbxGoods_Unit_4.Text.Trim(), tbxGoods_Comment_4.Text.Trim(), Contribute_IsStock);
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
                    Contribute_IssueData_AddNew(Issue_Id, HFD_Goods_Id_5.Value, tbxGoods_Name_5.Text.Trim(), tbxGoods_Qty_5.Text.Trim(), tbxGoods_Unit_5.Text.Trim(), tbxGoods_Comment_5.Text.Trim(), Contribute_IsStock);
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
                SetSysMsg("領用資料新增成功！");
                Response.Redirect(Util.RedirectByTime("ContributeIssue_Detail.aspx", "Issue_Id=" + Issue_Id));
            }
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("ContributeIssueList.aspx"));
    }
    public string CheckStock()
    {
        string strRet = "";
        if (tbxGoods_Name_1.Text != "")
        {
            strRet +=  Check(HFD_Goods_Id_1.Value, int.Parse(tbxGoods_Qty_1.Text.Trim()));
        }
        if (tbxGoods_Name_2.Text != "")
        {
            strRet += Check(HFD_Goods_Id_2.Value, int.Parse(tbxGoods_Qty_2.Text.Trim()));
        }
        if (tbxGoods_Name_3.Text != "")
        {
            strRet += Check(HFD_Goods_Id_3.Value, int.Parse(tbxGoods_Qty_3.Text.Trim()));
        }
        if (tbxGoods_Name_4.Text != "")
        {
            strRet += Check(HFD_Goods_Id_4.Value, int.Parse(tbxGoods_Qty_4.Text.Trim()));
        }
        if (tbxGoods_Name_5.Text != "")
        {
            strRet += Check(HFD_Goods_Id_5.Value, int.Parse(tbxGoods_Qty_5.Text.Trim()));
        }
        if (strRet != "")
        {
            strRet += "的物品庫存量不足無法領用";
        }

        return strRet;
    }
    public void Contribute_AddNew()
    {
        string strSql = "insert into  Contribute_Issue\n";
        strSql += "( Dept_Id, Issue_Date, Issue_Processor, Issue_Purpose, Issue_Org, Issue_Comment, Issue_Type\n";
        strSql += " ,Issue_Pre, Issue_No, Issue_Print, Export, Create_Date, Create_User, Create_IP) values\n";
        strSql += "( @Dept_Id,@Issue_Date,@Issue_Processor,@Issue_Purpose,@Issue_Org,@Issue_Comment,@Issue_Type\n";
        strSql += " ,@Issue_Pre,@Issue_No,@Issue_Print,@Export,@Create_Date,@Create_User,@Create_IP)";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Dept_Id", ddlDept.SelectedValue);
        dict.Add("Issue_Date", tbxIssue_Date.Text.Trim());
        dict.Add("Issue_Processor", tbxIssue_Processor.Text.Trim());
        dict.Add("Issue_Purpose", ddlIssue_Purpose.SelectedItem.Text);
        dict.Add("Issue_Org", tbxIssue_Org.Text.Trim());
        dict.Add("Issue_Pre", "領A");
        //收據編號
        if (cbxIssue_Pre.Checked == false)
        {
            dict.Add("Issue_Type", "");
            //******設定收據編號******//
            string strSql2 = @"select isnull(MAX(Issue_No),'') as Issue_No from Contribute_Issue
                          where Issue_No like '%" + DateTime.Now.Year.ToString() + "%'";
            //****執行查詢流水號語法****//
            DataTable dt2 = NpoDB.QueryGetTable(strSql2);
            string Issue_No = "";
            if (dt2.Rows.Count > 0)
            {
                string Issue_No_value = dt2.Rows[0]["Issue_No"].ToString();

                if (Issue_No_value == "")
                {
                    Issue_No = DateTime.Now.Year.ToString() + "001";
                }
                else
                {
                    Issue_No = (Convert.ToInt64(Issue_No_value.Replace("A","")) + 1).ToString();
                }
            }
            //************************//
            dict.Add("Issue_No", Issue_No);
            HFD_Issue_No.Value = Issue_No;
        }
        if (cbxIssue_Pre.Checked == true)
        {
            dict.Add("Issue_Type", "M");
            dict.Add("Issue_No", tbxIssue_No.Text.Trim());
            HFD_Issue_No.Value = tbxIssue_No.Text.Trim();
        }
        dict.Add("Issue_Comment", tbxComment.Text.Trim());
        dict.Add("Issue_Print", "0");
        dict.Add("Export", "N");
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());

        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void Contribute_IssueData_AddNew(string Issue_Id, string Goods_Id, string Goods_Name, string Goods_Qty, string Goods_Unit, string Goods_Comment, bool Contribute_IsStock)
    {
        string strSql = "insert into Contribute_IssueData\n";
        strSql += "( Issue_Id, Goods_Id, Goods_Name, Goods_Qty, Goods_Unit, Goods_Comment \n";
        strSql += " , Contribute_IsStock, Create_Date, Create_User, Create_IP) values\n";
        strSql += "( @Issue_Id,@Goods_Id,@Goods_Name,@Goods_Qty,@Goods_Unit,@Goods_Comment, \n";
        strSql += " @Contribute_IsStock,@Create_Date,@Create_User,@Create_IP)";
        strSql += "\n";
        strSql += "select @@IDENTITY";
        
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Issue_Id", Issue_Id);
        dict.Add("Goods_Id", Goods_Id);
        dict.Add("Goods_Name", Goods_Name);
        dict.Add("Goods_Qty", Goods_Qty);
        dict.Add("Goods_Unit", Goods_Unit);
        dict.Add("Goods_Comment", Goods_Comment);
        if (Contribute_IsStock==true)
        {
            dict.Add("Contribute_IsStock", "Y");
        }
        else
        {
            dict.Add("Contribute_IsStock", "N");
        }
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
        strSql += "  Goods_Qty = Goods_Qty -'" + Goods_Qty + "'";
        strSql += " where Goods_Id = @Goods_Id";
        dict.Add("Goods_Id", Goods_Id);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public string Check(string Goods_Id, int tbxGoods_Qty)
    {
        string strRet = "";
        string strSql = " select *  from Goods where Goods_Id='" + Goods_Id + "'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        int Goods_Qty = int.Parse(dr["Goods_Qty"].ToString());//庫存量
        string Goods_Name = dr["Goods_Name"].ToString();
        if (Goods_Qty < tbxGoods_Qty)
        {
            strRet += "『" + Goods_Name + "』 ";
        }
        return strRet;
    }
}