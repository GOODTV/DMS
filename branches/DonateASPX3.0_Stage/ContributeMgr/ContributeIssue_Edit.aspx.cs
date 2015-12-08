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
using System.Web.UI.WebControls.WebParts;
using System.Web.Security;

public partial class ContributeMgr_ContributeIssue_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            HFD_Uid.Value = Util.GetQueryString("Issue_Id");
            LoadDropDownListData();
            Form_DataBind();
        }
        //紀錄原始收據編號
        HFD_Issue_No.Value = tbxIssue_No.Text;
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " ReadOnly();", true);
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);

        //領取用途
        Util.FillDropDownList(ddlIssue_Purpose, Util.GetDataTable("CaseCode", "GroupName", "物品領取用途", "", ""), "CaseName", "CaseName", false);
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
                    from Contribute_Issue CI left join Dept D on CI.Dept_Id = D.DeptID 
                    where CI.Issue_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);

        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("ContributeIssueList.aspx");

        DataRow dr = dt.Rows[0];

        //機構
        ddlDept.Text = dr["DeptShortName"].ToString().Trim();
        //領取人
        tbxIssue_Processor.Text = dr["Issue_Processor"].ToString().Trim();
        //領用日期
        tbxIssue_Date.Text = DateTime.Parse(dr["Issue_Date"].ToString()).ToString("yyyy/MM/dd");
        //領用用途
        ddlIssue_Purpose.Text = dr["Issue_Purpose"].ToString().Trim();
        //出貨單位
        tbxIssue_Org.Text = dr["Issue_Org"].ToString().Trim();
        //收據號碼
        tbxIssue_No.Text = dr["Issue_Pre"].ToString().Trim() + dr["Issue_No"].ToString().Trim();
        //捐款備註
        tbxComment.Text = dr["Issue_Comment"].ToString().Trim();

        //載入領用內容
        string HLink, LinkParam, LinkTarget, DataLink, DataParam, DataTarget, strSql2, uid2 = "";
        DataTable dt2;
        uid2 = HFD_Uid.Value;
        strSql2 = " select Ser_No as [Ser_No], ROW_NUMBER() OVER(ORDER BY Ser_No) as [序號], \n";
        strSql2 += " CID.Goods_Name as [物品名稱], CID.Goods_Qty as [數量] , CID.Goods_Unit as [單位], Goods_Comment as [備註], \n";
        strSql2 += " (case when Contribute_IsStock='Y' then 'V' else '' end) as [庫存品]\n";
        strSql2 += " from Contribute_IssueData CID right join Goods G on CID.Goods_Id = G.Goods_Id\n";
        strSql2 += " where Issue_Id='" + uid2 + "'";

        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        dt2 = NpoDB.GetDataTableS(strSql2, dict2);

        HLink = "ContributeIssueData_Edit.aspx?Ser_No=";
        LinkParam = "Ser_No";
        LinkTarget = "_blank";
        DataLink = "ContributeIssueData_Delete.aspx?Ser_No=";
        DataParam = "Ser_No";
        DataTarget = "_self";

        lblGridList.Text = DataLinkGrids(dt2, HLink, LinkParam, LinkTarget, DataLink, DataParam, DataTarget); 
    }
    //增加選項欄位
    public string DataLinkGrids(DataTable dt, string HLink, string LinkParam, string LinkTarget, string DataLink, string DataParam, string DataTarget)
    {
        int i;
        StringBuilder sb = new StringBuilder();
        sb.AppendLine(@"<table width='100%' align='center' class='table_h'>");
        sb.AppendLine(@"<tr>");
        i = 0;
        foreach (DataColumn dc in dt.Columns)
        {
            if (i < 1)
            {
                i++;
                continue;

            }
            sb.AppendLine(@"<th>" + dc.ColumnName + "</th>");
        }
        //sb.AppendLine(@"<th>選項</th>");
        sb.AppendLine(@"<th>編輯</th>");
        sb.AppendLine(@"<th>刪除</th>");
        sb.AppendLine("</tr>");

        foreach (DataRow dr in dt.Rows)
        {
            i = 0;
            sb.AppendLine(@"<tr>");
            foreach (DataColumn dc in dt.Columns)
            {
                if (i < 1)
                {
                    i++;
                    continue;
                }
                else
                {
                    //顯示資料
                    sb.AppendLine("<td>" + dr[dc.ColumnName] + "</td>");
                }
                i++;
            }
            //sb.AppendLine(@"<td><a href='" + HLink + dr[LinkParam] + @"' target='" + LinkTarget + @"'><img src='../images/save.gif' border=0 align='absmiddle' alt='修改'>修改</a><a href='JavaScript:if(confirm(""是否確定要刪除 ?"")){window.location.href=""" + DataLink + dr[DataParam] + @""";}' target='" + DataTarget + @"'><img src='../images/delete2.gif' align='absmiddle' border=0 alt='刪除'></a></td>");
            sb.AppendLine(@"<td><a href onclick=""window.open('" + HLink + dr[LinkParam] + @"','ContributeIssueData_Edit','status=no,scrollbars=no,top=100,left=120,width=500,height=250')""><img src='../images/save.gif' border=0 align='absmiddle' alt='編輯'>編輯</a></td>");
            //sb.AppendLine(@"<td><a href='" + HLink + dr[LinkParam] + @"' target='" + LinkTarget + @"'><img src='../images/save.gif' border=0 align='absmiddle' alt='修改'>修改</a>");
            sb.AppendLine(@"<td><a href='JavaScript:if(confirm(""是否確定要刪除 ?"")){window.location.href=""" + DataLink + dr[DataParam] + @""";}' target='" + DataTarget + @"'><img src='../images/delete2.gif' align='absmiddle' border=0 alt='刪除'></a></td>");
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</table>");
        return sb.ToString();
    }
    //---------------------------------------------------------------------
    protected void btnEdit_Click(object sender, EventArgs e)
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
             try
             {
                 Edit();
                 flag = true;
             }
             catch (Exception ex)
             {
                 throw ex;
             }
             if (flag == true)
             {
                 SetSysMsg("修改成功！");
                 Response.Redirect(Util.RedirectByTime("ContributeIssue_Detail.aspx", "Issue_Id=" + HFD_Uid.Value));
             }
         }
    }
    protected void btnDel_Click(object sender, EventArgs e)
    {
        //刪除Contribute_Issue表的資料
        string strSql = "delete from Contribute_Issue where Issue_Id=@Issue_Id";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Issue_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);

        //找出Contribute_IssueData有幾筆資料
        //****設定查詢****//
        strSql = " select *  from Contribute_IssueData where Issue_Id='" + HFD_Uid.Value + "'";

        //****執行語法****//
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            DataRow dr = dt.Rows[i];
            string Goods_Id = dr["Goods_Id"].ToString();
            string Goods_Qty = dr["Goods_Qty"].ToString();

            //******修改Goods的資料******//
            Dictionary<string, object> dict_Goods = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Goods = " update Goods set ";
            strSql_Goods += "  Goods_Qty = Goods_Qty +'" + Goods_Qty + "'";
            strSql_Goods += " where Goods_Id = @Goods_Id";
            dict_Goods.Add("Goods_Id", Goods_Id);
            NpoDB.ExecuteSQLS(strSql_Goods, dict_Goods);
        }

        //刪除ContributeData表的資料
        string strSql2 = "delete from Contribute_IssueData where Issue_Id=@Issue_Id";
        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        dict2.Add("Issue_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql2, dict2);

        SetSysMsg("領用資料刪除成功！");
        Response.Redirect(Util.RedirectByTime("ContributeIssueList.aspx"));
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("ContributeIssue_Detail.aspx", "Issue_Id=" + HFD_Uid.Value));
    }
    public void Edit()
    {
        Contribute_Issue_AddNew();

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
            Contribute_IssueData_AddNew(HFD_Goods_Id_1.Value, tbxGoods_Name_1.Text.Trim(), tbxGoods_Qty_1.Text.Trim(), tbxGoods_Unit_1.Text.Trim(), tbxGoods_Comment_1.Text.Trim(), Contribute_IsStock);
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
            Contribute_IssueData_AddNew(HFD_Goods_Id_2.Value, tbxGoods_Name_2.Text.Trim(), tbxGoods_Qty_2.Text.Trim(), tbxGoods_Unit_2.Text.Trim(), tbxGoods_Comment_2.Text.Trim(), Contribute_IsStock);
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
            Contribute_IssueData_AddNew(HFD_Goods_Id_3.Value, tbxGoods_Name_3.Text.Trim(), tbxGoods_Qty_3.Text.Trim(), tbxGoods_Unit_3.Text.Trim(), tbxGoods_Comment_3.Text.Trim(), Contribute_IsStock);
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
            Contribute_IssueData_AddNew(HFD_Goods_Id_4.Value, tbxGoods_Name_4.Text.Trim(), tbxGoods_Qty_4.Text.Trim(), tbxGoods_Unit_4.Text.Trim(), tbxGoods_Comment_4.Text.Trim(), Contribute_IsStock);
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
            Contribute_IssueData_AddNew(HFD_Goods_Id_5.Value, tbxGoods_Name_5.Text.Trim(), tbxGoods_Qty_5.Text.Trim(), tbxGoods_Unit_5.Text.Trim(), tbxGoods_Comment_5.Text.Trim(), Contribute_IsStock);
            //******修改Goods的資料******//
            Goods_Edit(tbxGoods_Qty_5.Text.Trim(), HFD_Goods_Id_5.Value);
        }
    }
    public string CheckStock()
    {
        string strRet = "";
        if (tbxGoods_Name_1.Text != "")
        {
            strRet += Check(HFD_Goods_Id_1.Value, int.Parse(tbxGoods_Qty_1.Text.Trim()));
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
    public void Contribute_Issue_AddNew()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update Contribute_Issue set ";

        strSql += "  Dept_Id = @Dept_Id";
        strSql += ", Issue_Processor = @Issue_Processor";
        strSql += ", Issue_Date = @Issue_Date";
        strSql += ", Issue_Purpose = @Issue_Purpose";
        strSql += ", Issue_Org = @Issue_Org";
        strSql += ", Issue_Comment = @Issue_Comment";
        strSql += ", Issue_Pre = @Issue_Pre";
        strSql += ", Issue_No = @Issue_No";
        strSql += ", Issue_Type = @Issue_Type";
        strSql += " where Issue_Id = @Issue_Id";

        dict.Add("Dept_Id", ddlDept.SelectedValue);
        dict.Add("Issue_Processor", tbxIssue_Processor.Text.Trim());
        dict.Add("Issue_Date", tbxIssue_Date.Text.Trim());
        dict.Add("Issue_Purpose", ddlIssue_Purpose.SelectedItem.Text);
        dict.Add("Issue_Org", tbxIssue_Org.Text.Trim());
        dict.Add("Issue_Comment", tbxComment.Text.Trim());
        dict.Add("Issue_Pre", "領A");
        if (cbxIssue_Pre.Checked == false)
        {
            dict.Add("Issue_Type", "");
            dict.Add("Issue_No", tbxIssue_No.Text.Trim().Replace("A",""));
        }
        if (cbxIssue_Pre.Checked == true)
        {
            dict.Add("Issue_Type", "M");
            dict.Add("Issue_No", tbxIssue_No.Text.Trim());
        }

        dict.Add("Issue_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void Contribute_IssueData_AddNew(string Goods_Id, string Goods_Name, string Goods_Qty, string Goods_Unit, string Goods_Comment, bool Contribute_IsStock)
    {
        string strSql = "insert into Contribute_IssueData\n";
        strSql += "( Issue_Id, Goods_Id, Goods_Name, Goods_Qty, Goods_Unit, Goods_Comment \n";
        strSql += " , Contribute_IsStock, Create_Date, Create_User, Create_IP) values\n";
        strSql += "( @Issue_Id,@Goods_Id,@Goods_Name,@Goods_Qty,@Goods_Unit,@Goods_Comment, \n";
        strSql += " @Contribute_IsStock,@Create_Date,@Create_User,@Create_IP)";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Issue_Id", HFD_Uid.Value);
        dict.Add("Goods_Id", Goods_Id);
        dict.Add("Goods_Name", Goods_Name);
        dict.Add("Goods_Qty", Goods_Qty);
        dict.Add("Goods_Unit", Goods_Unit);
        dict.Add("Goods_Comment", Goods_Comment);
        if (Contribute_IsStock == true)
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