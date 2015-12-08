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

public partial class ContributeMgr_Contribute_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Util.ShowSysMsgWithScript("");
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Contribute_Id");
            LoadDropDownListData();
            Form_DataBind();
        }
        //紀錄原始領用編號
        HFD_Invoice_No.Value = tbxInvoice_No.Text;
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
        ddlInvoice_Type.Items.Insert(1, new ListItem("不需要", "不需要"));
        ddlInvoice_Type.Items.Insert(2, new ListItem("國外格式(A)", "國外格式(A)"));
        ddlInvoice_Type.Items.Insert(3, new ListItem("國外格式(B)", "國外格式(B)"));
        ddlInvoice_Type.SelectedIndex = 2;

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
        ////類別
        //tbxCategory.Text = dr["Category"].ToString().Trim();
        //身份別
        tbxDonor_Type.Text = dr["Donor_Type"].ToString().Trim();
        //收據地址
        tbxAddress.Text = dr["通訊地址"].ToString().Trim();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //捐贈日期
        tbxContribute_Date.Text = DateTime.Parse(dr["Contribute_Date"].ToString()).ToString("yyyy/MM/dd");
        //捐款方式
        ddlContribute_Payment.Text = dr["Contribute_Payment"].ToString().Trim();
        //捐款用途
        ddlContribute_Purpose.Text = dr["Contribute_Purpose"].ToString().Trim();
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
        ddlAccounting_Title.Text = dr["Accounting_Title"].ToString().Trim();
        //專案活動
        ddlAct_Id.Text = dr["Act_ShortName"].ToString().Trim();
        //捐款備註
        tbxComment.Text = dr["Comment"].ToString().Trim();
        //收據備註
        tbxInvoice_PrintComment.Text = dr["Invoice_PrintComment"].ToString().Trim();

        //載入捐贈內容
        string HLink, LinkParam, LinkTarget, DataLink, DataParam, DataTarget, strSql2, uid2 = "";
        DataTable dt2;
        uid2 = HFD_Uid.Value;
        strSql2 = " select Ser_No as [Ser_No], ROW_NUMBER() OVER(ORDER BY Ser_No) as [序號], \n";
        strSql2 += " CD.Goods_Name as [物品名稱], CD.Goods_Qty as [數量] , CD.Goods_Unit as [單位], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Goods_Amt),1),'.00','')  as [折合現金], \n";
        strSql2 += " (case when Goods_DueDate!='' or CONVERT(VARCHAR(10) , Goods_DueDate, 111 )!='1900/01/01' then CONVERT(VARCHAR(10), \n";
        strSql2 += " Goods_DueDate, 111 ) else '' end)  as [保存期限], Goods_Comment as [備註], \n";
        strSql2 += " (case when Contribute_IsStock='Y' then 'V' else '' end) as 寫入庫存\n";
        strSql2 += " from ContributeData CD right join Goods G on CD.Goods_Id = G.Goods_Id \n";
        strSql2 += " where Contribute_Id='" + uid2 + "'\n";
        strSql2 += " union \n";
        strSql2 += " select '9999',null,'',null,'折合現金合計：',REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,sum (case when Goods_Amt!='0' then Goods_Amt else '0' end)),1),'.00','') ,'','',''\n";
        strSql2 += " from ContributeData where Contribute_Id='" + uid2 + "'";

        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        dt2 = NpoDB.GetDataTableS(strSql2, dict2);

        HLink = "ContributeData_Edit.aspx?Ser_No=";
        LinkParam = "Ser_No";
        LinkTarget = "_blank";
        DataLink = "ContributeData_Delete.aspx?Ser_No=";
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
            if (dr[LinkParam].ToString() == "9999")
            {
                //折合現金合計不能做刪除和修改的動作
            }
            else
            {
                //sb.AppendLine(@"<td><a href='" + HLink + dr[LinkParam] + @"' target='" + LinkTarget + @"'><img src='../images/save.gif' border=0 align='absmiddle' alt='修改'>修改</a><a href='JavaScript:if(confirm(""是否確定要刪除 ?"")){window.location.href=""" + DataLink + dr[DataParam] + @""";}' target='" + DataTarget + @"'><img src='../images/delete2.gif' align='absmiddle' border=0 alt='刪除'></a></td>");
                sb.AppendLine(@"<td><a href onclick=""window.open('" + HLink + dr[LinkParam] + @"','ContributeData_Edit','status=no,scrollbars=no,top=100,left=120,width=500,height=250')""><img src='../images/save.gif' border=0 align='absmiddle' alt='編輯'>編輯</a></td>");
                //sb.AppendLine(@"<td><a href='" + HLink + dr[LinkParam] + @"' target='" + LinkTarget + @"'><img src='../images/save.gif' border=0 align='absmiddle' alt='修改'>修改</a>");
                sb.AppendLine(@"<td><a href='JavaScript:if(confirm(""是否確定要刪除 ?"")){window.location.href=""" + DataLink + dr[DataParam] + @""";}' target='" + DataTarget + @"'><img src='../images/delete2.gif' align='absmiddle' border=0 alt='刪除'></a></td>");
            }
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</table>");
        return sb.ToString();
    }
    //---------------------------------------------------------------------
    protected void btnEdit_Click(object sender, EventArgs e)
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
            SetSysMsg("捐贈資料修改成功！");
            if (Session["cType"] == "ContributeList")
            {
                Response.Redirect(Util.RedirectByTime("Contribute_Detail.aspx", "Contribute_Id=" + HFD_Uid.Value));
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
    }
    protected void btnDel_Click(object sender, EventArgs e)
    {
        //找出Contribute表中的總金額
        string strSql = "select Contribute_Amt from Contribute where Contribute_Id='" + HFD_Uid.Value +"'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Contribute_Amt_S = (Convert.ToInt64(dr["Contribute_Amt"])).ToString();
        int Contribute_Amt = int.Parse(Contribute_Amt_S);
        
        //刪除Contribute表的資料
        string strSql2 = "delete from Contribute where Contribute_Id=@Contribute_Id";
        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        dict2.Add("Contribute_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql2, dict2);

        //找出ContributeData有幾筆資料
        //****設定查詢****//
        string strSql3 = " select *  from ContributeData where Contribute_Id='" + HFD_Uid.Value + "'";

        //****執行語法****//
        DataTable dt3;
        dt3 = NpoDB.QueryGetTable(strSql);
        for (int i = 0; i < dt3.Rows.Count; i++)
        {
            DataRow dr3 = dt3.Rows[i];
            string Goods_Id = dr3["Goods_Id"].ToString();
            string Goods_Qty = dr3["Goods_Qty"].ToString();

            //******修改Goods的資料******//
            Dictionary<string, object> dict_Goods = new Dictionary<string, object>();
            //****設定SQL指令****//
            string strSql_Goods = " update Goods set ";
            strSql_Goods += "  Goods_Qty = Goods_Qty +'" + Goods_Qty + "'";
            strSql_Goods += " where Goods_Id = @Goods_Id";
            dict_Goods.Add("Goods_Id", Goods_Id );
            NpoDB.ExecuteSQLS(strSql_Goods, dict_Goods);
        }

        //刪除ContributeData表的資料
        string strSql4 = "delete from ContributeData where Contribute_Id=@Contribute_Id";
        Dictionary<string, object> dict4 = new Dictionary<string, object>();
        dict4.Add("Contribute_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql4, dict4);

        //******修改Donor的資料******//
        Dictionary<string, object> dict_Donor = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql_Donor = " update Donor set ";
        strSql_Donor += "  Donate_NoC = Donate_NoC -1";


        //找尋同個人是否有其他筆捐款紀錄
        string strSql5 = @"select isnull(MAX(CONVERT(NVARCHAR, Create_Date, 111)),'') as Create_Date from Contribute ";
        strSql5 += " where Donor_Id='" + HFD_Donor_Id.Value + "'";
        DataTable dt5 = NpoDB.GetDataTableS(strSql5, null);
        DataRow dr5 = dt5.Rows[0];
        //找不到其他筆
        if (dr5["Create_Date"].ToString() == "")
        {
            strSql_Donor += "  ,Begin_DonateDateC = ''";
            strSql_Donor += "  ,Last_DonateDateC = ''";
        }
        else
        {
            strSql_Donor += "  ,Last_DonateDateC = '" + dr5["Create_Date"].ToString() + "'";
        }
        strSql_Donor += "  ,Donate_TotalC = Donate_TotalC -'" + Contribute_Amt + "'";
        strSql_Donor += " where Donor_Id = @Donor_Id";
        dict_Donor.Add("Donor_Id", HFD_Donor_Id.Value);
        NpoDB.ExecuteSQLS(strSql_Donor, dict_Donor);

        SetSysMsg("捐贈資料刪除成功！");
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
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("Contribute_Detail.aspx", "Contribute_Id=" + HFD_Uid.Value));
    }
    public void Edit()
    {
        bool Contribute_IsStock;
        Contribute_Edit();
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
            ContributeData_AddNew(HFD_Goods_Id_1.Value, tbxGoods_Name_1.Text.Trim(), tbxGoods_Qty_1.Text.Trim(), tbxGoods_Unit_1.Text.Trim(), tbxGoods_Amt_1.Text.Trim(), tbxGoods_DueDate_1.Text.Trim(), tbxGoods_Comment_1.Text.Trim(), Contribute_IsStock);
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
            ContributeData_AddNew(HFD_Goods_Id_2.Value, tbxGoods_Name_2.Text.Trim(), tbxGoods_Qty_2.Text.Trim(), tbxGoods_Unit_2.Text.Trim(), tbxGoods_Amt_2.Text.Trim(), tbxGoods_DueDate_2.Text.Trim(), tbxGoods_Comment_2.Text.Trim(), Contribute_IsStock);
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
            ContributeData_AddNew(HFD_Goods_Id_3.Value, tbxGoods_Name_3.Text.Trim(), tbxGoods_Qty_3.Text.Trim(), tbxGoods_Unit_3.Text.Trim(), tbxGoods_Amt_3.Text.Trim(), tbxGoods_DueDate_3.Text.Trim(), tbxGoods_Comment_3.Text.Trim(), Contribute_IsStock);
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
            ContributeData_AddNew(HFD_Goods_Id_4.Value, tbxGoods_Name_4.Text.Trim(), tbxGoods_Qty_4.Text.Trim(), tbxGoods_Unit_4.Text.Trim(), tbxGoods_Amt_4.Text.Trim(), tbxGoods_DueDate_4.Text.Trim(), tbxGoods_Comment_4.Text.Trim(), Contribute_IsStock);
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
            ContributeData_AddNew(HFD_Goods_Id_5.Value, tbxGoods_Name_5.Text.Trim(), tbxGoods_Qty_5.Text.Trim(), tbxGoods_Unit_5.Text.Trim(), tbxGoods_Amt_5.Text.Trim(), tbxGoods_DueDate_5.Text.Trim(), tbxGoods_Comment_5.Text.Trim(), Contribute_IsStock);
            //******修改Goods的資料******//
            Goods_Edit(tbxGoods_Qty_5.Text.Trim(), HFD_Goods_Id_5.Value);
        }

        int Donate_TotalC = int.Parse(tbxGoods_Amt_1.Text.Trim()) + int.Parse(tbxGoods_Amt_2.Text.Trim()) + int.Parse(tbxGoods_Amt_3.Text.Trim()) + int.Parse(tbxGoods_Amt_4.Text.Trim()) + int.Parse(tbxGoods_Amt_5.Text.Trim());
        //******修改Donor的Donate_TotalC******//
        Dictionary<string, object> dict_Donor = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql_Donor = " update Donor set ";
        strSql_Donor += "  Donate_TotalC = Donate_TotalC +'" + Donate_TotalC + "'";
        strSql_Donor += " where Donor_Id = @Donor_Id";
        dict_Donor.Add("Donor_Id", HFD_Donor_Id.Value);
        NpoDB.ExecuteSQLS(strSql_Donor, dict_Donor);

        //******修改Contribute的Contribute_Amt******//
        //****變數宣告****//
        Dictionary<string, object> dict_Contribute = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql_Contribute = " update Contribute set ";
        strSql_Contribute += " Contribute_Amt = Contribute_Amt +'" + Donate_TotalC + "'";
        strSql_Contribute += " where Contribute_Id = @Contribute_Id";
        dict_Contribute.Add("Contribute_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql_Contribute, dict_Contribute);
    }
    public void Contribute_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update Contribute set ";

        strSql += "  Contribute_Date = @Contribute_Date";
        strSql += ", Contribute_Payment = @Contribute_Payment";
        strSql += ", Contribute_Purpose = @Contribute_Purpose";
        strSql += ", Invoice_Type = @Invoice_Type";
        strSql += ", Invoice_Title = @Invoice_Title";
        strSql += ", Invoice_Pre = @Invoice_Pre";
        strSql += ", Invoice_No = @Invoice_No";
        strSql += ", Accoun_Date = @Accoun_Date";
        strSql += ", Accounting_Title = @Accounting_Title";
        strSql += ", Act_Id = @Act_Id";
        strSql += ", Comment = @Comment";
        strSql += ", Invoice_PrintComment= @Invoice_PrintComment";
        strSql += ", Issue_Type = @Issue_Type";
        strSql += " where Contribute_Id = @Contribute_Id";

        dict.Add("Contribute_Date", Util.DateTime2String(tbxContribute_Date.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty));
        dict.Add("Contribute_Payment", ddlContribute_Payment.SelectedItem.Text);
        dict.Add("Contribute_Purpose", ddlContribute_Purpose.SelectedItem.Text);
        dict.Add("Invoice_Type", ddlInvoice_Type.SelectedItem.Text);
        dict.Add("Invoice_Title", tbxInvoice_Title.Text.Trim());
        dict.Add("Invoice_Pre", "物A");
        if (cbxInvoice_Pre.Checked == false)
        {
            dict.Add("Issue_Type", "");
            dict.Add("Invoice_No", tbxInvoice_No.Text.Trim().Replace("A", "").Replace("物",""));
        }
        if (cbxInvoice_Pre.Checked == true)
        {
            dict.Add("Issue_Type", "M");
            dict.Add("Invoice_No", tbxInvoice_No.Text.Trim());
        }
        dict.Add("Accoun_Date", Util.DateTime2String(tbxAccoun_Date.Text.Trim(), DateType.yyyyMMdd, EmptyType.ReturnEmpty));
        dict.Add("Accounting_Title", ddlAccounting_Title.SelectedItem.Text);
        dict.Add("Act_Id", ddlAct_Id.SelectedValue);
        dict.Add("Comment", tbxComment.Text.Trim());
        dict.Add("Invoice_PrintComment", tbxInvoice_PrintComment.Text.Trim());

        dict.Add("Contribute_Id", HFD_Uid.Value);
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
    public void ContributeData_AddNew(string Goods_Id, string Goods_Name, string Goods_Qty, string Goods_Unit, string Goods_Amt, string Goods_DueDate, string Goods_Comment, bool Contribute_IsStock)
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
        dict.Add("Contribute_Id", HFD_Uid.Value);
        dict.Add("Donor_Id", HFD_Donor_Id.Value);
        dict.Add("Goods_Id",Goods_Id);
        dict.Add("Goods_Name", Goods_Name);
        dict.Add("Goods_Qty", Goods_Qty);
        dict.Add("Goods_Unit", Goods_Unit);
        dict.Add("Goods_Amt", Goods_Amt);
        dict.Add("Goods_DueDate", Util.DateTime2String(Goods_DueDate, DateType.yyyyMMdd, EmptyType.ReturnEmpty));
        dict.Add("Goods_Comment", Goods_Comment);
        if (Contribute_IsStock == true )
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
}