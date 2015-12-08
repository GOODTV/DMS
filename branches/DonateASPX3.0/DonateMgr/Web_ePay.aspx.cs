using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class DonateMgr_Web_ePay : System.Web.UI.Page
{

    string strSqlWhere = "";
    string strOrderIds = "";

    protected void Page_Load(object sender, EventArgs e)
    {

        Session["ProgID"] = "Web_ePay";
        ////有 npoGridView 時才需要
        //Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
        
        if (!IsPostBack)
        {
            lblQueryCnt.Text = "";
            btnInput.Enabled = false;
            lblGridList.Text = "** 請先輸入查詢條件 **";
            lblGridList.ForeColor = System.Drawing.Color.Red;
        }

    }

    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }

    protected void btnToxls_Click(object sender, EventArgs e)
    {
        //Response.Redirect("DonateInfo_Print_Excel.aspx");
    }

    protected void btnInput_Click(object sender, EventArgs e)
    {

        strOrderIds = "";
        string strId = "";

        int count = 0; // 合計筆數
        foreach (string str in Request.Form.Keys)
        {
            if (str.StartsWith("orderid") && str != "orderidAll")
            {
                count++;
                strId = str.Split('_')[1];//將勾選的Pledge_Id串成字串
                if (count == 1) strOrderIds = strId;
                else strOrderIds += "','" + strId;
            }
        }

        Dictionary<string, object> dict = new Dictionary<string, object>();
        //dict.Add("OrderIds", "'"+strOrderIds+"'");
        dict.Add("Invoice_No", GetInvoiceNo());
        dict.Add("Donate_Date", txtDonateDate.Text.Trim());

        string strSql = @"
            INSERT INTO [dbo].[DONATE] ([Donor_Id],[Donate_Date],[Donate_Amt],[Donate_AmtD],[Donate_Fee],[Donate_FeeD]
            ,[Donate_Accou],[Donate_AccouD],[Donate_AmtM],[Donate_Payment],[Donate_Purpose],[Donate_FeeM],[Donate_AccouM]
            ,[Donate_AmtA],[Donate_AccouA],[Donate_FeeA],[Donate_AmtS],[Donate_FeeS],[Donate_RateS],[Donate_AccouS]
            ,[Donate_Purpose_Type],[Invoice_Type],[IsBeductible],[Donate_Amt2],[Dept_Id],[Invoice_Title],[Invoice_Pre]
            ,[Invoice_No],[Invoice_Print],[Invoice_Print_Add],[Invoice_Print_Yearly_Add],[Accoun_Bank],[Donate_Type]
            ,[Issue_Type],[Issue_Type_Keep],[od_sob],[Export],[Create_Date],[Create_DateTime],[Create_User],[Create_IP])
            SELECT W.[Donor_Id],@Donate_Date as Donate_Date,account as Donate_Amt
            ,account as Donate_AmtD,0 as Donate_Fee,0 as Donate_FeeD,account as Donate_Accou
            ,account as Donate_AccouD,0 as Donate_AmtM,'網路信用卡' as Donate_Payment,W.Donate_Purpose
            ,0 as Donate_FeeM,0 as Donate_AccouM,0 as Donate_AmtA,0 as Donate_AccouA,0 as Donate_FeeA
            ,0 as Donate_AmtS,0 as Donate_FeeS,0 as Donate_RateS,0 as Donate_AccouS,'D' as Donate_Purpose_Type
            ,[Donate_Invoice_Type] as Invoice_Type,'N' as IsBeductible,0 as Donate_Amt2,W.[Dept_Id]
            ,[Donate_Invoice_Title] as Invoice_Title,(select Invoice_Pre from Dept where DeptId = W.[Dept_Id]) as Invoice_Pre
            ,cast(@Invoice_No as numeric)-1+row_number() over (order by W.od_sob)  as [Invoice_No]
            ,0 as Invoice_Print,0 as Invoice_Print_Add,'0' as Invoice_Print_Yearly_Add,'國泰世華銀行' as Accoun_Bank
            ,W.[Donate_Type],'' as Issue_Type,'' as Issue_Type_Keep,W.[od_sob],'N' as Export,getdate() as Create_Date
            ,getdate() as Create_DateTime,'線上金流' as Create_User,[Donate_CreateIP] as Create_IP
            From (select * from (select ROW_NUMBER() OVER(PARTITION by od_sob ORDER BY Donate_Purpose) AS ROWID,*
                from DONATE_WEB where [IsDelete] is null) as W1 where ROWID = 1) as W
            join (select * from (select ROW_NUMBER() OVER(PARTITION by orderid ORDER BY Ser_No) AS ROWID,*
                from DONATE_IEPAY) as P1 where ROWID = 1) as P on W.od_sob = P.orderid and P.[Status] = '0' and paytype <> '1'
			where od_sob in ('" + strOrderIds + "') ;";

        int intsqlcnt = NpoDB.ExecuteSQLS(strSql, dict);
        //int intsqlcnt = 1;    //test

        if (intsqlcnt > 0)
        {

            //strSql = @"select count(account) as Donate_cnt,sum(account) as Donate_Amt
            //        from DONATE_WEB where orderid in ('" + strOrderIds + "') ;";
            strSql = @"select count(Donate_Amt) as Donate_Cnt,sum(Donate_Amt) as Donate_Amt
                    from donate where od_sob in ('" + strOrderIds + "') ;";
            DataTable dt = NpoDB.QueryGetTable(strSql);

            int Total_cnt = 0;
            decimal Total_Amount = 0;

            DataRow dr = dt.Rows[0];
            Total_cnt = (int)dr["Donate_Cnt"];
            if (Total_cnt > 0)
            {
                Total_Amount = (decimal)dr["Donate_Amt"];
                string msg = String.Format("轉入捐款記錄成功！\\n轉入筆數：{0:#,0} \\n轉入捐款金額：{1:#,0} 元", Total_cnt, Total_Amount);
                this.Page.RegisterStartupScript("s", "<script>alert('" + msg + "');</script>");
                UpdateDonor();

            }
            else
            {
                this.Page.RegisterStartupScript("s", "<script>alert('轉捐款記錄失敗！');</script>");
                //Util.ShowMsg("轉捐款記錄失敗！");
            }

            //lblCheckCnt.Text = strSql;
        }
        else
        {
            this.Page.RegisterStartupScript("s", "<script>alert('轉捐款記錄失敗！');</script>");
            //Util.ShowMsg("轉捐款記錄失敗！");
        }
        LoadFormData();

    }

    private string GetInvoiceNo()
    {
        string strRet = "";
        string strInvoicePre = Util.GetDBValue("Dept", "Invoice_Pre", "Dept_Id", "C001");
        // 2014/6/4 修正SQL語法
        //string strSql = "select top 1 invoice_no from donate where invoice_no like @InvoiceNo";
        string strSql = "select top 1 invoice_no from donate where invoice_no like @InvoiceNo order by invoice_no desc";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        // 2014/9/3 修正收據編號定義 by 詩儀
        //dict.Add("InvoiceNo", strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd") + "%");
        //dict.Add("InvoiceNo", Util.GetDBDateTime().ToString("yyyyMMdd") + "%");
        dict.Add("InvoiceNo", DateTime.Parse(txtDonateDate.Text).ToString("yyyyMMdd") + "%");

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            //取得當日最大流水號
            // 2014/6/4 修正收據編號定義
            //string strSN = dt.Rows[0]["invoice_no"].ToString().Replace(strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd"), "");
            //string strSN = dt.Rows[0]["invoice_no"].ToString().Replace(Util.GetDBDateTime().ToString("yyyyMMdd"), "");
            string strSN = dt.Rows[0]["invoice_no"].ToString().Replace(DateTime.Parse(txtDonateDate.Text).ToString("yyyyMMdd"), "");
            //取得當日最新流水號
            strSN = (Convert.ToInt16(strSN) + 1).ToString("0000");
            //20140515修正InvoiceNo，不加Invoice_Pre
            //strRet = strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd") + strSN;
            //strRet = Util.GetDBDateTime().ToString("yyyyMMdd") + strSN;
            strRet = DateTime.Parse(txtDonateDate.Text).ToString("yyyyMMdd") + strSN;
        }
        else
        {
            //strRet = strInvoicePre + Util.GetDBDateTime().ToString("yyyyMMdd") + "0001";
            //strRet = Util.GetDBDateTime().ToString("yyyyMMdd") + "0001";
            strRet = DateTime.Parse(txtDonateDate.Text).ToString("yyyyMMdd") + "0001";
        }

        return strRet;
    }

    private void UpdateDonor()
    {

        string strSQL = @"Select distinct cast(Donor_Id as int) as Donor_Id from DONATE_WEB where od_sob in ('" + strOrderIds + "') ;";
        DataTable dt = NpoDB.QueryGetTable(strSQL);
        DataRow dr;

        if (dt.Rows.Count > 0)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dr = dt.Rows[i];
                int Donor_Id = (int)dr["Donor_Id"];
                Dictionary<string, object> dict2 = new Dictionary<string, object>();
                dict2.Add("Donor_Id", Donor_Id);
                string strSQL2 = @"DECLARE @Begin_DonateDate datetime " +
               "DECLARE @Last_DonateDate datetime " +
               "DECLARE @Donate_No numeric " +
               "DECLARE @Donate_Total numeric " +
               "Select Top 1 @Begin_DonateDate=Donate_Date From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 Order By Donate_Date " +
               "Select Top 1 @Last_DonateDate=Donate_Date From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 Order By Donate_Date Desc " +
               "Select @Donate_No=Count(*) From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 " +
               "Select @Donate_Total=IsNull(Sum(Donate_Amt),0) From DONATE Where Donor_Id=@Donor_Id And Issue_Type<>'D' And Donate_Amt>0 " +
               "Update DONOR Set Begin_DonateDate=@Begin_DonateDate,Last_DonateDate=@Last_DonateDate,Donate_No=@Donate_No,Donate_Total=@Donate_Total Where Donor_Id=@Donor_Id ; ";
                NpoDB.ExecuteSQLS(strSQL2, dict2);
            }

        }


    }

    public void LoadFormData()
    {

        string strSql;
        DataTable dt;
        strSql = @"
        select W.od_sob as [訂單編號],CONVERT(VarChar,Donate_CreateDate,111) as [捐款日期],cast(account as int) as [捐款金額]
        ,(case when isnull(I.codename,'') <> '' then i.codename 
		else I.CodeID end ) as [付款方式] 
        ,W.Donor_Id as [捐款人編號], Donor_Name as [捐款人] 
        ,(Case When D.IsAbroad = 'N' then B.mValue + C.mValue + D.[Address] Else D.[Address] End) as [地址] 
        ,(case when ISNULL(D.Cellular_Phone,'') <> '' then D.Cellular_Phone 
        when ISNULL(D.Tel_Office_Loc,'') <> '' and ISNULL(D.Tel_Office_Ext,'') <> '' then '('+D.Tel_Office_Loc+')'+D.Tel_Office+'#'+D.Tel_Office_Ext 
        when ISNULL(D.Tel_Office_Loc,'') <> '' and ISNULL(D.Tel_Office_Ext,'') = '' then '('+D.Tel_Office_Loc+')'+D.Tel_Office 
        when ISNULL(D.Tel_Office_Loc,'') = '' and ISNULL(D.Tel_Office_Ext,'') = '' then D.Tel_Office 
        else '' end) as [電話] ,W.Donate_Purpose as [捐款用途] ,Case when P.IEPAY_returnOK is not null then '<font color=cornflowerblue>金流處理中</font>('+IEPAY_returnOK+')' else PendingMsg end as [狀態] ,MsgHelp
        From (select Donor_Id,Donate_DonorName,od_sob,Donate_CreateDate,sum(Donate_Amount) as Donate_Amount,min(Donate_Purpose) as Donate_Purpose
            from DONATE_WEB where [IsDelete] is null group by Donor_Id,Donate_DonorName,od_sob,Donate_CreateDate) as W
        join DONOR as D on W.Donor_Id = D.Donor_Id 
        join DONATE_IEPAY as P on W.od_sob = P.orderid and P.[Status] = '0' and paytype <> '1'
        left Join CODECITY As B on D.City = B.mCode 
        left Join CODECITY As C on D.Area = C.mCode 
        left Join Donate_IePayType AS I on P.paytype = I.CodeID
        left join Donate as N on W.od_sob = N.od_sob 
        ";

        Dictionary<string, object> dict = new Dictionary<string, object>();

        strSqlWhere = "where N.Donate_Id is null  ";

        if (txtDonateDateS.Text.Trim() != "")
        {
            strSqlWhere += " and Donate_CreateDate >= '" + txtDonateDateS.Text.Trim() + "' ";
        }
        if (txtDonateDateE.Text.Trim() != "")
        {
            strSqlWhere += " and Donate_CreateDate <= '" + txtDonateDateE.Text.Trim() + "' ";
        }

        strSql += strSqlWhere + " Order By Donate_CreateDate Desc,W.od_sob Desc ; ";
        dt = NpoDB.GetDataTableS(strSql, dict);
        lblGridList.Text = "";
        if (dt.Rows.Count == 0)
        {
            lblGridList.Text = "** 沒有符合條件的資料 **";
            lblGridList.ForeColor = System.Drawing.Color.Red;
            btnInput.Enabled = false;
        }
        else
        {

            btnInput.Enabled = true;
            lblGridList.ForeColor = System.Drawing.Color.Black;

            //int Donate_cnt = Convert.ToInt32(dt.Compute("count([捐款金額])", "true"));
            int Donate_cnt = dt.Rows.Count;
            double Donate_Amt = Convert.ToDouble(dt.Compute("Sum([捐款金額])", "true"));

            //2014/07/30 換掉原先的 Html Table
            string strBody = "<table class='table_h' width='100%'>";
            strBody += @"<TR><TH noWrap><SPAN><INPUT id=checkboxAll CHECKED type=checkbox name=orderidAll></SPAN></TH>
                            <TH noWrap><SPAN>訂單編號</SPAN></TH>
                            <TH noWrap><SPAN>線上捐款日期</SPAN></TH>
                            <TH noWrap><SPAN>捐款金額</SPAN></TH>
                            <TH noWrap><SPAN>付款方式</SPAN></TH>
                            <TH noWrap><SPAN>捐款人編號</SPAN></TH>
                            <TH noWrap><SPAN>捐款人</SPAN></TH>
                            <TH noWrap><SPAN>地址</SPAN></TH>
                            <TH noWrap><SPAN>電話</SPAN></TH>
                            <TH noWrap><SPAN>捐款用途</SPAN></TH>
                            <TH noWrap><SPAN>狀態</SPAN></TH>
                        </TR>";
            foreach (DataRow dr in dt.Rows)
            {

                //int Donate_Amt = Convert.ToInt32(dr["捐款金額"].ToString());
                //Donate_Amt_Sum += Donate_Amt;

                strBody += String.Format("<TR align='left'><TD noWrap align='center'><INPUT id=checkbox CHECKED type=checkbox value={1} name=orderid_{0} style=\"cursor: pointer; cursor: hand;\" ></TD>", dr["訂單編號"].ToString().Trim(), Convert.ToInt32(dr["捐款金額"].ToString()));
                strBody += "<TD noWrap><SPAN>" + dr["訂單編號"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["捐款日期"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap align='right'><SPAN id='Donate_Amt'>" + String.Format("{0:#,0}", Convert.ToInt32(dr["捐款金額"].ToString())) + "&nbsp;</SPAN></TD>";
                if (dr["付款方式"].ToString() == "非信用卡")
                {
                    //線上付款方式
                    DropDownList dpl = new DropDownList();
                    System.IO.StringWriter sw = new System.IO.StringWriter();
                    HtmlTextWriter htw = new HtmlTextWriter(sw);
                    dpl.ID = "ddl_" + dr["訂單編號"].ToString();
                    Util.FillDropDownList(dpl, Util.GetDataTable("Donate_IePayType", "Display", "1", "", ""), "CodeName", "CodeID", false);
                    dpl.Items.RemoveAt(0);  //移除信用卡選項
                    dpl.Items.Insert(0, new ListItem("非信用卡", "99"));
                    dpl.SelectedIndex = 0;
                    dpl.RenderControl(htw);
                    strBody += "<TD noWrap><SPAN>" + sw.ToString() + "</SPAN></TD>";
                }
                else
                {
                    strBody += "<TD noWrap><SPAN>" + dr["付款方式"].ToString() + "&nbsp;</SPAN></TD>";
                }
                strBody += "<TD noWrap><SPAN>" + dr["捐款人編號"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["捐款人"].ToString() + "</SPAN></TD>";
                strBody += "<TD width='35%'><SPAN>" + dr["地址"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["電話"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["捐款用途"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN title='" + dr["MsgHelp"].ToString() + "'>" + dr["狀態"].ToString() + "</SPAN></TD></TR>";

            }

            strBody += "</table>";
            lblQueryCnt.Text = String.Format("<p align='left'>捐款筆數：{0:#,0} 筆 / 捐款金額合計：{1:#,0} 元</p>", Donate_cnt, Donate_Amt);
            lblGridList.Text = strBody;
            Session["strSql"] = strSql;

            //lblQueryCnt.Text = String.Format("捐款筆數：{0:#,0} 筆 / 捐款金額合計：{1:#,0} 元", Donate_cnt, Donate_Amt);
            //Session["strSql"] = strSql;

            /* 2014/07/30 換掉自訂 Html Table
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.Keys.Add("訂單編號");
            //npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            npoGridView.ShowPage = false;
            npoGridView.DisableColumn.Add("訂單編號");
            //-------------------------------------------------------------------------
            NPOGridViewColumn col = new NPOGridViewColumn();
            col.ColumnType = NPOColumnType.Checkbox;
            col.ControlKeyColumn.Add("訂單編號");
            col.CaptionWithControl = true;
            col.ColumnName.Add("訂單編號");

            col.ControlText.Add("");
            col.ControlName.Add("orderid");
            col.ControlValue.Add("");
            col.ControlId.Add("checkbox");
            col.DisableValue = "";
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("訂單編號");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("訂單編號");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("線上捐款日期");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("線上捐款日期");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("捐款金額");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("捐款金額");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("付款方式");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("付款方式");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("捐款人編號");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("捐款人編號");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("捐款人");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("捐款人");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("地址");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("地址");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("電話");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("電話");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("捐款用途");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("捐款用途");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("狀態");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("狀態");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            lblGridList.Text = npoGridView.Render();
            */

        }


    }


}