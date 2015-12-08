using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class DonateMgr_Web_ePayList : BasePage
{

    string strSqlWhere = "where 1=1 ";

    protected void Page_Load(object sender, EventArgs e)
    {

        Session["ProgID"] = "Web_ePayList";
        ////有 npoGridView 時才需要
        //Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
        HFD_Uid.Value = Util.GetQueryString("Ser_No");
        
        if (!IsPostBack)
        {
            lblQueryCnt.Text = "";
            lblGridList.Text = "** 請先輸入查詢條件 **";
            lblGridList.ForeColor = System.Drawing.Color.Red;
            btnPrint.Enabled = false;
            btnToxls.Enabled = false;

            Util.FillDropDownList(ddlDonateIeType, Util.GetDataTable("Donate_IePayType", "Display", "1", "", ""), "CodeName", "CodeID", false);
            ddlDonateIeType.Items.Insert(0, new ListItem("請選擇", ""));
            ddlDonateIeType.SelectedIndex = 0;
            //20141024若接收到刪除指令，再重新回到該搜尋結果頁面
            if (Util.GetQueryString("Ser_No") != "")
            {
                DeleteData();
                if (Session["strStatus"] != "" && Session["strStatus"] != null)
                {
                    Status.SelectedValue = Session["strStatus"].ToString();
                }
                if (Session["DonateDateS"] != null)
                {
                    txtDonateDateS.Text = Session["DonateDateS"].ToString();
                }
                if (Session["DonateDateE"] != null)
                {
                    txtDonateDateE.Text = Session["DonateDateE"].ToString();
                }
                if (Session["HistoryRecord"] == "true")
                    cbxHistoryRecord.Checked = true;
                else
                    cbxHistoryRecord.Checked = false;
                
                LoadFormData();
            }
        }

    }

    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }

    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Response.Redirect("Web_ePayList_Print_Excel.aspx");
    }

    public void LoadFormData()
    {
        string strSql;
        DataTable dt;
        strSql = @"
        select W.Ser_No,W.od_sob as [訂單編號],CONVERT(VarChar,W.Donate_CreateDate,111) as [捐款日期]
        ,cast(w.Donate_Amount as int) as [捐款金額]
        ,(case when isnull(I.codename,'') <> '' then i.codename 
		else I.CodeID end ) as [付款方式] 
        ,W.Donor_Id as [捐款人編號], w.Donate_DonorName as [捐款人],D.Donor_Name 
        ,(Case When D.IsAbroad = 'N' then B.mValue + C.mValue + D.[Address] Else D.[Address] End) as [地址] 
        ,(case when ISNULL(D.Cellular_Phone,'') <> '' then D.Cellular_Phone 
        when ISNULL(D.Tel_Office_Loc,'') <> '' and ISNULL(D.Tel_Office_Ext,'') <> '' then '('+D.Tel_Office_Loc+')'+D.Tel_Office+'#'+D.Tel_Office_Ext 
        when ISNULL(D.Tel_Office_Loc,'') <> '' and ISNULL(D.Tel_Office_Ext,'') = '' then '('+D.Tel_Office_Loc+')'+D.Tel_Office 
        when ISNULL(D.Tel_Office_Loc,'') = '' and ISNULL(D.Tel_Office_Ext,'') = '' then D.Tel_Office 
        else '' end) as [電話] ,Case when Q.OrderNumber is not NULL then 'V' else '' end as [問卷] ,W.Donate_Purpose as [捐款用途] 
        ,CASE when P.[Status]='0' and N.Donate_Id is not null then '已轉捐款紀錄' when Status='0' and P.IEPAY_returnOK is not null then '<font color=cornflowerblue>金流處理中</font>('+IEPAY_returnOK+')' when Status='0' and N.Donate_Id is null then I.PendingMsg when Status IS NULL then '<font color=red>未完成刷卡流程</font>' else '授權失敗 代碼：'+ISNULL(RTRIM(P.errcode),'') end as [狀態] 
        ,CASE when P.[Status]='0' and N.Donate_Id is not null then '此筆已轉入捐款紀錄，可至捐款紀錄頁面查詢。' 
            when P.[Status]='0' and N.Donate_Id is null then I.MsgHelp when Status IS NULL then 
            '在金流繳款流程中有可能遇到問題，而未完成繳款流程，請關心這位奉獻天使的問題。' else 
            '在金流繳款流程中有可能輸入錯誤，而造成的金流交易的失敗，請關心這位奉獻天使的問題。' end as [MsgHelp]
        ,N.Donate_Id,N.Invoice_Pre + N.Invoice_No as [收據編號],W.[Donate_CreateDateTime] as [捐款時間]
        From DONATE_WEB as W
        join DONOR as D on W.Donor_Id = D.Donor_Id 
        left join DONATE_IEPAY as P on W.od_sob = P.orderid
        left Join CODECITY As B on D.City = B.mCode 
        left Join CODECITY As C on D.Area = C.mCode
        left Join Donate_IePayType AS I on P.paytype = I.CodeID
        left join Donate as N on W.od_sob = N.od_sob and W.Donate_Purpose = N.Donate_Purpose
        left join Pledge_Status as PS on PS.ErrorCode = P.errcode
        left join Donate_OnlineQuestion as Q on W.od_sob=Q.OrderNumber
        ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strStatus = Status.SelectedValue;

        if (strStatus != "")
        {

           	switch (strStatus)
            {
                case "0":
                    strSqlWhere += " and Status='0' and N.Donate_Id is not null "; //已轉捐款紀錄
                    break;
                case "1":
                    strSqlWhere += " and Status='0' and N.Donate_Id is null "; //授權成功
                    break;
                case "2":
                    strSqlWhere += " and Status <> '0' "; //授權失敗
                    break;
                case "3":
                    strSqlWhere += " and Status is null "; //未完成金流流程
                    break;
            }
            Session["strStatus"] = strStatus;
        }
        if (ddlDonateIeType.SelectedIndex != 0)
        {
            strSqlWhere += " and I.CodeID= '" + ddlDonateIeType.SelectedValue + "'";
        }
        if (txtDonateDateS.Text.Trim() != "")
        {
            strSqlWhere += " and Donate_CreateDate >= '" + txtDonateDateS.Text.Trim() + "' ";
            Session["DonateDateS"] = txtDonateDateS.Text.Trim();
        }
        if (txtDonateDateE.Text.Trim() != "")
        {
            strSqlWhere += " and Donate_CreateDate <= '" + txtDonateDateE.Text.Trim() + "' ";
            Session["DonateDateE"] = txtDonateDateE.Text.Trim();
        }
        //增加　查詢近3個月記錄
        if (cbxHistoryRecord.Checked == false)
        {
            strSqlWhere += " and CONVERT(VarChar,W.Donate_CreateDate ,111) > dateadd(mm,-3,CONVERT(VarChar,getdate(),111))";
            Session["HistoryRecord"] = "false";
        }
        else
        {
            Session["HistoryRecord"] = "true";
        }
        strSqlWhere += " and ISNULL(W.IsDelete,'') <> 'Y' ";
        //strSql += strSqlWhere + " Order By Donate_CreateDate Desc,W.od_sob Desc ; ";
        strSql += strSqlWhere + " Order By Donate_CreateDateTime Desc ; ";
        dt = NpoDB.GetDataTableS(strSql, dict);
        lblGridList.Text = "";
        if (dt.Rows.Count == 0)
        {
            lblQueryCnt.Text = "";
            lblGridList.Text = "** 沒有符合條件的資料 **";
            lblGridList.ForeColor = System.Drawing.Color.Red;
            btnPrint.Enabled = false;
            btnToxls.Enabled = false;
        }
        else
        {
            btnPrint.Enabled = true;
            btnToxls.Enabled = true;
            lblGridList.ForeColor = System.Drawing.Color.Black;

            //int Donate_cnt = Convert.ToInt32(dt.Compute("count([捐款金額])", "true"));
            int Donate_cnt = dt.Rows.Count;
            double Donate_Amt = Convert.ToDouble(dt.Compute("Sum([捐款金額])", "true"));

            //2014/07/30 換掉原先的 Html Table
            string strBody = "<table class='table_h' width='100%'>";
            strBody += @"<TR><TH noWrap>訂單編號</SPAN></TH>
                            <TH noWrap><SPAN>捐款日期</SPAN></TH>
                            <TH noWrap><SPAN>捐款金額</SPAN></TH>
                            <TH noWrap><SPAN>付款方式</SPAN></TH>
                            <TH noWrap><SPAN>收據編號</SPAN></TH>
                            <TH noWrap><SPAN>捐款人編號</SPAN></TH>
                            <TH noWrap><SPAN>捐款人</SPAN></TH>
                            <TH noWrap><SPAN>地址</SPAN></TH>
                            <TH noWrap><SPAN>電話</SPAN></TH>
                            <TH noWrap><SPAN>問卷</SPAN></TH>
                            <TH noWrap><SPAN>捐款用途</SPAN></TH>
                            <TH noWrap><SPAN>狀態</SPAN></TH>";
                if (Authrity.RightCheck("_Delete"))
                {
                    strBody += "<TH noWrap><SPAN>操作</SPAN></TH>";
                }
                strBody += "</TR>";
            foreach (DataRow dr in dt.Rows)
            {

                strBody += "<TR style=\"cursor: pointer; cursor: hand;\" align='left' title='點下可進入捐款人繳款紀錄清單' onclick =\"window.open(\'" + Util.RedirectByTime("DonateDataList.aspx", "Donor_Id=" + dr["捐款人編號"].ToString()) + "\',\'_self\',\'\')\">";
                strBody += "<TD noWrap><SPAN>" + dr["訂單編號"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["捐款日期"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap align='right'><SPAN>" + String.Format("{0:#,0}", Convert.ToInt32(dr["捐款金額"].ToString())) + "&nbsp;</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["付款方式"].ToString() + "&nbsp;</SPAN></TD>";
                if (dr["Donate_Id"].ToString() != "")
                {
                    strBody += "<TD noWrap><SPAN title='點下可進入繳款紀錄' onclick =\"window.event.cancelBubble=true;window.open(\'" + Util.RedirectByTime("Donate_Detail.aspx", "Donate_Id=" + dr["Donate_Id"].ToString()) + "\',\'_self\',\'\')\"><font color='blue'>" + dr["收據編號"].ToString() + "</font></SPAN></TD>";
                }
                else
                {
                    strBody += "<TD noWrap><SPAN>" + dr["收據編號"].ToString() + "&nbsp;</SPAN></TD>";
                }
                //strBody += "<TD noWrap><SPAN>" + dr["收據編號"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["捐款人編號"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["捐款人"].ToString() + "</SPAN></TD>";
                strBody += "<TD width='30%'><SPAN>" + dr["地址"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["電話"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD align='center'><SPAN>" + dr["問卷"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["捐款用途"].ToString() + "</SPAN></TD>";
                if (dr["狀態"].ToString().IndexOf("未完成刷卡流程") > 0)
                {

                    strBody += "<TD noWrap><SPAN title='" + dr["MsgHelp"].ToString() + "目前奉獻天使至金流網站操作已有" + DateDiff2(Convert.ToDateTime(dr["捐款時間"]), DateTime.Now) + "的時間。" + "'>" + dr["狀態"].ToString() + "(" + DateDiff(Convert.ToDateTime(dr["捐款時間"]), DateTime.Now) + "前)</SPAN></TD>";
                }
                else
                {
                    strBody += "<TD noWrap><SPAN title='" + dr["MsgHelp"].ToString() + "'>" + dr["狀態"].ToString() + "</SPAN></TD>";
                }

                if (Authrity.RightCheck("_Delete"))
                {
                    strBody += "<TD noWrap><a id=\"lbn_" + dr["Ser_No"].ToString() + "\" onclick=\"window.event.cancelBubble=true;if(confirm('是否確定要刪除 ?')){window.open(\'Web_ePayList.aspx?Ser_No=" + dr["Ser_No"].ToString() + "\',\'_self\',\'\');}\"><font color='blue'>刪除</font></a></TD>";
                }
                strBody += "</TR>";
            }

            strBody += "</table>";
            lblQueryCnt.Text = String.Format("<p align='left'>捐款筆數：{0:#,0} 筆 / 捐款金額合計：{1:#,0} 元</p>", Donate_cnt, Donate_Amt);
            lblGridList.Text = strBody;
            Session["strSql"] = strSql;

            /* 2014/07/30 換掉自訂 Html Table
            //Grid initial
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dt;
            npoGridView.Keys.Add("訂單編號");
            //npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            npoGridView.ShowPage = false;
            npoGridView.DisableColumn.Add("訂單編號");
            //npoGridView.DisableColumn.Add("訂單編號");
            //-------------------------------------------------------------------------
            NPOGridViewColumn col = new NPOGridViewColumn();
            col = new NPOGridViewColumn("訂單編號");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("訂單編號");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("捐款日期");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("捐款日期");
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
            //lblGridList.Text = npoGridView.Render();
            */

        }


    }
    //20141024 若刪除，更新刪除標記
    public void DeleteData()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update Donate_Web set ";

        strSql += "  IsDelete = @IsDelete";
        strSql += " where Ser_No = @Ser_No";
        dict.Add("IsDelete", "Y");
        dict.Add("Ser_No", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
        AjaxShowMessage("已刪除成功！");
    }

    private string DateDiff(DateTime DateTime1, DateTime DateTime2)
    {
        string dateDiff = null;
        TimeSpan ts1 = new TimeSpan(DateTime1.Ticks);
        TimeSpan ts2 = new TimeSpan(DateTime2.Ticks);
        TimeSpan ts = ts1.Subtract(ts2).Duration();
        dateDiff = (ts.Days > 0 ? ts.Days.ToString() + "天" : (ts.Hours > 0 ? ts.Hours.ToString() + "小時" : (ts.Minutes > 0 ? ts.Minutes.ToString() + "分鐘" : "")));
        return dateDiff;
    }

    private string DateDiff2(DateTime DateTime1, DateTime DateTime2)
    {
        string dateDiff = null;
        TimeSpan ts1 = new TimeSpan(DateTime1.Ticks);
        TimeSpan ts2 = new TimeSpan(DateTime2.Ticks);
        TimeSpan ts = ts1.Subtract(ts2).Duration();
        dateDiff = (ts.Days > 0 ? ts.Days.ToString() + "天" : "") + (ts.Hours > 0 ? ts.Hours.ToString() + "小時" : "") + (ts.Minutes > 0 ? ts.Minutes.ToString() + "分鐘" : "");
        return dateDiff;
    }

}