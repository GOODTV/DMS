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

public partial class DonateMgr_Pledge_Bot_Transfer : BasePage
{

    string TableWidth = "1450px";

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {
            //權控處理
            AuthrityControl();

            Session["ProgID"] = "Pledge_Bot_Transfer";
            Session["Transfer_AspxName"] = "Transfer_bot.aspx";
            //HFD_Dept_Id.Value = SessionInfo.DeptID;
            //筆數和金額為0
            lblAccount_Amount.Text = "0";
            lblAccount_Count.Text = "0";
            lblCard_Amount.Text = "0";
            lblCard_Count.Text = "0";
            lblError_Count.Text = "0";
            lblError_Amount.Text = "0";

            //Transfer_Date();
            btnImport_CheckFalse.Enabled = false;
        }

    }

    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_AddNew", btnImport);
        Authrity.CheckButtonRight("_Query", btnImport_CheckFalse);
    }

    private Tuple<string, string> Account()
    {
        string strSql = @"Select Count(*) as Donate_No,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY, Sum(Donate_Amt)),1),'.00','') as Donate_Amt 
                          From PLEDGE P Join DONOR D On P.Donor_Id=D.Donor_Id 
                          Where Status='授權中' And Post_SavingsNo<>'' And Post_AccountNo<>'' And Donate_Payment like '%帳戶轉帳授權書%'";
        strSql += SqlCondition();
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Donate_No = dr["Donate_No"].ToString();
        string Donate_Amt = dr["Donate_Amt"].ToString() == "" ? "0" : dr["Donate_Amt"].ToString();
        Tuple<string, string> Count = new Tuple<string, string>(Donate_No, Donate_Amt);
        return Count;
    }

    private Tuple<string, string> Card()
    {
        string Year = "14"; // tbxDonateDate.Text == "" ? "0" : (Convert.ToDateTime(tbxDonateDate.Text).Year).ToString().Substring(2, 2);
        string Month = "10"; // tbxDonateDate.Text == "" ? "0" : (Convert.ToDateTime(tbxDonateDate.Text).Month).ToString();
        string strSql = @"Select Count(*) as Donate_No, REPLACE(CONVERT(VARCHAR,CONVERT(MONEY, Sum(Donate_Amt)),1),'.00','') as Donate_Amt 
                          From PLEDGE P Join DONOR D On P.Donor_Id=D.Donor_Id 
                          Where Status='授權中' And Account_No<>'' And Donate_Payment like '%信用卡授權書%' And (Convert(int,RIGHT(Valid_Date,2))>" + Year +
                                " Or (Convert(int,RIGHT(Valid_Date,2))=" + Year + " And Convert(int,Left(Valid_Date,2))>=0))";
        strSql += SqlCondition();
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Donate_No = dr["Donate_No"].ToString();
        string Donate_Amt = dr["Donate_Amt"].ToString() == "" ? "0" : dr["Donate_Amt"].ToString();
        Tuple<string, string> Count = new Tuple<string, string>(Donate_No, Donate_Amt);
        return Count;
    }

    private Tuple<string, string> ACH()
    {
        string strSql = @"Select Count(*) as Donate_No, REPLACE(CONVERT(VARCHAR,CONVERT(MONEY, Sum(Donate_Amt)),1),'.00','') as Donate_Amt 
                          From PLEDGE P Join DONOR D On P.Donor_Id=D.Donor_Id 
                          Where Status='授權中' And Donate_Payment like '%ACH轉帳授權書%' And P_BANK<>'' And P_RCLNO<>'' And P_PID<>''";
        strSql += SqlCondition();
        DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Donate_No = dr["Donate_No"].ToString();
        string Donate_Amt = dr["Donate_Amt"].ToString() == "" ? "0" : dr["Donate_Amt"].ToString();
        Tuple<string, string> Count = new Tuple<string, string>(Donate_No, Donate_Amt);
        return Count;
    }

    /*
    public void Transfer_Date()
    {

        string strSql = "Select * From DEPT Where DeptId='" + SessionInfo.DeptID + "'";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];
        string Transfer_Date = dr["Transfer_Date"].ToString();
        DateTime LastTransfer_Date = Convert.ToDateTime(dr["LastTransfer_Date"].ToString());

    }
    */

    public string Sql()
    {
        string strSql = "";
        //20140513 修改SQL 有可能沒有信用卡的有效月年 若沒有則帶入2099/01/01以免發生Bug
        strSql = @"Select '0' as 'disabled',Pledge_Id,Pledge_Id as '授權編號',D.Donor_Name as 捐款人,Donate_Payment as 授權方式,
                          REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','')  as 扣款金額,
                          CONVERT(VarChar,Donate_FromDate,111) as 授權起日, CONVERT(VarChar,Donate_ToDate,111) as 授權迄日,Donate_Period as 轉帳週期,CONVERT(VarChar,Next_DonateDate,111) as 下次扣款日,
                          (Case When Valid_Date<>'' Then Left(Valid_Date,2)+'/'+Right(Valid_Date,2) Else '' End) as 有效月年, CONVERT(VarChar,Transfer_Date,111)  as 最後扣款日, (Case When Valid_Date<>'' Then '20' + Right(Valid_Date,2)+'/'+Left(Valid_Date,2)+'/01' Else '2099/01/01' End) as 有效年月日期
                   From PLEDGE P Join DONOR D On P.Donor_Id=D.Donor_Id 
                   Where Status='授權中' And ((Post_SavingsNo<>'' And Post_AccountNo<>'') Or Account_No<>'' Or P_RCLNO<>'') ";
        return strSql;
    }

    public string SqlCondition()
    {

        string strSql = "";
        /*
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " And D.dept_id='" + ddlDept.SelectedValue + "' ";
        }
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " And (D.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%' Or D.NickName like '%" + tbxDonor_Name.Text.Trim() + "%') ";
            strSql += " And (D.Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%' Or D.NickName like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%') ";
        }
        if (tbxDonateDate.Text != "")
        {
            strSql += " And Donate_FromDate <= '" + tbxDonateDate.Text + "' And Donate_ToDate>= '" + tbxDonateDate.Text +"' And ((Year(Next_DonateDate)<'" + Convert.ToDateTime(tbxDonateDate.Text).Year + "') Or (Year(Next_DonateDate)='" + Convert.ToDateTime(tbxDonateDate.Text).Year + "' And Month(Next_DonateDate)<='" + Convert.ToDateTime(tbxDonateDate.Text).Month + "')) ";
        }
        if (ddlDonate_Payment.SelectedIndex != 0)
        {
            strSql += " And Donate_Payment = '" + ddlDonate_Payment.SelectedItem.Text + "' ";
        }
        if (ddlDonate_Period.SelectedIndex != 0)
        {
            strSql += " And Donate_Period = '" + ddlDonate_Period.SelectedItem.Text + "' ";
        }
        */
        return strSql;
    }


    protected void Import_Check(string Donate_Date_yyyyMMDD)
    {

        string Donate_Date = Mid(Donate_Date_yyyyMMDD, 1, 4) + "-" + Mid(Donate_Date_yyyyMMDD, 5, 2) + "-" + Mid(Donate_Date_yyyyMMDD, 7, 2);
        lblGridList.Text = "";//先清空再填入
        var flag = false;
        //撈出Temp Table:PLEDGE_SEND_RETURN裡的Pledge_Id串成陣列 Pledge_Ids
        string strSql = "Select Pledge_Id From Pledge_Send_Return Where Return_Status='Y'";
        /*
        SELECT a.[Pledge_Id]
              ,a.[Donate_Amt]
	          ,b.[Donate_Amt]
	          ,CASE WHEN b.[Donate_Amt] is not null THEN 1 ELSE 0 END as [金額比對]
	          ,CASE WHEN c.Pledge_Id is not null THEN 1 ELSE 0 END as [收據比對]
          FROM [DonationGT].[dbo].[Pledge_Send_Return] as a
          LEFT JOIN [DonationGT].[dbo].[Pledge] as b
          on a.Pledge_Id = b.Pledge_Id
          and a.[Donate_Amt] = b.[Donate_Amt]
          LEFT JOIN [DonationGT].[dbo].[Donate] as c
          on a.Pledge_Id = c.Pledge_Id
          and CAST('2012-12-25' as datetime) = CAST(c.[Donate_Date] as datetime)
          Where Return_Status='Y'
         * */

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donate_Date", Donate_Date);

        DataTable dt;
        DataRow dr;

        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count != 0)
        {
            string[] Pledge_Ids = new string[dt.Rows.Count];
            bool first = true;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dr = dt.Rows[i];
                Pledge_Ids[i] = dr["Pledge_Id"].ToString();
            }

            //判斷是否PLEDGE_SEND_RETURN裡的Pledge_Id有在PLEDGE中
            reload_list(true, first, Pledge_Ids);//有
            if (lblGridList.Text != "")
            {
                flag = true;
            }

            //增加授權失敗的清單
            Pledge_Err_List();

            if (flag == false)
            {
                Page.Controls.Add(new LiteralControl("<script>alert(\"授權清單與回覆檔案資料不符！\");</script>"));
            }
            else
            {
                //SendEmail_Excel();
                //SendEmail_Excel_OK();
                strSql = "Select Count(*) AS Count,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,SUM(Donate_Amt)),1),'.00','') AS Sum_Amt From dbo.PLEDGE_SEND_RETURN Where Return_Status='Y'";
                DataTable dt2 = NpoDB.GetDataTableS(strSql, null);
                DataRow dr2 = dt2.Rows[0];
                Page.Controls.Add(new LiteralControl("<script>alert(\"回覆檔資料匯入成功，共" + dr2["Count"] + "筆，金額" + dr2["Sum_Amt"] + "元。\\n授權成功資料已自動勾選！\");</script>"));
            }
        }

    }

    private void Pledge_Err_List()
    {

        string strSql = "";
        strSql = @"
            Select ROW_NUMBER() OVER(ORDER BY P.Pledge_Id) AS [序號],p.Pledge_Id as '授權編號',D.Donor_Name as 捐款人
            ,Donate_Payment as 授權方式,cast(p.Donate_Amt as int)  as 扣款金額
            ,CONVERT(VarChar,Donate_FromDate,111) as 授權起日, CONVERT(VarChar,Donate_ToDate,111) as 授權迄日
            ,Donate_Period as 轉帳週期,Authorize as [末三碼CVV]
            ,(Case When Valid_Date<>'' Then Left(Valid_Date,2)+'/'+Right(Valid_Date,2) Else '' End) as 有效月年
            ,r.[Return_Status_No] as [授權失敗碼],s.Note as [授權失敗原因],p.[Status]
            From PLEDGE P 
            Join DONOR D On P.Donor_Id=D.Donor_Id 
            join Pledge_Send_Return r on p.Pledge_Id=r.Pledge_Id and r.[Return_Status] = 'N'
            left join Pledge_Status S on R.Return_Status_No = S.ErrorCode
            order by p.Pledge_Id
         ;   ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count > 0)
        {

            //lblGridList.ForeColor = System.Drawing.Color.Black;
            string strBody = "<table class='table_h' width='100%'>";
            strBody += @"<TR><TH noWrap><SPAN>序號</SPAN></TH>
                            <TH noWrap><SPAN>授權編號</SPAN></TH>
                            <TH noWrap><SPAN>捐款人</SPAN></TH>
                            <TH noWrap><SPAN>授權方式</SPAN></TH>
                            <TH noWrap><SPAN>扣款金額</SPAN></TH>
                            <TH noWrap><SPAN>授權起日</SPAN></TH>
                            <TH noWrap><SPAN>授權迄日</SPAN></TH>
                            <TH noWrap><SPAN>轉帳週期</SPAN></TH>
                            <TH noWrap><SPAN>有效月年</SPAN></TH>
                            <TH noWrap><SPAN>末三碼CVV</SPAN></TH>
                            <TH noWrap><SPAN>失敗碼</SPAN></TH>
                            <TH noWrap><SPAN>授權失敗原因</SPAN></TH>
                        </TR>";
            foreach (DataRow dr in dt.Rows)
            {

                strBody += "<TR " + (Convert.ToInt64(dr["扣款金額"].ToString()) >= 10000 ? "style=\"COLOR: white; BACKGROUND-COLOR: darkorange\" " : "") + "align='center'><TD noWrap align='right'><SPAN>" + dr["序號"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["授權編號"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["捐款人"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["授權方式"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD noWrap align='right'><SPAN>" + String.Format("{0:#,0}", Convert.ToInt64(dr["扣款金額"].ToString())) + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["授權起日"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["授權迄日"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["轉帳週期"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["有效月年"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["末三碼CVV"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["授權失敗碼"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap align='left'><SPAN" + (dr["授權失敗碼"].ToString() == "05" && Convert.ToInt64(dr["扣款金額"].ToString()) < 10000 ? " style=\"COLOR: red\"" : "") + ">" + dr["授權失敗原因"].ToString() + "</SPAN></TD></TR>";

            }

            strBody += "</table>";
            lblGridList.Text += "<BR/><TABLE cellSpacing='0' width='100%' border=0 height='30px'><TR style='BACKGROUND-COLOR: mistyrose'>" +
            "<TD width='5%' align=right><DIV style='HEIGHT: 15px; WIDTH: 20px; BACKGROUND-COLOR: darkorange'></DIV></TD><TD width='25%' align=left><SPAN style='FONT-SIZE: 20px; FONT-WEIGHT: bold; COLOR: red;'>大額警示</SPAN></TD>" +
            "<TD align=center><SPAN style='FONT-SIZE: 18px; COLOR: blue;'>本次匯入授權失敗</SPAN></TD><TD width='25%'></TD><TD width='5%'></TD></TR></TABLE>" + strBody;
            //Session["strSql"] = strSql;
        }
    }

    protected void btnImport_CheckFalse_Click(object sender, EventArgs e)
    {

        Response.Redirect("PledgeReturnError_Print_Excel.aspx");

    }

    private void reload_list(bool check, bool first, string[] pledge_id)
    {

        string strSql = Sql();
        strSql += SqlCondition();
        if (check == true)
        {
            strSql += " AND Pledge_Id IN ( ";
        }
        else
        {
            strSql += " AND Pledge_Id NOT IN ( ";
        }
        for (int i = 0; i < pledge_id.Length; i++)
        {
            if (i == pledge_id.Length - 1)
            {
                strSql += "'" + pledge_id[i] + "'";
            }
            else
            {
                strSql += "'" + pledge_id[i] + "',";
            }
        }
        strSql += " ) ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (check == true && dt.Rows.Count == 0)
        {
            lblGridList.Text = "";
            return;
        }
        else if (check == false && dt.Rows.Count == 0)
        {
            return;
        }

        if (dt.Rows.Count > 0)
        {

            //lblGridList.ForeColor = System.Drawing.Color.Black;
            string strBody = "<table class='table_h' width='100%'>";
            strBody += @"<TR><TH noWrap><SPAN>序號</SPAN></TH>
                            <TH noWrap><SPAN>授權編號</SPAN></TH>
                            <TH noWrap><SPAN>捐款人</SPAN></TH>
                            <TH noWrap><SPAN>授權方式</SPAN></TH>
                            <TH noWrap><SPAN>扣款金額</SPAN></TH>
                            <TH noWrap><SPAN>授權起日</SPAN></TH>
                            <TH noWrap><SPAN>授權迄日</SPAN></TH>
                            <TH noWrap><SPAN>轉帳週期</SPAN></TH>
                            <TH noWrap><SPAN>下次扣款日</SPAN></TH>
                            <TH noWrap><SPAN>有效月年</SPAN></TH>
                            <TH noWrap><SPAN>最後扣款日</SPAN></TH>
                        </TR>";
            foreach (DataRow dr in dt.Rows)
            {

                strBody += "<TR " + (Convert.ToInt64(dr["扣款金額"].ToString()) >= 10000 ? "style=\"COLOR: white; BACKGROUND-COLOR: darkorange\" " : "") + "align='center'><TD noWrap align='right'><SPAN>" + dr["序號"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["授權編號"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["捐款人"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["授權方式"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD noWrap align='right'><SPAN>" + String.Format("{0:#,0}", Convert.ToInt64(dr["扣款金額"].ToString())) + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["授權起日"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["授權迄日"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["轉帳週期"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["下次扣款日"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["有效月年"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["最後扣款日"].ToString() + "</SPAN></TD></TR>";

            }

            strBody += "</table>";
            lblGridList.Text += (check == true ? "<BR/><DIV style=\"FONT-SIZE: 18px; COLOR: blue; LINE-HEIGHT: 30px; BACKGROUND-COLOR: skyblue\">本次匯入授權成功</DIV>" : "<BR/><DIV style=\"FONT-SIZE: 18px; COLOR: blue; LINE-HEIGHT: 30px; BACKGROUND-COLOR: palegreen\">扣款日期前累計等待授權紀錄</DIV>") + strBody;
            //lblGridList.Text += "<BR/><TABLE cellSpacing='0' width='100%' border=0 height='30px'><TR style='BACKGROUND-COLOR: mistyrose'>" +
            //"<TD width='5%' align=right><DIV style='HEIGHT: 15px; WIDTH: 20px; BACKGROUND-COLOR: darkorange'></DIV></TD><TD width='25%' align=left><SPAN style='FONT-SIZE: 20px; FONT-WEIGHT: bold; COLOR: red;'>大額警示</SPAN></TD>" +
            //"<TD align=center><SPAN style='FONT-SIZE: 18px; COLOR: blue;'>本次匯入授權失敗</SPAN></TD><TD width='25%'></TD><TD width='5%'></TD></TR></TABLE>" + strBody;
            //Session["strSql"] = strSql;
            Session["List"] = lblGridList.Text;
        }


    }

    private void SendEmail_Excel()
    {

        string strSql = @"select ROW_NUMBER() OVER(ORDER BY P.Pledge_Id) AS '序號'
                        ,P.Pledge_Id as '授權編號',CAST(SR.Donate_Amt as int) as '授權金額'
                        ,D.Donor_Id as '捐款人編號', D.Donor_Name as '捐款人姓名'
						,p.[Card_Bank] as [發卡銀行],p.[Account_No] as [信用卡卡號],p.[Valid_Date] as [效期],Authorize as [末三碼CVV]
                        ,IsNull(D.ZipCode,'') AS '郵遞區號'
                        ,Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') 
		                        else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
                        ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
			                        (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
			                        Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
		                        Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
			                        Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話'  
                        ,IsNull(Cellular_Phone,'') AS '手機'
                        ,PS.ErrorCode as '授權失敗碼' ,PS.Note as '授權失敗原因'
                        from Pledge_Send_Return SR
                        left join Pledge P on SR.Pledge_Id = P.Pledge_Id
                        left join Donor D on P.Donor_Id = D.Donor_Id
                        LEFT JOIN dbo.CODECITYNew C1 ON D.City = C1.ZipCode
                        LEFT JOIN dbo.CODECITYNew C2 ON D.Area = C2.ZipCode
                        left join Pledge_Status PS on SR.Return_Status_No = PS.ErrorCode
                        where SR.Return_Status='N'
                        order by SR.Pledge_Id";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            //ShowSysMsg("查無資料!!!");
            return;
        }

        //string style = "<style> .text { mso-number-format:\\@; } </style> ";
        string style = "<style type=text/css>td{mso-number-format:\"\\@\";}</style>";
        GridList.Text = style + GetTitle(dt.Rows[0], false) + GetTable1(dt.Rows[0], strSql, false);
        long longSum = Convert.ToInt64(dt.Compute("Sum([授權金額])", "true"));
        byte[] byteArray = Encoding.UTF8.GetBytes(GridList.Text);
        MemoryStream sm = new MemoryStream(byteArray);

        string StrEmailToDonations = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];
        //發送內部通知mail*****************************************
        string strBody = "親愛的同工平安,<BR/><BR/>";
        //"2014-01-01"
        strBody += "以下是本次扣款日期: " + "2014/08/25" + " 回覆之授權失敗記錄 <BR/><BR/>";
        strBody += "<font color='blue'>合計授權失敗的筆數為：" + String.Format("{0:#,0}", dt.Rows.Count) + " 筆，金額為: " + String.Format("{0:#,0}", longSum) + " 元 </font><BR/><BR/>";
        strBody += GridList.Text;
        strBody += "<BR/>請盡速追蹤處理!<BR/><BR/>";

        SendEMailObject MailObject = new SendEMailObject();
        string MailSubject = "信用卡授權失敗提示(批次) - 扣款日期:" + "2014/08/25";
        string MailBody = strBody;
        string strFilename = "PledgeReturnError.xls";
        string result = MailObject.SendEmailAttachment(StrEmailToDonations, strFilename, MailSubject, strBody, sm);

        if (MailObject.ErrorCode != 0)
        {
            this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        }

    }

    private string GetTitle(DataRow dr, bool OKERR)
    {
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        table.Border = 0;
        table.CellPadding = 0;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "15px");
        css.Add("font-family", "標楷體");
        css.Add("width", TableWidth);

        string strFontSize = "24px";

        string strTitle = OKERR ? "台銀授權成功回覆檔暨捐款人資料<br/> " : "台銀授權失敗回覆檔暨捐款人資料<br/> ";
        row = new HtmlTableRow();

        cell = new HtmlTableCell();
        cell.InnerHtml = strTitle;
        css = cell.Style;
        css.Add("text-align", "center");
        css.Add("font-size", strFontSize);
        cell.ColSpan = OKERR ? 12 : 17;
        row.Cells.Add(cell);

        table.Rows.Add(row);

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }

    private string GetTable1(DataRow dr, string strSql, bool OKERR)
    {
        //組 table
        HtmlTable table = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        Dictionary<string, object> dict = new Dictionary<string, object>();

        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        int count = dt.Rows.Count;

        DataTable dtRet = OKERR ? CaseUtil.PledgeReturnOK_Print(dt) : CaseUtil.PledgeReturnError_Print(dt);

        table.Border = 1;
        table.CellPadding = 3;
        table.CellSpacing = 0;
        css = table.Style;
        css.Add("font-size", "15px");
        css.Add("font-family", "新細明體");
        css.Add("width", TableWidth);
        css.Add("line-height", "25px");

        row = new HtmlTableRow();

        int iCtrl = 0;
        foreach (DataColumn dc in dtRet.Columns)
        {
            cell = new HtmlTableCell();
            cell.Style.Add("background-color", " #FFE4C4 ");
            /*
            cell.Style.Add("border-left", ".5pt solid windowtext");
            cell.Style.Add("border-top", ".5pt solid windowtext");
            cell.Style.Add("border-right", ".5pt solid windowtext");
            cell.Style.Add("border-bottom", ".5pt solid windowtext");
            */

            if (iCtrl == 0)
            {
                cell.Width = "30";
            }
            else
            {
                cell.Width = "90mm";
            }
            cell.InnerHtml = dc.ColumnName == "" ? "&nbsp;" : dc.ColumnName;
            row.Cells.Add(cell);
            iCtrl++;
        }
        table.Rows.Add(row);

        foreach (DataRow drow in dtRet.Rows)
        {

            row = new HtmlTableRow();
            iCtrl = 0;
            foreach (object objItem in drow.ItemArray)
            {
                cell = new HtmlTableCell();
                /*
                cell.Style.Add("border-left", ".5pt solid windowtext");
                cell.Style.Add("border-top", ".5pt solid windowtext");
                cell.Style.Add("border-right", ".5pt solid windowtext");
                cell.Style.Add("border-bottom", ".5pt solid windowtext");
                */

                if (iCtrl == 2)
                {
                    cell.InnerHtml = objItem.ToString() == "" ? "&nbsp;" : String.Format("{0:#,0}", Convert.ToInt64(objItem.ToString()));
                }
                else
                {
                    cell.InnerHtml = objItem.ToString() == "" ? "&nbsp;" : objItem.ToString();
                }
                //cell.InnerHtml = objItem.ToString() == "" ? "&nbsp;" : objItem.ToString();
                row.Cells.Add(cell);
                iCtrl++;
            }
            table.Rows.Add(row);
        }

        //轉成 html 碼
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }

    private void SendEmail_Excel_OK()
    {

        string strSql = @"select ROW_NUMBER() OVER(ORDER BY P.Pledge_Id) AS '序號'
                        ,P.Pledge_Id as '授權編號',CAST(SR.Donate_Amt as int) as '授權金額'
                        ,D.Donor_Id as '捐款人編號', D.Donor_Name as '捐款人姓名'
						,p.[Card_Bank] as [發卡銀行],p.[Account_No] as [信用卡卡號],p.[Valid_Date] as [效期]
                        ,IsNull(D.ZipCode,'') AS '郵遞區號'  
                        ,Case When IsAbroad = 'Y' Then IsNull(OverseasCountry,'') + ' ' + IsNull(OverseasAddress,'') 
		                        else IsNull(C1.Name,'') + IsNull(C2.Name,'') + Address END '地址'  
                        ,Case When Tel_Office_Loc Is Null Or Tel_Office_Loc = '' Then  
			                        (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then Tel_Office  
			                        Else Tel_Office + ' Ext.' + Tel_Office_Ext End)  
		                        Else (Case When Tel_Office_Ext Is Null Or Tel_Office_Ext ='' Then '(' + Tel_Office_Loc + ')' + Tel_Office  
			                        Else '(' + Tel_Office_Loc + ')' + Tel_Office + ' Ext.' + Tel_Office_Ext End) End '電話'  
                        ,IsNull(Cellular_Phone,'') AS '手機'
                        from Pledge_Send_Return SR
                        left join Pledge P on SR.Pledge_Id = P.Pledge_Id
                        left join Donor D on P.Donor_Id = D.Donor_Id
                        LEFT JOIN dbo.CODECITYNew C1 ON D.City = C1.ZipCode
                        LEFT JOIN dbo.CODECITYNew C2 ON D.Area = C2.ZipCode
                        left join Pledge_Status PS on SR.Return_Status_No = PS.ErrorCode
                        where SR.Return_Status='Y'
                        order by SR.Pledge_Id";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            //ShowSysMsg("查無資料!!!");
            return;
        }
       
        //string style = "<style> .text { mso-number-format:\\@; } </style> ";
        string style = "<style type=text/css>td{mso-number-format:\"\\@\";}</style>";
        string strReplaceOldValue = GetTable1(dt.Rows[0], strSql, true);
        long longSum = Convert.ToInt64(dt.Compute("Sum([授權金額])", "true"));
        string strReplaceNewValue = "<tr><td colspan='2'>合計</td><td>" + String.Format("{0:#,0}", longSum) + "</td><td colspan='9'>&nbsp;</td></tr></table>";
        GridList.Text = style + GetTitle(dt.Rows[0], true) + strReplaceOldValue.Replace("</table>", strReplaceNewValue);
        byte[] byteArray = Encoding.UTF8.GetBytes(GridList.Text);
        MemoryStream sm = new MemoryStream(byteArray);

        string StrEmailToDonations = System.Configuration.ConfigurationManager.AppSettings["EmailToDonations"];
        //發送內部通知mail*****************************************
        string strBody = "親愛的同工平安,<BR/><BR/>";
        //"2014-01-01"
        strBody += "以下是本次扣款日期: " + "2014/08/58" + " 回覆之授權成功記錄 <BR/>";
        strBody += "<font color='blue'>合計授權成功的筆數為：" + String.Format("{0:#,0}", dt.Rows.Count) + " 筆，金額為： " + String.Format("{0:#,0}", longSum) + " 元 </font><BR/><BR/>";
        strBody += GridList.Text;
        strBody += "<BR/>請盡速進行轉捐款收據作業!<BR/><BR/>";

        SendEMailObject MailObject = new SendEMailObject();
        string MailSubject = "信用卡授權成功通知(批次) - 扣款日期:" + "2014/08/58";
        string MailBody = strBody;
        string strFilename = "PledgeReturnOK.xls";
        string result = MailObject.SendEmailAttachment(StrEmailToDonations, strFilename, MailSubject, strBody, sm);

        if (MailObject.ErrorCode != 0)
        {
            this.Page.RegisterStartupScript("s", "<script>alert('" + MailObject.ErrorMessage + MailObject.ErrorCode + "=>" + result + "');</script>");
        }

    }

    protected void btnFileImport_Click(object sender, EventArgs e)
    {

        if (!FileUpload.HasFile)
        {
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript'>alert('請選擇檔案');</script>");
            return;
        }
        if (FileNameValid() == false)
            return;
        if (!AllowSave())
            return;

        string txt_filePath = "";

        txt_filePath = SaveFileAndReturnPath();//先上傳TXT檔案給Server
        SaveOrInsertDB();//讀取TXT中的資料寫入DB

    }

    public bool FileNameValid()
    {
        string strfilePath = FileUpload.FileName;
        string strfileName = strfilePath.Substring(strfilePath.LastIndexOf(@"\") + 1);
        if (strfileName.Contains(" ") || strfileName.Contains("　"))
        {
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('檔案名稱不能有空白');</script>");
            return false;
        }
        return true;
    }

    //檢查及限定上傳檔案類型與大小
    public bool AllowSave()
    {
        int DenyMbSize = 4;//限制大小30MB
        bool flag = false;
        //判斷副檔名
        switch (System.IO.Path.GetExtension(FileUpload.PostedFile.FileName).ToUpper())
        {
            case ".TXT":
                flag = true;
                break;
            case ".01R":
                flag = true;
                break;
            case ".02R":
                flag = true;
                break;
            case ".03R":
                flag = true;
                break;
            default:
                //this.RegisterStartupScript("js", @"<script language='javascript' >alert('此檔案類型不允許上傳');</script>"); 
                this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('此檔案類型不允許上傳');</script>");
                flag = false;
                break;
        }
        if (Util.GetFileSizeMB(FileUpload.FileBytes.Length) > DenyMbSize)
        {
            //this.RegisterStartupScript("js", @"<script language='javascript' >alert('檔案大小不得超過" + DenyMbSize + @"MB');</script>");
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('檔案大小不得超過" + DenyMbSize + @"MB');</script>");
            flag = false;
        }
        return flag;
    }

    //儲存TXT檔案給Server
    private string SaveFileAndReturnPath()
    {
        string return_file_path = "";//上傳的Excel檔在Server上的位置
        if (FileUpload.FileName != "")
        {
            String appPath = Request.PhysicalApplicationPath;
            return_file_path = System.IO.Path.Combine(appPath + System.Configuration.ConfigurationManager.AppSettings["UploadPath"], Guid.NewGuid().ToString() + ".txt");

            FileUpload.SaveAs(return_file_path);
        }
        return return_file_path;
    }

    //把TXT資料Insert into Table
    private void SaveOrInsertDB()
    {

        //先將PLEDGE_SEND_RETURN中資料刪除
        string strSql = "Delete from Pledge_Send_Return where Pledge_Id>0";
        NpoDB.ExecuteSQLS(strSql, null);
        //讀取TXT檔
        bool flag = false;
        System.IO.StreamReader red = new System.IO.StreamReader(this.FileUpload.PostedFile.InputStream, System.Text.Encoding.Default);
        //作檢查用的變數
        int SumAmt = 0;
        int Total_Amt = 0;
        string Donate_Date_yyyyMMDD = "";

        while (red.Peek() > 0)
        {
            string strContent = red.ReadLine();
            //基本資料
            string SourcePath = "";
            string SourceName = "";
            string UploadName = "";
            string UploadSize = "";
            string ExtName = "";
            if (strContent.Trim() != "")
            {
                UploadName = Guid.NewGuid().ToString() + ".txt";
                SourcePath = FileUpload.PostedFile.FileName.Replace(UploadName, "");
                SourceName = FileUpload.FileName;
                UploadSize = string.Format((FileUpload.PostedFile.ContentLength / 1024).ToString(), "0.00");
                ExtName = System.IO.Path.GetExtension(FileUpload.PostedFile.FileName).ToUpper();
            }
            //Data
            string Account_No = "";
            string Donate_Amt = "";
            string Return_Status = "";
            string Pledge_Id = "";
            if (strContent.IndexOf("T", 0) == -1)
            {
                Account_No = (Mid(strContent, 19, 19).Trim());
                Donate_Amt = Trim0(Mid(strContent, 45, 10));
                Return_Status = (Mid(strContent, 57, 2).Trim());
                Pledge_Id = Trim0(Mid(strContent, 76, 8));

                SumAmt = SumAmt + Convert.ToInt32(Donate_Amt);
                //insert Data
                strSql = "insert into  Pledge_Send_Return\n";
                strSql += "( Pledge_Id, Account_No, Return_Status, Donate_Amt,Return_Status_No ) values\n";
                strSql += "( @Pledge_Id,@Account_No,@Return_Status,@Donate_Amt,@Return_Status_No ) ";
                strSql += "\n";
                strSql += "select @@IDENTITY";

                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Pledge_Id", Pledge_Id);
                dict.Add("Account_No", Account_No);
                if (Return_Status == "00")
                {
                    dict.Add("Return_Status", "Y");
                }
                else
                {
                    dict.Add("Return_Status", "N");
                }
                dict.Add("Donate_Amt", Donate_Amt);
                dict.Add("Return_Status_No", Return_Status);
                NpoDB.ExecuteSQLS(strSql, dict);
                flag = true;

                HFD_Pledge_Import.Value = "1";
            }
            else
            {
                Total_Amt = Convert.ToInt32(Trim0(Mid(strContent, 7, 10)));
                Donate_Date_yyyyMMDD = Mid(strContent, 19, 8);
            }

        }

        if (SumAmt != Total_Amt)
        {
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('台銀回覆檔格式有問題！！');</script>");
        }
        else
        {
            if (flag == true)
            {

                Import_Check(Donate_Date_yyyyMMDD);

            }
            else
            {
                ShowSysMsg("無符合的資料！");
            }
        }
    }

    // 取得起始位置至結束之字串
    public static string Mid(string sSource, int iStart, int iLength)
    {
        if (sSource.Trim().Length > 0)
        {
            int iStartPoint = iStart > sSource.Length ? sSource.Length : iStart;
            return sSource.Substring(iStartPoint, iStartPoint + iLength > sSource.Length ? sSource.Length - iStartPoint : iLength);
        }
        else
            return "";
    }

    public static string Trim0(string sSource)
    {
        int r = 0;
        for (int i = 0; i < sSource.Length; i++)
        {
            if (sSource[i].ToString().Equals("0"))
            {
                r += 1;
            }
            else
            {
                break;
            }
        }
        return sSource.Substring(r);
    }  


}
