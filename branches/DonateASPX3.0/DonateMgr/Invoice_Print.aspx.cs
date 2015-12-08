using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonateMgr_Invoice_Print : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "Invoice_Print";
        //權控處理
        AuthrityControl();
        if (!IsPostBack)
        {
            LoadDropDownListData();
            LoadCheckBoxListData();
        }
    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_Print", btnAddress);
        Authrity.CheckButtonRight("_Print", btnReport);
        Authrity.CheckButtonRight("_Print", btnPost);
    }
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);

        //經手人
        string strSql = "select Distinct Create_User from Donate ";
        DataTable dt = NpoDB.QueryGetTable(strSql);
        int count = dt.Rows.Count;
        for (int i = 0; i < count; i++)
        {
            DataRow dr = dt.Rows[i];
            ddlCreate_User.Items.Insert(i,new ListItem(dr["Create_User"].ToString(),dr["Create_User"].ToString()));
        }
        ddlCreate_User.Items.Insert(0, new ListItem("", ""));
        ddlCreate_User.SelectedIndex = 0;
        //募款活動
        Util.FillDropDownList(ddlActName, Util.GetDataTable("Act", "1", "1", "", ""), "Act_ShortName", "Act_Id", false);
        ddlActName.Items.Insert(0, new ListItem("", ""));
        ddlActName.SelectedIndex = 0;

        //名條格式
        Util.FillDropDownList(ddlFormat, Util.GetDataTable("CaseCode", "GroupName", "郵遞標籤", "", ""), "CaseName", "CaseID", false);
        ddlFormat.Items.Insert(0, new ListItem("", ""));
        ddlFormat.SelectedIndex = 8;
    }
    public void LoadCheckBoxListData()
    {
        ////收據開立
        //20140124改
        cblInvoice_Type.Items.Insert(0, new ListItem("不寄", "不寄"));
        cblInvoice_Type.Items.Insert(1, new ListItem("單次收據", "單次收據"));
        //Util.FillCheckBoxList(cblInvoice_Type, Util.GetDataTable("CaseCode", "GroupName", "收據開立", "", ""), "CaseName", "CaseID", false);
        //cblInvoice_Type.Items[0].Selected = false;
        //捐款方式
        Util.FillCheckBoxList(cblDonate_Payment, Util.GetDataTable("CaseCode", "GroupName", "捐款方式", "ABS(CaseID)", ""), "CaseName", "CaseID", false);
        cblDonate_Payment.Items[0].Selected = false;
        //捐款用途
        Util.FillCheckBoxList(cblDonate_Purpose, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseID", false);
        cblDonate_Purpose.Items[0].Selected = false;
        //其他
        cblExp_Pre.Items.Insert(0, new ListItem("手開收據", "手開收據"));
        cblExp_Pre.Items.Insert(1, new ListItem("作廢收據", "作廢收據"));
    }
    //---------------------------------------------------------------------------
    protected void btnReport_Click(object sender, EventArgs e)
    {
        if (cblInvoice_Type.Items[0].Selected == false && cblInvoice_Type.Items[1].Selected == false)
        {
            AjaxShowMessage("收據開立 至少需勾選一個！");
            return;
        }
        string strSql = Sql_Print();
        //收據重印
        if (cbxInvoice.Checked)
        {
            strSql += " and do.Invoice_Print = '1' ";
        }
        else
        {
            strSql += " and (do.Invoice_Print = '0' or do.Invoice_Print is null) ";
        }
        strSql = Sort(strSql);
        Session["strSql_Print"] = strSql;
        Sql_Edit(1);
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Report_OnClick();", true);
    }
    protected void btnAddress_Click(object sender, EventArgs e)
    {
        string strSql = Sql_Print();
        //地址名條重印
        if (cbxAddress.Checked)
        {
            strSql += " and do.Invoice_Print_Add = '1' ";
        }
        else
        {
            strSql += " and (do.Invoice_Print_Add = '0' or do.Invoice_Print_Add is null) ";
        }
        strSql = Sort(strSql);
        Session["strSql_Print"] = strSql;
        Sql_Edit(2);
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Address_OnClick(" + ddlFormat.SelectedIndex + ");", true);
    }
    protected void btnPost_Click(object sender, EventArgs e)
    {
        string strSql = Sql_Print();
        //收據重印
        if (cbxInvoice.Checked)
        {
            strSql += " and do.Invoice_Print = '1' ";
        }
        else
        {
            strSql += " and (do.Invoice_Print = '0' or do.Invoice_Print is null) ";
        }
        //地址名條重印
        if (cbxAddress.Checked)
        {
            strSql += " and do.Invoice_Print_Add = '1' ";
        }
        else
        {
            strSql += " and (do.Invoice_Print_Add = '0' or do.Invoice_Print_Add is null) ";
        }
        strSql = Sort(strSql);
        Session["strSql"] = strSql;
        Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", " Post_OnClick();", true);
    }
    private string Sql_Print()
    {
        string strSql_temp = @"select Distinct dr.Donor_Id as [編號],
                                 dr.Donor_Name as [捐款人],dr.Title as [稱謂],dr.Title2 as [稱謂2],do.Invoice_No as [收據編號],Donate_Forign as [外幣],Donate_ForignAmt as [外幣金額],do.Donate_Date,CONVERT(nvarchar(4000),Invoice_PrintComment) as Invoice_PrintComment,
                                 do.Donate_Purpose as [款項用途],dr.Cellular_Phone as [電話], REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,do.Donate_Amt),1),'.00','') as [金額],
                                 dr.IsAbroad_Invoice,dr.Invoice_Attn,dr.Attn,dr.Invoice_OverseasCountry,do.Invoice_Title,
                                 WhereClause1     
                                 WhereClause2
                                 on do.Donor_Id = dr.Donor_Id
                                 where dr.DeleteDate is null and do.Issue_Type != 'D'";
        string strSql = Condition(strSql_temp);
        return strSql;
    }
    private void Sql_Edit(int i)
    {
        string strSql_temp = @"select do.Donate_Id as [捐款編號],
                                 WhereClause1     
                                 WhereClause2
                                 on do.Donor_Id = dr.Donor_Id
                                 where dr.DeleteDate is null and do.Issue_Type != 'D' WhereClause3";
        string strSql = Condition(strSql_temp);
        strSql = Sort(strSql);
        if (i == 1)
        {
            if (cbxInvoice.Checked)
            {
                strSql = strSql.Replace("WhereClause3", " and do.Invoice_Print = '1' ");
            }
            else
            {
                strSql = strSql.Replace("WhereClause3", " and (do.Invoice_Print = '0' or do.Invoice_Print is null) ");
            }
        }
        else if (i == 2)
        {
            if (cbxAddress.Checked)
            {
                strSql = strSql.Replace("WhereClause3", " and do.Invoice_Print_Add = '1' ");
            }
            else
            {
                strSql = strSql.Replace("WhereClause3", " and (do.Invoice_Print_Add = '0' or do.Invoice_Print_Add is null) ");
            }
        }
        Session["strSql_Edit"] = strSql;
    }
    private string Condition(string strSql)
    {
        //機構
        if (ddlDept.SelectedIndex != 0)
        {
            strSql += " and do.Dept_Id = '" + ddlDept.SelectedValue + "'";
        }
        //經手人
        if (ddlCreate_User.SelectedIndex != 0)
        {
            strSql += " and do.Create_User = '" + ddlCreate_User.SelectedValue + "'";
        }
        //捐款日期
        if (tbxDonate_DateS.Text.Trim() != "")
        {
            strSql += " and do.Donate_Date >='" + tbxDonate_DateS.Text.Trim() + "'";
        }
        if (tbxDonate_DateE.Text.Trim() != "")
        {
            strSql += " and Donate_Date <='" + tbxDonate_DateE.Text.Trim() + "'";
        }
        /* 2014/5/23 修改為捐款人編號
        //捐款人
        if (tbxDonor_Name.Text.Trim() != "")
        {
            //20140515 修改by Ian_Kao 修改傳入參數時 ' 可能造成錯誤的bug
            //strSql += " and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim() + "%'";
            strSql += " and dr.Donor_Name like '%" + tbxDonor_Name.Text.Trim().Replace("'", "''") + "%'";
        }
        */
        //捐款人編號
        if (tbxDonor_Id.Text.Trim() != "")
        {
            strSql += " and dr.Donor_Id like '%" + tbxDonor_Id.Text.Trim().Replace("'", "''") + "%'";
        }
        //募款活動
        if (ddlActName.SelectedIndex != 0)
        {
            strSql += " and do.Act_Id ='" + ddlActName.SelectedValue + "' ";
        }
        //收據編號
        if (tbxInvoice_NoS.Text.Trim() != "")
        {
            strSql += " and do.Invoice_No >='" + tbxInvoice_NoS.Text.Trim() + "'";
        }
        if (tbxInvoice_NoE.Text.Trim() != "")
        {
            strSql += " and do.Invoice_No <='" + tbxInvoice_NoE.Text.Trim() + "'";
        }
        //收據開立
        bool first = true; int cnt = 0;
        for (int i = 0; i < cblInvoice_Type.Items.Count; i++)
        {
            if (cblInvoice_Type.Items[i].Selected && first)
            {
                strSql += " and (do.Invoice_Type='" + cblInvoice_Type.Items[i].ToString() + "' "; first = false; cnt++;
            }
            else if (cblInvoice_Type.Items[i].Selected)
            {
                strSql += " or do.Invoice_Type='" + cblInvoice_Type.Items[i].ToString() + "' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //捐款方式
        first = true; cnt = 0;
        for (int i = 0; i < cblDonate_Payment.Items.Count; i++)
        {
            if (cblDonate_Payment.Items[i].Selected && first)
            {
                if (i != cblDonate_Payment.Items.Count - 1)
                {
                    strSql += " and (do.Donate_Payment='" + cblDonate_Payment.Items[i].ToString() + "' "; first = false; cnt++;
                }
                else
                {
                    strSql += @" and (do.Donate_Payment<>'現金' and do.Donate_Payment<>'劃撥' and do.Donate_Payment<>'網路信用卡' and do.Donate_Payment<>'支票' and 
                                 do.Donate_Payment<>'ACH轉帳授權書' and do.Donate_Payment<>'匯款' and do.Donate_Payment<>'ATM' and 
                                 do.Donate_Payment<>'信用卡授權書(聯信)' and do.Donate_Payment<>'信用卡授權書(一般)' and do.Donate_Payment<>'郵局帳戶轉帳授權書'  "; first = false; cnt++;
                }
            }
            else if (cblDonate_Payment.Items[i].Selected)
            {
                if (i != cblDonate_Payment.Items.Count - 1)
                {
                    strSql += " or do.Donate_Payment='" + cblDonate_Payment.Items[i].ToString() + "' "; cnt++;
                }
                else
                {
                    strSql += @" or (do.Donate_Payment<>'現金' and do.Donate_Payment<>'劃撥' and do.Donate_Payment<>'網路信用卡' and do.Donate_Payment<>'支票' and 
                                 do.Donate_Payment<>'ACH轉帳授權書' and do.Donate_Payment<>'匯款' and do.Donate_Payment<>'ATM' and 
                                 do.Donate_Payment<>'信用卡授權書(聯信)' and do.Donate_Payment<>'信用卡授權書(一般)' and do.Donate_Payment<>'郵局帳戶轉帳授權書' ) "; first = false; cnt++;
                }
            }
        }
        if (cnt != 0) strSql += ')';
        //捐款用途
        first = true; cnt = 0;
        for (int i = 0; i < cblDonate_Purpose.Items.Count; i++)
        {
            if (cblDonate_Purpose.Items[i].Selected && first)
            {
                strSql += " and (do.Donate_Purpose='" + cblDonate_Purpose.Items[i].ToString() + "' "; first = false; cnt++;
            }
            else if (cblDonate_Purpose.Items[i].Selected)
            {
                strSql += " or do.Donate_Purpose='" + cblDonate_Purpose.Items[i].ToString() + "' "; cnt++;
            }
        }
        if (cnt != 0) strSql += ')';
        //地址
        if (rblLocal.SelectedValue == "1")
        {
            strSql += " and dr.IsAbroad = 'N' ";
        }
        if (rblLocal.SelectedValue == "2")
        {
            strSql += " and dr.IsAbroad = 'Y' ";
        }
        //其他
        if (cblExp_Pre.Items[0].Selected)
        {
            strSql += " and do.Invoice_Pre='B' ";
        }
        else if (cblExp_Pre.Items[1].Selected)
        {
            strSql += " and do.Export='Y' ";
        }
        //名條內容 
        if (rblContent.SelectedValue == "1")
        {
            strSql = strSql.Replace("WhereClause1", "IsNull(dr.Invoice_ZipCode,'') as [郵遞區號],(Case When dr.Invoice_City='' Then dr.Invoice_Address Else Case When C.mValue<>D.mValue Then IsNull(C.mValue,'') + IsNull(D.mValue,'')+dr.Invoice_Address Else IsNull(C.mValue,'')+dr.Invoice_Address End End) as [地址],dr.Invoice_ZipCode as [Invoice_ZipCode],(Case When dr.Invoice_City='' Then dr.Invoice_Address Else Case When C.mValue<>D.mValue Then C.mValue+D.mValue+dr.Invoice_Address Else C.mValue+dr.Invoice_Address End End) as [Invoice_Address]");
            strSql += " and IsNull(dr.Invoice_City,'') + IsNull(dr.Invoice_Area,'') + dr.Invoice_Address<>'' ";
        }
        if (rblContent.SelectedValue == "2")
        {
            strSql = strSql.Replace("WhereClause1", "IsNull(dr.ZipCode,'') as [郵遞區號],(Case When dr.City='' Then dr.Address Else Case When A.mValue<>B.mValue Then IsNull(A.mValue,'') + IsNull(B.mValue,'')+dr.Address Else IsNull(A.mValue,'')+dr.Address End End) as [地址],dr.ZipCode as [ZipCode],(Case When dr.City='' Then dr.Address Else Case When A.mValue<>B.mValue Then A.mValue+B.mValue+dr.Address Else A.mValue+dr.Address End End) as [Address]");
            strSql += " and  IsNull(dr.City,'') + IsNull(dr.Area,'') + dr.Address<>'' ";
        }
        //不含身份別主知名
        if (cbxPrint_NoName.Checked)
        {
            strSql += " and IsNull(dr.Donor_Type,'') Not Like '%主知名%' ";
        }
        return strSql;
    }
    private string Sort(string strSql)
    {
        ////排序方式
        for (int i = 0; i < rblSort.Items.Count; i++)
        {
            if (rblSort.Items[i].Selected)
            {
                switch (i)
                {
                    case 0:
                        /*strSql = strSql.Replace("WhereClause2", @" from (select top 1000 * from Donate order by Invoice_No)do left join Donor dr  
                                                                                                Left Join CODECITY As A On dr.City=A.mCode Left Join CODECITY As B On dr.Area=B.mCode
                                                                                                Left Join CODECITY As C On dr.Invoice_City=C.mCode Left Join CODECITY As D On dr.Invoice_Area=D.mCode"); */
                        //20140521修改SQL by 詩儀
                        strSql = strSql.Replace("WhereClause2", @" from Donate do left join Donor dr  
                                                                                                Left Join CODECITY As A On dr.City=A.mCode Left Join CODECITY As B On dr.Area=B.mCode
                                                                                                Left Join CODECITY As C On dr.Invoice_City=C.mCode Left Join CODECITY As D On dr.Invoice_Area=D.mCode");
                        strSql += " order by do.Invoice_No"; break;
                    case 1:
                        strSql = strSql.Replace("WhereClause2", @" from Donate do Left Join Donor dr 
                                                                                  Left Join CODECITY As A On dr.City=A.mCode Left Join CODECITY As B On dr.Area=B.mCode
                                                                                  Left Join CODECITY As C On dr.Invoice_City=C.mCode Left Join CODECITY As D On dr.Invoice_Area=D.mCode");
                        strSql += " order by dr.Donor_Id"; break;
                    case 2:
                        strSql = strSql.Replace("WhereClause2", @" from Donate do Left Join Donor dr 
                                                                                  Left Join CODECITY As A On dr.City=A.mCode Left Join CODECITY As B On dr.Area=B.mCode
                                                                                  Left Join CODECITY As C On dr.Invoice_City=C.mCode Left Join CODECITY As D On dr.Invoice_Area=D.mCode");
                        strSql += " order by [郵遞區號]"; break;
                }
            }
        }
        return strSql;
    }
}