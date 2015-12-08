using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;

public partial class Online_DonateQry : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {
            //get Server Time
            DateTime dtNow = Util.GetDBDateTime();
            //2013-05-06 Modify by GoodTV Tanya
            if (dtNow.Year >= 2013)
                ddlYearMS.Items.Add(new ListItem(dtNow.Year.ToString(), dtNow.Year.ToString()));

            if (dtNow.Year > 2013)
                ddlYearMS.Items.Add(new ListItem(dtNow.AddYears(-1).Year.ToString(), dtNow.AddYears(-1).Year.ToString()));

            if (ddlYearMS.Items.Count > 0)
            {
                ddlYearMS.SelectedIndex = 0;
                AddDdlMonth(ddlYearMS, ddlMonthS);
            }

            //            lblFootMsg.Text = @"
            //                好消息電視台 製作 版權所有 GOOD TV Broadcasting Corp.<br/>
            //                新北市中和區中正路911號6樓 TEL:02-8024-3911 FAX:02-8024-3938<br/>
            //                本網站所載之所有著作、音樂及網頁畫面資料之安排，其著作權、<br/>
            //                均已獲得其權利人所授權，任何人若非事先經權利人及本網站授權，不得任意轉載。<br/>
            //                ";
        }        
    }    
    //---------------------------------------------------------------------------
    private void AddDdlMonth(DropDownList ddlYearTarget, DropDownList ddlMonthTarget)
    {
        int iTopMonth = 12;
        //get Server Time
        DateTime dtNow = Util.GetDBDateTime();
        //最長查詢期間為去年度及到目前時間的上個月
        if (ddlYearTarget.SelectedValue == dtNow.Year.ToString())
        {
            //並每個月15號之後才能開始查詢上個月的記錄
            if (dtNow.Day > 15)
            {
                iTopMonth = dtNow.Month - 1;
            }
            else
            {
                iTopMonth = dtNow.Month - 2;
            }
        }

        ddlMonthTarget.Items.Clear();
        for (int i = 1; i <= iTopMonth; i++)
        {
            ddlMonthTarget.Items.Add(new ListItem(i.ToString("00"),i.ToString("00")));
        }
    }
    //---------------------------------------------------------------------------
    private string CheckQueryItem()
    {
        string strRet = "";
        if (ddlMonthS.SelectedValue == "")
        {
            strRet = "請選擇查詢月份！";
        }
        return strRet;
    }
    //---------------------------------------------------------------------------
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        string strRet = CheckQueryItem();
        if (strRet != "")
        {
            ScriptManager.RegisterClientScriptBlock(UpdatePanel1, typeof(UpdatePanel), "Warning", "alert('" + strRet + "');", true);
            return;
        }
        
        if (txtDonorName.Text.Trim().Length <= 1)
        {
            ScriptManager.RegisterClientScriptBlock(UpdatePanel1, typeof(UpdatePanel), "Warning", "alert('【奉獻姓名/團體】至少需輸入2個字！');", true);
            return;
        }
        else if (txtDonorName.Text.IndexOf('%') > -1)
        {
            ScriptManager.RegisterClientScriptBlock(UpdatePanel1, typeof(UpdatePanel), "Warning", "alert('【奉獻姓名/團體】請勿輸入特殊字元！');", true);
            return;
        }
        //}       

        //2013-05-02 Modify by GoodTV Tanya:Call DB SP
        string strConn = System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
        SqlConnection conn = new SqlConnection();
        conn.ConnectionString = strConn;
        conn.Open();
        //SqlCommand cmd = new SqlCommand(strSql, conn);

        SqlCommand cmd = new SqlCommand("dbo.[getDonateRecord]", conn);
        cmd.CommandType = CommandType.StoredProcedure;

        cmd.Parameters.Add("@NAME", SqlDbType.VarChar).Value = txtDonorName.Text.Trim();
        cmd.Parameters.Add("@YYYY", SqlDbType.VarChar).Value = ddlYearMS.SelectedValue;
        cmd.Parameters.Add("@MM", SqlDbType.VarChar).Value = ddlMonthS.SelectedValue.PadLeft(2, '0');

        //SqlParameter param = new SqlParameter();
        //param.ParameterName = "@REVDT";
        //param.Value = ddlYearMS.SelectedValue + ddlMonthS.SelectedValue + "%";
        //cmd.Parameters.Add(param);
        //SqlParameter param1 = new SqlParameter();
        //param1.ParameterName = "@NAME";
        //param1.Value = "%" + txtDonorName.Text.Trim() + "%";
        //cmd.Parameters.Add(param1);

        SqlDataReader sdr = cmd.ExecuteReader();

        DataTable dt = new DataTable();
        dt.Load(sdr);

        gdvQry.DataSource = dt;
        gdvQry.DataBind();
    }
    //---------------------------------------------------------------------------
    protected void gdvQry_DataBound(object sender, EventArgs e)
    {
        for (int i = 0; i < gdvQry.Rows.Count; i++)
        {
            gdvQry.Rows[i].Cells[0].CssClass = "aligncenter";
            gdvQry.Rows[i].Cells[1].CssClass = "aligncenter";
            gdvQry.Rows[i].Cells[2].CssClass = "alignright";
            gdvQry.Rows[i].Cells[3].CssClass = "aligncenter";
            //string strMask = GetMaskResult(gdvQry.Rows[i].Cells[1].Text, "X");
            string strMask = GetMaskResult(gdvQry.Rows[i].Cells[1].Text, "Ο");
            gdvQry.Rows[i].Cells[1].Text = strMask;
        }

        if (gdvQry.Rows.Count > 0)
        {
            gdvQry.HeaderRow.Cells[0].CssClass = "aligncenter";
            gdvQry.HeaderRow.Cells[1].CssClass = "aligncenter";
            gdvQry.HeaderRow.Cells[2].CssClass = "alignright";
            gdvQry.HeaderRow.Cells[3].CssClass = "aligncenter";
        }
    }
    //---------------------------------------------------------------------------
    private string GetMaskResult(string strTarget, string strMask)
    {
        string strRet = "";
        string strTmp = strTarget.Trim();
        string strReplace = "";

        for (int i = 0; i < strTmp.Length - 2; i++)
        {
            strReplace += strMask;
        }

        strRet = strTmp.Substring(0, 1) + strReplace + strTmp.Substring(strTmp.Length - 1, 1);

        //2個字時
        if (strRet.Length == 2)
        {
            strRet = strRet.Substring(0, 1) + strMask;
        }
        return strRet;
    }
    //---------------------------------------------------------------------------
    protected void ddlYearMS_SelectedIndexChanged(object sender, EventArgs e)
    {
        AddDdlMonth(ddlYearMS, ddlMonthS);
    }
}