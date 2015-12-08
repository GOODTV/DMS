using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateOnlineAll : System.Web.UI.Page
{
    DataTable dtOnce = new DataTable();
    DataTable dtPeriod = new DataTable();

    protected void Page_Load(object sender, EventArgs e)
    {
        //顯示前夜愈
        if (Session["Msg"] != null)
        {
            Util.ShowMsg(Session["Msg"].ToString());
            Session["Msg"] = null;
        }
        if (Session["ItemOnce"] != null)
        {
            dtOnce = (DataTable)Session["ItemOnce"];
        }
        if (Session["ItemPeriod"] != null)
        {
            dtPeriod = (DataTable)Session["ItemPeriod"];
        }
        //檢查是否已有奉獻項目，以判定前端檢核機制
        if (dtOnce.Rows.Count > 0 || dtPeriod.Rows.Count > 0)
        {
            HFD_NeedCheck.Value = "N";
        }
        else
        {
            HFD_NeedCheck.Value = "Y";
        }

        // 2014/7/17 停留過久移除Session
        if (dtOnce.Rows.Count == 0)
        {
            Session.Remove("ItemOnce");
        }
        if (dtPeriod.Rows.Count == 0)
        {
            Session.Remove("ItemPeriod");
        }

        if (!IsPostBack)
        {
            List<ControlData> list = new List<ControlData>();
            //list.Add(new ControlData("Checkbox", "Item", "CHK_ItemOnce", "單筆奉獻", "單筆奉獻", false, ""));
            list.Add(new ControlData("Radio", "Item", "RDO_Item1", "單筆奉獻", "One time Donation", false, ""));
            lblItemOnce.Text = HtmlUtil.RenderControl(list);
            lblItemOnce.Font.Size = 10;
            list = new List<ControlData>();
            //list.Add(new ControlData("Checkbox", "Item", "CHK_ItemPeriod", "定期定額奉獻", "定期定額奉獻", false, ""));
            list.Add(new ControlData("Radio", "Item", "RDO_Item2", "定期定額奉獻", "Recurring Donation", false, ""));
            lblItemPeriod.Text = HtmlUtil.RenderControl(list);
            lblItemPeriod.Font.Size = 10;

            LoadDropDownListData();

            //if (Session["DonorName"] != null)
            //{
            //    lblTitle.Text = Session["DonorName"].ToString() + lblTitle.Text;
            //}
        }
    }
    //---------------------------------------------------------------------------
    private void LoadDropDownListData()
    {
        //ddlYearMS.Items.Add(new ListItem("請選擇", "請選擇"));
        //ddlYearME.Items.Add(new ListItem("請選擇", "請選擇"));
        //ddlYearQS.Items.Add(new ListItem("請選擇", "請選擇"));
        //ddlYearQE.Items.Add(new ListItem("請選擇", "請選擇"));
        //ddlYearS.Items.Add(new ListItem("請選擇", "請選擇"));
        //ddlYearE.Items.Add(new ListItem("請選擇", "請選擇"));
        //for (int i = DateTime.Now.Year; i < DateTime.Now.Year + 5; i++)
        //{
        //    ddlYearMS.Items.Add(new ListItem(i.ToString(), i.ToString()));
        //    ddlYearME.Items.Add(new ListItem(i.ToString(), i.ToString()));
        //    ddlYearQS.Items.Add(new ListItem(i.ToString(), i.ToString()));
        //    ddlYearQE.Items.Add(new ListItem(i.ToString(), i.ToString()));
        //    ddlYearS.Items.Add(new ListItem(i.ToString(), i.ToString()));
        //    ddlYearE.Items.Add(new ListItem(i.ToString(), i.ToString()));
        //}
        //ddlMonthS.Items.Add(new ListItem("請選擇", "請選擇"));
        //ddlMonthE.Items.Add(new ListItem("請選擇", "請選擇"));
        //for (int i = 1; i < 13; i++)
        //{
        //    ddlMonthS.Items.Add(new ListItem(i.ToString(), i.ToString()));
        //    ddlMonthE.Items.Add(new ListItem(i.ToString(), i.ToString()));
        //}
        //ddlQtrS.Items.Add(new ListItem("請選擇", "請選擇"));
        //ddlQtrE.Items.Add(new ListItem("請選擇", "請選擇"));
        //for (int i = 1; i < 5; i++)
        //{
        //    ddlQtrS.Items.Add(new ListItem(i.ToString(), i.ToString()));
        //    ddlQtrE.Items.Add(new ListItem(i.ToString(), i.ToString()));
        //}
        //List<ControlData> list = new List<ControlData>();
        //list.Add(new ControlData("Radio", "PayType", "rdoPayType1", "月繳", "月繳", false, ""));
        //list.Add(new ControlData("Radio", "PayType", "rdoPayType2", "季繳", "季繳", false, ""));
        //list.Add(new ControlData("Radio", "PayType", "rdoPayType3", "年繳", "年繳", false, ""));
        //lblPayType.Text = HtmlUtil.RenderControl(list);
    }
    //---------------------------------------------------------------------------
    protected void btnNext_Click(object sender, EventArgs e)
    {
        DataTable dtOnce = new DataTable();
        DataTable dtPeriod = new DataTable();
        if (Session["ItemOnce"] != null)
        {
            dtOnce = (DataTable)Session["ItemOnce"];
        }
        else
        {
            dtOnce.Columns.Add("Uid");
            dtOnce.Columns.Add("奉獻項目");
            dtOnce.Columns.Add("奉獻金額");
        }

        if (Session["ItemPeriod"] != null)
        {
            dtPeriod = (DataTable)Session["ItemPeriod"];
        }
        else
        {
            dtPeriod.Columns.Add("Uid");
            dtPeriod.Columns.Add("奉獻項目");
            dtPeriod.Columns.Add("繳費方式");
            dtPeriod.Columns.Add("開始年");
            dtPeriod.Columns.Add("開始月");
            dtPeriod.Columns.Add("結束年");
            dtPeriod.Columns.Add("結束月");
            dtPeriod.Columns.Add("奉獻金額");
        }

        //string strPurpose = rblPurpose.SelectedValue;
        string strPurpose = "Gift for GOODTV";//string.Empty;
        /*if (rdoPurpose1.Checked)
            strPurpose = rdoPurpose1.Text;
        else if (rdoPurpose2.Checked)
            strPurpose = rdoPurpose2.Text;*/
        
        if (HFD_chkItem.Value.Contains("單筆奉獻"))
        {
            DataRow dr = dtOnce.NewRow();
            dr["Uid"] = dtOnce.Rows.Count;
            dr["奉獻項目"] = strPurpose;
            dr["奉獻金額"] = txtAmountOnce.Text;
            dtOnce.Rows.Add(dr);
        }
        if (HFD_chkItem.Value.Contains("定期定額奉獻"))
        {
            DataRow dr = dtPeriod.NewRow();
            dr["Uid"] = dtPeriod.Rows.Count;
            dr["奉獻項目"] = strPurpose;
            dr["繳費方式"] = HFD_PayType.Value;
            if (HFD_PayType.Value == "月繳")
            {
                dr["開始年"] = Util.GetDBDateTime().Year.ToString();
                dr["開始月"] = Util.GetDBDateTime().Month.ToString();
                dr["結束年"] = "2099";
                dr["結束月"] = "12";
            }
            if (HFD_PayType.Value == "季繳")
            {
                //dr["開始年"] = ddlYearQS.SelectedValue;
                //dr["開始月"] = ddlQtrS.SelectedValue;
                //dr["結束年"] = ddlYearQE.SelectedValue;
                //dr["結束月"] = ddlQtrE.SelectedValue;
                dr["開始年"] = Util.GetDBDateTime().Year.ToString();
                dr["開始月"] = Util.GetDBDateTime().Month.ToString();
                dr["結束年"] = "2099";
                dr["結束月"] = "12";
            }
            if (HFD_PayType.Value == "年繳")
            {
                dr["開始年"] = Util.GetDBDateTime().Year.ToString();
                dr["開始月"] = "";
                dr["結束年"] = "2099";
                dr["結束月"] = "";
            }
            dr["奉獻金額"] = txtAmountPeriod.Text;
            dtPeriod.Rows.Add(dr);
        }

        Session["ItemOnce"] = dtOnce;
        Session["ItemPeriod"] = dtPeriod;

        Response.Redirect("ShoppingCart.aspx");
    }
    //---------------------------------------------------------------------------
    protected void btnCancel_Click(object sender, EventArgs e)
    {
        Response.Redirect("ShoppingCart.aspx");
    }
}