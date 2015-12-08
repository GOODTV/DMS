using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateOnlineMedia : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            List<ControlData> list = new List<ControlData>();
            list.Add(new ControlData("Checkbox", "Item", "CHK_ItemOnce", "單筆奉獻", "單筆奉獻", false, ""));
            lblItemOnce.Text = HtmlUtil.RenderControl(list);
            list = new List<ControlData>();
            list.Add(new ControlData("Checkbox", "Item", "CHK_ItemPeriod", "定期定額奉獻", "定期定額奉獻", false, ""));
            lblItemPeriod.Text = HtmlUtil.RenderControl(list);

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
        ddlYearMS.Items.Add(new ListItem("請選擇", "請選擇"));
        ddlYearME.Items.Add(new ListItem("請選擇", "請選擇"));
        ddlYearQS.Items.Add(new ListItem("請選擇", "請選擇"));
        ddlYearQE.Items.Add(new ListItem("請選擇", "請選擇"));
        ddlYearS.Items.Add(new ListItem("請選擇", "請選擇"));
        ddlYearE.Items.Add(new ListItem("請選擇", "請選擇"));
        for (int i = DateTime.Now.Year; i < DateTime.Now.Year + 5; i++)
        {
            ddlYearMS.Items.Add(new ListItem(i.ToString(), i.ToString()));
            ddlYearME.Items.Add(new ListItem(i.ToString(), i.ToString()));
            ddlYearQS.Items.Add(new ListItem(i.ToString(), i.ToString()));
            ddlYearQE.Items.Add(new ListItem(i.ToString(), i.ToString()));
            ddlYearS.Items.Add(new ListItem(i.ToString(), i.ToString()));
            ddlYearE.Items.Add(new ListItem(i.ToString(), i.ToString()));
        }
        ddlMonthS.Items.Add(new ListItem("請選擇", "請選擇"));
        ddlMonthE.Items.Add(new ListItem("請選擇", "請選擇"));
        for (int i = 1; i < 13; i++)
        {
            ddlMonthS.Items.Add(new ListItem(i.ToString(), i.ToString()));
            ddlMonthE.Items.Add(new ListItem(i.ToString(), i.ToString()));
        }
        ddlQtrS.Items.Add(new ListItem("請選擇", "請選擇"));
        ddlQtrE.Items.Add(new ListItem("請選擇", "請選擇"));
        for (int i = 1; i < 5; i++)
        {
            ddlQtrS.Items.Add(new ListItem(i.ToString(), i.ToString()));
            ddlQtrE.Items.Add(new ListItem(i.ToString(), i.ToString()));
        }
        List<ControlData> list = new List<ControlData>();
        list.Add(new ControlData("Radio", "PayType", "rdoPayType1", "月繳", "月繳", false, ""));
        list.Add(new ControlData("Radio", "PayType", "rdoPayType2", "季繳", "季繳", false, ""));
        list.Add(new ControlData("Radio", "PayType", "rdoPayType3", "年繳", "年繳", false, ""));
        lblPayType.Text = HtmlUtil.RenderControl(list);
    }
    //---------------------------------------------------------------------------
    protected void btnNext_Click(object sender, EventArgs e)
    {
        //Response.Redirect("CheckOut.aspx?QryItemList=" + HFD_ItemList.Value);
        //if (Session["ItemList"] != null)
        //{
        //    Session["ItemList"] = HFD_ItemList.Value;
        //}
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
        if (HFD_chkItem.Value.Contains("單筆奉獻"))
        {
            DataRow dr = dtOnce.NewRow();
            dr["Uid"] = dtOnce.Rows.Count;
            dr["奉獻項目"] = "優質媒體奉獻 ";
            dr["奉獻金額"] = txtAmountOnce.Text;
            dtOnce.Rows.Add(dr);
        }
        if (HFD_chkItem.Value.Contains("定期定額奉獻"))
        {
            DataRow dr = dtPeriod.NewRow();
            dr["Uid"] = dtPeriod.Rows.Count;
            dr["奉獻項目"] = "優質媒體奉獻 ";
            dr["繳費方式"] = HFD_PayType.Value;
            if (HFD_PayType.Value == "月繳")
            {
                dr["開始年"] = ddlYearMS.SelectedValue;
                dr["開始月"] = ddlMonthS.SelectedValue;
                dr["結束年"] = ddlYearME.SelectedValue;
                dr["結束月"] = ddlMonthE.SelectedValue;
            }
            if (HFD_PayType.Value == "季繳")
            {
                dr["開始年"] = ddlYearQS.SelectedValue;
                dr["開始月"] = ddlQtrS.SelectedValue;
                dr["結束年"] = ddlYearQE.SelectedValue;
                dr["結束月"] = ddlQtrE.SelectedValue;
            }
            if (HFD_PayType.Value == "年繳")
            {
                dr["開始年"] = ddlYearS.SelectedValue;
                dr["開始月"] = "";
                dr["結束年"] = ddlYearE.SelectedValue;
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
    protected void btnPrev_Click(object sender, EventArgs e)
    {
        Response.Redirect("DonateOnlineDefault.aspx");
    }
}