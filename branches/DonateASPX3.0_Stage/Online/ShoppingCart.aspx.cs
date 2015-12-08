using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_ShoppingCart : System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {

            //捐款方式
            if (!String.IsNullOrEmpty(Request.Params["HFD_chkItem"]))
            {
                HFD_chkItem.Value = Request.Params["HFD_chkItem"];
            }
            //奉獻金額
            if (!String.IsNullOrEmpty(Request.Params["HFD_Amount"]))
            {
                HFD_Amount.Value = Request.Params["HFD_Amount"];
            }
            //帳號(Email)
            if (!String.IsNullOrEmpty(Request.Params["HFD_Email"]))
            {
                txtAccount.Text = Request.Params["HFD_Email"];
            }
            if (HFD_chkItem.Value.Contains("單筆奉獻"))
            {
                ShowCartOnce();
                lblStep3.Text = " 確認奉獻明細 ";
                lblStep4.Visible = false;
                lblStep5.Visible = false;
                PeriodMemo.Text = "";
            }
            else if (HFD_chkItem.Value.Contains("定期定額奉獻"))
            {
                HFD_PayType.Value = "月繳";
                ShowCartPeriod();
                PeriodMemo.Text = "<span style='color: blue;'>1.定期定額捐款扣款方式：捐款金額月繳是每月扣款，季繳為每三個月扣款一次，以此類推。<br/><br/>"
                //+ "<!--信用卡僅提供<font color='red'>VISA、MASTER、JCB</font> 3種卡別。<br><br>-->"
                + "2.若需終止定期定額奉獻，請電洽-奉獻服務專線(02)8024-3911 捐款服務組，謝謝您的支持。</span><br/><br/>";
            }
            else 
            {
                Response.Redirect("DonateOnlineAll.aspx");
                //lblGrid.Text = @"<span align='center' style='color: Black; font-weight: bold;'>※ 您的清單目前沒有項目。</span><br/>";
            }

        }


    }

    //-------------------------------------------------------------------------
    public void ShowCartOnce()
    {

        lblGrid.Text = @"<span style='color: black; font-weight: bold;' align='center'>※ 單 筆 奉 獻</span>";
        lblGrid.Text += "<table width='100%' class='table_h'><tr style='background-color: rgb(205, 225, 226);'>";
        lblGrid.Text += "<th>奉獻項目</th><th>金額</th></tr><tr><td style='text-align: center;'>為GOODTV奉獻</td>";
        lblGrid.Text += "<td style='text-align: right;'><span><font color='red'>NT$ " + 
            String.Format("{0:0,0}", Convert.ToInt32(HFD_Amount.Value)) + "</font></span></td></tr></table>";

    }

    //-------------------------------------------------------------------------
    public void ShowCartPeriod()
    {

        lblGrid.Text = @"<span style='color: black; font-weight: bold;' align='center'>※ 定 期 定 額 奉 獻</span>";
        lblGrid.Text += "<table width='100%' class='table_h'><tr><th>奉獻項目</th>";
        lblGrid.Text += "<th>金額</th><th>奉獻週期</th></tr><tr><td style='text-align: center;'>為GOODTV奉獻</td>";
        lblGrid.Text += "<td style='text-align: right;'><span><font color='red'>NT$ " +
            String.Format("{0:0,0}", Convert.ToInt32(HFD_Amount.Value)) + "</font></span></td>";
        lblGrid.Text += "<td style='text-align: center;'><span><select name='ddlPeriod' id='ddlPeriod'>";
        lblGrid.Text += "<option value='年繳'>年繳</option><option value='季繳'>季繳</option>";
        lblGrid.Text += "<option value='月繳' selected=''>月繳</option></select></span></td></tr></table>";

    }

}
