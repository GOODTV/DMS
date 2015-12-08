using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateOnline : System.Web.UI.Page
{
    DataTable dtCart = new DataTable();

    protected void Page_Load(object sender, EventArgs e)
    {
        //有 npoGridView 時才需要
        //Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (!IsPostBack)
        {
            dtCart.Columns.Add("Uid");
            dtCart.Columns.Add("奉獻項目");
            dtCart.Columns.Add("奉獻金額");
            Session["dtCart"] = dtCart;
        }
        else
        {
            if (Session["dtCart"] != null)
            {
                dtCart = (DataTable)Session["dtCart"];
            }
        }
    }
    //-------------------------------------------------------------------------
    public void ShowCart(DataTable dtSrc)
    {
        if (dtSrc.Rows.Count > 0)
        {
            //Grid initial
            //沒有需要特別處理的欄位時
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dtSrc;
            npoGridView.Keys.Add("Uid");
            npoGridView.DisableColumn.Add("Uid");
            npoGridView.ShowPage = false;
            
            

            //npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
            //npoGridView.EditLink = Util.RedirectByTime("ReportInputDay_Edit.aspx", "Uid=");
//            lblGridList.Text = @"                
//                <tr>
//                    <td align='center' 
//                        style='background-color: #D3D3D3; color: Black; font-weight: bold'>
//                        奉　獻　項　目　清　單
//                    </td>
//                </tr>
//                " + npoGridView.Render();
            lblGridList.Text = @"                
                    <span align='center' width='100%' 
                        style='background-color: #D3D3D3; color: Black; font-weight: bold'>
                        奉　獻　項　目　清　單
                    </span>
                " + npoGridView.Render();
            //lblGridList.Text = npoGridView.Render();
        }
        else
        {
            lblGridList.Text = "";
        }
    }
    //---------------------------------------------------------------------------
    private void dtCartChange(CheckBox chkSrc,TextBox txtSrc,CompareValidator cvSrc)
    {
        if (chkSrc.Checked)
        {
            cvSrc.Validate();
            if (!cvSrc.IsValid)
            {
                // ScriptManager.RegisterClientScriptBlock(UpdatePanel1, typeof(UpdatePanel), "test", "alert('test');", true);
                //chkSrc.Checked = false;
                txtSrc.Focus();
                txtSrc.Text = "";
            }
            DataRow dr = dtCart.NewRow();
            dr["Uid"] = chkSrc.ID.Substring(chkSrc.ID.Length-1, 1);
            dr["奉獻項目"] = chkSrc.Text;
            dr["奉獻金額"] = txtSrc.Text;
            dtCart.Rows.Add(dr);
        }
        else
        {
            DataRow dr = dtCart.Select("Uid='" + chkSrc.ID.Substring(chkSrc.ID.Length - 1, 1) + "'")[0];
            dtCart.Rows.Remove(dr);
            txtSrc.Text = "";
        }
       
        ShowCart(dtCart);
    }
    //---------------------------------------------------------------------------
    protected void chkItem1_CheckedChanged(object sender, EventArgs e)
    {
        dtCartChange(chkItem1, txtAmount1, cvAmount1);
    }
    //---------------------------------------------------------------------------
    protected void chkItem2_CheckedChanged(object sender, EventArgs e)
    {
        dtCartChange(chkItem2, txtAmount2, cvAmount2);
    }
    //---------------------------------------------------------------------------
    protected void chkItem3_CheckedChanged(object sender, EventArgs e)
    {
        dtCartChange(chkItem3, txtAmount3, cvAmount3);
    }
    //---------------------------------------------------------------------------
    protected void chkItem4_CheckedChanged(object sender, EventArgs e)
    {
        dtCartChange(chkItem4, txtAmount4, cvAmount4);
    }
    //---------------------------------------------------------------------------
    protected void chkItem5_CheckedChanged(object sender, EventArgs e)
    {
        dtCartChange(chkItem5, txtAmount5, cvAmount5);
    }
    //---------------------------------------------------------------------------
    private void UpdateCartAmount(string strUid, TextBox txtSrc,CompareValidator cvSrc)
    {
        cvSrc.Validate();
        if (cvSrc.IsValid)
        {
            if (dtCart.Select("Uid='" + strUid + "'").Length != 0)
            {
                DataRow dr = dtCart.Select("Uid='" + strUid + "'")[0];
                dr["奉獻金額"] = txtSrc.Text;
                ShowCart(dtCart);
            }
        }
        else
        {
            // ScriptManager.RegisterClientScriptBlock(UpdatePanel1, typeof(UpdatePanel), "test", "alert('test');", true);
            txtSrc.Focus();
            txtSrc.Text = "";
        }
    }
    //---------------------------------------------------------------------------
    protected void txtAmount1_TextChanged(object sender, EventArgs e)
    {
        UpdateCartAmount("1", txtAmount1, cvAmount1);
    }
    //---------------------------------------------------------------------------
    protected void txtAmount2_TextChanged(object sender, EventArgs e)
    {
        UpdateCartAmount("2", txtAmount2, cvAmount2);
    }
    //---------------------------------------------------------------------------
    protected void txtAmount3_TextChanged(object sender, EventArgs e)
    {
        UpdateCartAmount("3", txtAmount3, cvAmount3);
    }
    //---------------------------------------------------------------------------
    protected void txtAmount4_TextChanged(object sender, EventArgs e)
    {
        UpdateCartAmount("4", txtAmount4, cvAmount4);
    }
    //---------------------------------------------------------------------------
    protected void txtAmount5_TextChanged(object sender, EventArgs e)
    {
        UpdateCartAmount("5", txtAmount5, cvAmount5);
    }
    //---------------------------------------------------------------------------
}