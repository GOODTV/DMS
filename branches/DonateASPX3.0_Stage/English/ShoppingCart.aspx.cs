using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_ShoppingCart : System.Web.UI.Page
{
    DataTable dtOnce = new DataTable();
    DataTable dtPeriod = new DataTable();
    DataTable dtCart = new DataTable();

    protected void Page_Load(object sender, EventArgs e)
    {
        //有 npoGridView 時才需要
        //Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");

        if (Session["ItemOnce"] != null)
        {
            dtOnce = (DataTable)Session["ItemOnce"];
        }
        if (Session["ItemPeriod"] != null)
        {
            dtPeriod = (DataTable)Session["ItemPeriod"];
        }

        if (!IsPostBack)
        {
            // 用來判斷到是否有新增新的奉獻(沒有文字=沒有新增奉獻)，到ShoppingCart.aspx頁面時是否要跳出"已加入奉獻清單"的提示訊息 by Hilty 2013/12/18
            if (Session["AddShopping"] != null)
            {
                AddShopping.Value = Session["AddShopping"].ToString();
                Session.Remove("AddShopping"); 
            }

            //HFD_ItemList.Value = Util.GetQueryString("QryItemList");
            HFD_DelOnce.Value = Util.GetQueryString("DelOnce");
            HFD_DelPeriod.Value = Util.GetQueryString("DelPeriod");
            HFD_Mode.Value = Util.GetQueryString("Mode");
            //傳入刪除項目時，先移除該項目
            if (HFD_DelOnce.Value != "")
            {
                for (int i = 0; i < dtOnce.Rows.Count; i++)
                {
                    if (dtOnce.Rows[i]["Uid"].ToString() == HFD_DelOnce.Value)
                    {
                        dtOnce.Rows.RemoveAt(i);
                        break;
                    }
                }
                for (int i = 0; i < dtOnce.Rows.Count; i++)
                {
                    dtOnce.Rows[i]["Uid"] = i;
                }
                lblGridOnce.Text = @"                
                    <span align='center' style='color: Black; font-weight: bold; font-size: medium;'>
                        ※ There is no item in your cart.
                    </span>";
            }
            else
            {
                btnContinue.Visible = false;
            }
            if (HFD_DelPeriod.Value != "")
            {
                for (int i = 0; i < dtPeriod.Rows.Count; i++)
                {
                    if (dtPeriod.Rows[i]["Uid"].ToString() == HFD_DelPeriod.Value)
                    {
                        dtPeriod.Rows.RemoveAt(i);
                        break;
                    }
                }
                for (int i = 0; i < dtPeriod.Rows.Count; i++)
                {
                    dtPeriod.Rows[i]["Uid"] = i;
                }
                lblGridPeriod.Text = @"                
                    <span align='center' style='color: Black; font-weight: bold; font-size: medium;'>
                        ※ There is no item in your cart.
                    </span>";
            }
            else
            {
                btnContinue.Visible = false;
            }

            if (dtOnce.Rows.Count > 0)
            {
                ShowCartOnce(dtOnce);
            }

            if (dtPeriod.Rows.Count > 0)
            {
                ShowCartPeriod(dtPeriod);
            }

            if (dtOnce.Rows.Count == 0 && dtPeriod.Rows.Count == 0)
            {
                btnCheckOut.Visible = false;
                btnContinue.Visible = true;
            }
            else
            {
                btnCheckOut.Visible = true;
            }
        }

        int iOnce = dtOnce.Rows.Count;
        int iPeriod= dtPeriod.Rows.Count;
        lblCart.ForeColor = System.Drawing.Color.DarkRed;
        //lblCart.Text = "平安！ 您的奉獻明細如下：<br/>";
		lblCart.Text = "In your cart: one time donation (" + iOnce.ToString() + ")；recurring donation (" + iPeriod.ToString() + ")";
        //if (Session["DonorName"] != null)
        //{
          //  lblCart.Text = Session["DonorName"].ToString() + ", " + lblCart.Text;
        //}
        
    }
    //-------------------------------------------------------------------------
    public void ShowCartOnce(DataTable dtSrc)
    {
        //if (dtSrc.Rows.Count > 0)
        //{
            DataTable dtSum = dtSrc.Clone();
            //計算合計
            int iSum = 0;
            DataRow drSum = dtSum.NewRow();
            //DataColumn[] keys = new DataColumn[1];
            //keys[0] = dtSum.Columns[0];
            //dtSum.PrimaryKey = keys;
            if (dtSrc.Rows.Count > 0)
            {
                for (int i = 0; i < dtSrc.Rows.Count; i++)
                {
                    iSum += Convert.ToInt32(dtSrc.Rows[i]["奉獻金額"].ToString());
                    drSum = dtSum.NewRow();
                    drSum["Uid"] = dtSrc.Rows[i]["Uid"].ToString();
                    drSum["奉獻項目"] = dtSrc.Rows[i]["奉獻項目"].ToString();
                    drSum["奉獻金額"] = dtSrc.Rows[i]["奉獻金額"].ToString();
                    dtSum.Rows.Add(drSum);
                }
                drSum = dtSum.NewRow();
                drSum["奉獻項目"] = "Total";
                drSum["奉獻金額"] = iSum.ToString();
                dtSum.Rows.Add(drSum);
            }

            //Grid initial
            //沒有需要特別處理的欄位時
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dtSum;
            npoGridView.Keys.Add("Uid");
            npoGridView.DisableColumn.Add("Uid");
            npoGridView.ShowPage = false;
            npoGridView.TableID = "ItemOnce";

            //-------------------------------------------------------------------------
            NPOGridViewColumn col = new NPOGridViewColumn("Description");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("奉獻項目");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("Amount");
            col.ColumnType = NPOColumnType.CurrecyText;
            col.ColumnName.Add("奉獻金額");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            //欄位為 Button
            col = new NPOGridViewColumn("");
            col.ColumnType = NPOColumnType.Link;
            col.ControlKeyColumn.Add("Uid");
            col.Link = Util.RedirectByTime("ShoppingCart.aspx", "DelOnce=");
            col.LinkText="Delete";
            col.ShowConfirmDialog = true;
            col.ConfirmDialogMsg = "Are you sure to delete this item?";
            npoGridView.Columns.Add(col);

            lblGridOnce.Text = @"                
                    <span align='center' style='color: Black; font-weight: bold; font-size: medium;'>
                        ※ One Time Donation
                    </span>
                " + npoGridView.Render();
        //}
        //else
        //{
        //    lblGridOnce.Text = "";
        //}
        
    }
    //-------------------------------------------------------------------------
    public void ShowCartPeriod(DataTable dtSrc)
    {
        //if (dtSrc.Rows.Count > 0)
        //{
            DataTable dtNew = new DataTable();

            dtNew.Columns.Add("Uid");
            dtNew.Columns.Add("奉獻項目");
            dtNew.Columns.Add("繳費方式");
            dtNew.Columns.Add("奉獻金額");

            for (int i = 0; i < dtSrc.Rows.Count; i++)
            {
                DataRow drNew = dtNew.NewRow();
                drNew["Uid"] = dtSrc.Rows[i]["Uid"].ToString();
                drNew["奉獻項目"] = dtSrc.Rows[i]["奉獻項目"].ToString();
                string strContent = "";
                ////月繳
                //if (dtSrc.Rows[i]["繳費方式"].ToString() == "月繳")
                //{
                //    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年";
                //    //strContent += dtSrc.Rows[i]["開始月"].ToString() + "月 開始；";
                //    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年";
                //    //strContent += dtSrc.Rows[i]["開始月"].ToString() + "月至";
                //    //strContent += dtSrc.Rows[i]["結束年"].ToString() + "年";
                //    //strContent += dtSrc.Rows[i]["結束月"].ToString() + "月止；";
                //}
                ////季繳
                //if (dtSrc.Rows[i]["繳費方式"].ToString() == "季繳")
                //{
                //    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年第";
                //    //strContent += dtSrc.Rows[i]["開始月"].ToString() + "季至";
                //    //strContent += dtSrc.Rows[i]["結束年"].ToString() + "年第";
                //    //strContent += dtSrc.Rows[i]["結束月"].ToString() + "季止；";
                //    strContent = "";
                //}
                ////年繳
                //if (dtSrc.Rows[i]["繳費方式"].ToString() == "年繳")
                //{
                //    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年 開始；";
                //    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年至";
                //    //strContent += dtSrc.Rows[i]["結束年"].ToString() + "年止；";
                //}
                strContent += dtSrc.Rows[i]["繳費方式"].ToString();
                drNew["繳費方式"] = strContent;
                drNew["奉獻金額"] = dtSrc.Rows[i]["奉獻金額"].ToString();
                dtNew.Rows.Add(drNew);
            }

            //Grid initial
            //沒有需要特別處理的欄位時
            NPOGridView npoGridView = new NPOGridView();
            npoGridView.Source = NPOGridViewDataSource.fromDataTable;
            npoGridView.dataTable = dtNew;
            npoGridView.Keys.Add("Uid");
            npoGridView.DisableColumn.Add("Uid");
            npoGridView.ShowPage = false;
            npoGridView.TableID = "ItemPeriod";

            //-------------------------------------------------------------------------
            NPOGridViewColumn col = new NPOGridViewColumn("Description");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("奉獻項目");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("Amount");
            col.ColumnType = NPOColumnType.CurrecyText;
            col.ColumnName.Add("奉獻金額");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            //col = new NPOGridViewColumn("奉獻周期");
            //col.ColumnType = NPOColumnType.NormalText;
            //col.ColumnName.Add("奉獻內容");
            //npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            //add by geo 20140701
                //下拉選單欄位
                //填 Option 方法 1, 用 List 傳入,設定部份
                List<string> listText = new List<string>();
                List<string> listValue = new List<string>();
                //listText.Add("");
                listText.Add("Annual");
                listText.Add("Quarterly");
                listText.Add("Monthly");
                //listValue.Add("");
                listValue.Add("Annual");
                listValue.Add("Quarterly");
                listValue.Add("Monthly");

                col = new NPOGridViewColumn("Payment Option"); //產生所需欄位及設定 Title
                col.ColumnType = NPOColumnType.Dropdownlist;
                col.ColumnName.Add("繳費方式"); //如果 ColumnName 有設定,那就會與資料欄位勾稽
                col.ControlKeyColumn.Add("uid");
                //col.CaptionWithControl = true; //欄位抬頭是否要有全選 Checkbox, default 為 false

                col.ControlName.Add("ddlPeriod"); //與 ControlKeyColumn 配合組成唯一的 name 屬性(name='chkSelect_xxxx', xxxx 為uid)
                col.ControlId.Add("ddlPeriod"); //與 ControlKeyColumn 配合組成唯一的 id 屬性(id='chkSelect_xxxx', xxxx 為uid)

                //用 List 傳入
                col.ControlOptionText = listText;
                col.ControlOptionValue = listValue;

                npoGridView.Columns.Add(col);


            //col = new NPOGridViewColumn("奉獻周期");
            //col.ColumnType = NPOColumnType.Dropdownlist;
            //col.ControlName.Add("奉獻周期");
            //col.ControlOptionText.Add("年繳");
            //col.ControlOptionText.Add("季繳");
            //col.ControlOptionText.Add("月繳");
            ////
            //col.ControlOptionValue.Add("年繳");
            //col.ControlOptionValue.Add("季繳");
            //col.ControlOptionValue.Add("月繳");
            ////

            //col.ControlId.Add(dtPeriod.Rows.Count.ToString());
            ////col.ControlId.Add("2");
            ////col.ControlId.Add("1");
            ////
            //col.ColumnName.Add("");
            //col.ColumnName.Add("");
            //col.ColumnName.Add("");
            //npoGridView.Columns.Add(col);
            ////-------------------------------------------------------------------------
            //欄位為 Button
            col = new NPOGridViewColumn("");
            col.ColumnType = NPOColumnType.Link;
            col.ControlKeyColumn.Add("Uid");
            col.Link = Util.RedirectByTime("ShoppingCart.aspx", "DelPeriod=");
            col.LinkText = "Delete";
            col.ShowConfirmDialog = true;
            col.ConfirmDialogMsg = "Are you sure to delete this item?";
            npoGridView.Columns.Add(col);

            lblGridPeriod.Text = @"                
                    <span align='center' style='color: Black; font-weight: bold; font-size: medium;'>
                        ※ Recurring Donation
                    </span>
                " + npoGridView.Render() /*+ @"<br><span align='center' style='color: red; font-size: medium;'>
                        1.定期定額捐款扣款方式：捐款金額月繳是每月扣款，季繳為每三個月扣款一次，以此類推。<br><br>
                        2.若需終止定期定額奉獻，請電洽-奉獻服務專線(02)8024-3911 捐款服務組，謝謝您的支持。
                    </span>"*/;
        //}
        //else
        //{
        //    lblGridPeriod.Text = "";
        //}

    }
    //-------------------------------------------------------------------------
    protected void btnContinue_Click(object sender, EventArgs e)
    {
        UpdateCart();
        Response.Redirect("DonateOnlineAll.aspx");
    }
    //-------------------------------------------------------------------------
    protected void btnCheckOut_Click(object sender, EventArgs e)
    {
        UpdateCart();
        Response.Redirect("CheckOut.aspx");
    }
    //-------------------------------------------------------------------------
    //add by geo 20140701
    private void UpdateCart()
    {
        string uid = "";
        string[] strArray;

        foreach (string str in Request.Form.Keys)
        {
            if (str.StartsWith("ddlPeriod_"))
            {
                strArray = str.Split('_');
                uid = strArray[1];
                //繳費方式
                string Status = Util.GetControlValue("ddlPeriod_" + uid);

                for (int i = 0; i < dtPeriod.Rows.Count; i++)
                {
                    if (dtPeriod.Rows[i]["Uid"].ToString() == uid)
                    {
                        dtPeriod.Rows[i]["繳費方式"] = Status;

                        if (Status == "Monthly")
                        {
                            dtPeriod.Rows[i]["開始年"] = Util.GetDBDateTime().Year.ToString();
                            dtPeriod.Rows[i]["開始月"] = Util.GetDBDateTime().Month.ToString();
                            dtPeriod.Rows[i]["結束年"] = "2099";
                            dtPeriod.Rows[i]["結束月"] = "12";
                        }
                        if (Status == "Quarterly")
                        {
                            dtPeriod.Rows[i]["開始年"] = Util.GetDBDateTime().Year.ToString();
                            dtPeriod.Rows[i]["開始月"] = Util.GetDBDateTime().Month.ToString();
                            dtPeriod.Rows[i]["結束年"] = "2099";
                            dtPeriod.Rows[i]["結束月"] = "12";
                        }
                        if (Status == "Annual")
                        {
                            dtPeriod.Rows[i]["開始年"] = Util.GetDBDateTime().Year.ToString();
                            dtPeriod.Rows[i]["開始月"] = "";
                            dtPeriod.Rows[i]["結束年"] = "2099";
                            dtPeriod.Rows[i]["結束月"] = "";
                        }
                    }
                }
            }
        } //end of foreach (string str in Request.Form.Keys)

        Session["ItemPeriod"] = dtPeriod;
    }
}