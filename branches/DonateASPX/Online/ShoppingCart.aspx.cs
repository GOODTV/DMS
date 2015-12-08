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

        if (!IsPostBack)
        {
            if (Session["ItemOnce"] != null)
            {
                dtOnce = (DataTable)Session["ItemOnce"];
            }
            if (Session["ItemPeriod"] != null)
            {
                dtPeriod = (DataTable)Session["ItemPeriod"];
            }

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
            }

            //if (dtOnce.Rows.Count > 0)
            //{
                ShowCartOnce(dtOnce);
            //}

            //if (dtPeriod.Rows.Count > 0)
            //{
                ShowCartPeriod(dtPeriod);
            //}

            if (dtOnce.Rows.Count == 0 && dtPeriod.Rows.Count == 0)
            {
                btnCheckOut.Visible = false;
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
		lblCart.Text = "您的奉獻清單有: 單筆(" + iOnce.ToString() + ")項；定期定額(" + iPeriod.ToString() + ")項";
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
                drSum["奉獻項目"] = "奉獻小計NT$";
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
            NPOGridViewColumn col = new NPOGridViewColumn("用途");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("奉獻項目");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("金額");
            col.ColumnType = NPOColumnType.CurrecyText;
            col.ColumnName.Add("奉獻金額");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            //欄位為 Button
            col = new NPOGridViewColumn("");
            col.ColumnType = NPOColumnType.Link;
            col.ControlKeyColumn.Add("Uid");
            col.Link = Util.RedirectByTime("ShoppingCart.aspx", "DelOnce=");
            col.LinkText="刪除";
            col.ShowConfirmDialog = true;
            col.ConfirmDialogMsg = "確定要刪除此筆奉獻項目？";
            npoGridView.Columns.Add(col);

            lblGridOnce.Text = @"                
                    <span align='center' style='color: blue; font-weight: bold;'>
                        ※　單　筆　奉　獻
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
            dtNew.Columns.Add("奉獻內容");
            dtNew.Columns.Add("奉獻金額");

            for (int i = 0; i < dtSrc.Rows.Count; i++)
            {
                DataRow drNew = dtNew.NewRow();
                drNew["Uid"] = dtSrc.Rows[i]["Uid"].ToString();
                drNew["奉獻項目"] = dtSrc.Rows[i]["奉獻項目"].ToString();
                string strContent = "";
                //月繳
                if (dtSrc.Rows[i]["繳費方式"].ToString() == "月繳")
                {
                    strContent = dtSrc.Rows[i]["開始年"].ToString() + "年";
                    strContent += dtSrc.Rows[i]["開始月"].ToString() + "月至";
                    strContent += dtSrc.Rows[i]["結束年"].ToString() + "年";
                    strContent += dtSrc.Rows[i]["結束月"].ToString() + "月止；";
                }
                //季繳
                if (dtSrc.Rows[i]["繳費方式"].ToString() == "季繳")
                {
                    //strContent = dtSrc.Rows[i]["開始年"].ToString() + "年第";
                    //strContent += dtSrc.Rows[i]["開始月"].ToString() + "季至";
                    //strContent += dtSrc.Rows[i]["結束年"].ToString() + "年第";
                    //strContent += dtSrc.Rows[i]["結束月"].ToString() + "季止；";
                    strContent = "";
                }
                //年繳
                if (dtSrc.Rows[i]["繳費方式"].ToString() == "年繳")
                {
                    strContent = dtSrc.Rows[i]["開始年"].ToString() + "年至";
                    strContent += dtSrc.Rows[i]["結束年"].ToString() + "年止；";
                }
                strContent += dtSrc.Rows[i]["繳費方式"].ToString();
                drNew["奉獻內容"] = strContent;
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
            NPOGridViewColumn col = new NPOGridViewColumn("用途");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("奉獻項目");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("周期");
            col.ColumnType = NPOColumnType.NormalText;
            col.ColumnName.Add("奉獻內容");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            col = new NPOGridViewColumn("金額");
            col.ColumnType = NPOColumnType.CurrecyText;
            col.ColumnName.Add("奉獻金額");
            npoGridView.Columns.Add(col);
            //-------------------------------------------------------------------------
            //欄位為 Button
            col = new NPOGridViewColumn("");
            col.ColumnType = NPOColumnType.Link;
            col.ControlKeyColumn.Add("Uid");
            col.Link = Util.RedirectByTime("ShoppingCart.aspx", "DelPeriod=");
            col.LinkText = "刪除";
            col.ShowConfirmDialog = true;
            col.ConfirmDialogMsg = "確定要刪除此筆奉獻項目？";
            npoGridView.Columns.Add(col);

            lblGridPeriod.Text = @"                
                    <span align='center' style='color: Blue; font-weight: bold;'>
                        ※　定　期　定　額　奉　獻
                    </span>
                " + npoGridView.Render();
        //}
        //else
        //{
        //    lblGridPeriod.Text = "";
        //}

    }
    //-------------------------------------------------------------------------
    protected void btnContinue_Click(object sender, EventArgs e)
    {
        Response.Redirect("DonateOnlineAll.aspx");
    }
    protected void btnCheckOut_Click(object sender, EventArgs e)
    {
        Response.Redirect("CheckOut.aspx");
    }
}