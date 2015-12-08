using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;

public partial class DonorMgr_Gift_Edit : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Gift_Id");
            LoadDropDownListData();
            Form_DataBind();
        }
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
    }
    //----------------------------------------------------------------------
    //帶入資料
    public void Form_DataBind()
    {
        //****變數宣告****//
        string strSql, uid;
        DataTable dt;

        //****變數設定****//
        uid = HFD_Uid.Value;

        //****設定查詢****//
        strSql = @" select * 
                    from Gift C left join Donor Dr on C.Donor_Id = Dr.Donor_ID  
                        Left Join Act A on C.Act_Id = A.Act_Id 
                        Left Join Dept D on C.Dept_Id = D.DeptID 
                    where Dr.DeleteDate is null and C.Gift_Id='" + uid + "'";

        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);


        //資料異常
        if (dt.Rows.Count <= 0)
            //todo : add Default.aspx page
            Response.Redirect("ContributeList.aspx");

        DataRow dr = dt.Rows[0];

        HFD_Donor_Id.Value = dr["Donor_Id"].ToString().Trim();
        //捐款人
        tbxDonor_Name.Text = dr["Donor_Name"].ToString().Trim();
        //類別
        //tbxCategory.Text = dr["Category"].ToString().Trim();
        //身份別
        tbxDonor_Type.Text = dr["Donor_Type"].ToString().Trim();
        //收據地址
        //tbxAddress.Text = dr["通訊地址"].ToString().Trim();
        //身分證/統編
        tbxIDNo.Text = dr["IDNo"].ToString().Trim();
        //贈送日期
        tbxGift_Date.Text = DateTime.Parse(dr["Gift_Date"].ToString()).ToString("yyyy/MM/dd");
        //捐款備註
        tbxComment.Text = dr["Comment"].ToString().Trim();

        //載入捐贈內容
        string HLink, LinkParam, LinkTarget, DataLink, DataParam, DataTarget, strSql2, uid2 = "";
        DataTable dt2;
        uid2 = HFD_Uid.Value;
        strSql2 = " select CD.Ser_No as [Ser_No], ROW_NUMBER() OVER(ORDER BY CD.Ser_No) as [序號], \n";
        strSql2 += " L.Linked2_Name as [物品名稱], CD.Goods_Qty as [數量] from GiftData CD right join Linked2 L on CD.Goods_Id = L.Ser_No\n";
        strSql2 += " where Gift_Id='" + uid2 + "'\n";
        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        dt2 = NpoDB.GetDataTableS(strSql2, dict2);

        HLink = "GiftData_Edit.aspx?Ser_No=";
        LinkParam = "Ser_No";
        LinkTarget = "_blank";
        DataLink = "GiftData_Delete.aspx?Ser_No=";
        DataParam = "Ser_No";
        DataTarget = "_self";

        lblGridList.Text = DataLinkGrids(dt2, HLink, LinkParam, LinkTarget, DataLink, DataParam, DataTarget);
    }
    //增加選項欄位
    public string DataLinkGrids(DataTable dt, string HLink, string LinkParam, string LinkTarget, string DataLink, string DataParam, string DataTarget)
    {
        int i;
        StringBuilder sb = new StringBuilder();
        sb.AppendLine(@"<table width='100%' align='center' class='table_h'>");
        sb.AppendLine(@"<tr>");
        i = 0;
        foreach (DataColumn dc in dt.Columns)
        {
            if (i < 1)
            {
                i++;
                continue;

            }
            sb.AppendLine(@"<th>" + dc.ColumnName + "</th>");
        }
        //sb.AppendLine(@"<th>選項</th>");
        sb.AppendLine(@"<th>編輯</th>");
        sb.AppendLine(@"<th>刪除</th>");
        sb.AppendLine("</tr>");

        foreach (DataRow dr in dt.Rows)
        {
            i = 0;
            sb.AppendLine(@"<tr>");
            foreach (DataColumn dc in dt.Columns)
            {
                if (i < 1)
                {
                    i++;
                    continue;
                }
                else
                {
                    //顯示資料
                    sb.AppendLine("<td>" + dr[dc.ColumnName] + "</td>");
                }
                i++;
            }
            //sb.AppendLine(@"<td><a href='" + HLink + dr[LinkParam] + @"' target='" + LinkTarget + @"'><img src='../images/save.gif' border=0 align='absmiddle' alt='修改'>修改</a><a href='JavaScript:if(confirm(""是否確定要刪除 ?"")){window.location.href=""" + DataLink + dr[DataParam] + @""";}' target='" + DataTarget + @"'><img src='../images/delete2.gif' align='absmiddle' border=0 alt='刪除'></a></td>");
            sb.AppendLine(@"<td><a href onclick=""window.open('" + HLink + dr[LinkParam] + @"','GiftData_Edit','status=no,scrollbars=no,top=100,left=120,width=500,height=250')""><img src='../images/save.gif' border=0 align='absmiddle' alt='編輯'>編輯</a></td>");
            //sb.AppendLine(@"<td><a href='" + HLink + dr[LinkParam] + @"' target='" + LinkTarget + @"'><img src='../images/save.gif' border=0 align='absmiddle' alt='修改'>修改</a>");
            sb.AppendLine(@"<td><a href='JavaScript:if(confirm(""是否確定要刪除 ?"")){window.location.href=""" + DataLink + dr[DataParam] + @""";}' target='" + DataTarget + @"'><img src='../images/delete2.gif' align='absmiddle' border=0 alt='刪除'></a></td>");
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</table>");
        return sb.ToString();
    }
    //---------------------------------------------------------------------
    protected void btnEdit_Click(object sender, EventArgs e)
    {
        bool Ckeck_GoodID = true;

            string Goods_Message = "";
            if (tbxGifts_Name_1.Text != "")
            {
                Ckeck_GoodID = Get_Goods(tbxGifts_Name_1.Text);
                Goods_Message += tbxGifts_Name_1.Text + " ";
            }
            if (tbxGifts_Name_2.Text != "")
            {
                Ckeck_GoodID = Get_Goods(tbxGifts_Name_2.Text);
                Goods_Message += tbxGifts_Name_2.Text + " ";
            }
            if (tbxGifts_Name_3.Text != "")
            {
                Ckeck_GoodID = Get_Goods(tbxGifts_Name_3.Text);
                Goods_Message += tbxGifts_Name_3.Text + " ";
            }
            if (tbxGifts_Name_4.Text != "")
            {
                Ckeck_GoodID = Get_Goods(tbxGifts_Name_4.Text);
                Goods_Message += tbxGifts_Name_4.Text + " ";
            }
            if (tbxGifts_Name_5.Text != "")
            {
                Ckeck_GoodID = Get_Goods(tbxGifts_Name_5.Text);
                Goods_Message += tbxGifts_Name_5.Text + " ";
            }


            if (Ckeck_GoodID == false)
            {
                ShowSysMsg("您輸入的『 " + Goods_Message + " 』 該物品代號不存在");
                return;
            }
            else
            {
                bool flag = false;
                string Donor_Id = HFD_Donor_Id.Value;
                try
                {
                    Gift_Edit();

                    //******新增GiftData的資料******//
                    //--------------------------------------
                    //Gifts_1
                    if (tbxGifts_Name_1.Text != "")
                    {
                        GiftData_AddNew(Donor_Id, HFD_Gifts_Id_1.Value, tbxGifts_Name_1.Text.Trim(), tbxGifts_Qty_1.Text.Trim(), tbxGifts_Comment_1.Text.Trim());
                    }
                    //--------------------------------------
                    //Gifts_2
                    if (tbxGifts_Name_2.Text != "")
                    {
                        GiftData_AddNew(Donor_Id, HFD_Gifts_Id_2.Value, tbxGifts_Name_2.Text.Trim(), tbxGifts_Qty_2.Text.Trim(), tbxGifts_Comment_2.Text.Trim());
                    }
                    //--------------------------------------
                    //Gifts_3
                    if (tbxGifts_Name_3.Text != "")
                    {
                        GiftData_AddNew(Donor_Id, HFD_Gifts_Id_3.Value, tbxGifts_Name_3.Text.Trim(), tbxGifts_Qty_3.Text.Trim(), tbxGifts_Comment_3.Text.Trim());
                    }
                    //--------------------------------------
                    //Gifts_4
                    if (tbxGifts_Name_4.Text != "")
                    {
                        GiftData_AddNew(Donor_Id, HFD_Gifts_Id_4.Value, tbxGifts_Name_4.Text.Trim(), tbxGifts_Qty_4.Text.Trim(), tbxGifts_Comment_4.Text.Trim());
                    }
                    //--------------------------------------
                    //Gifts_5
                    if (tbxGifts_Name_5.Text != "")
                    {
                        GiftData_AddNew(Donor_Id, HFD_Gifts_Id_5.Value, tbxGifts_Name_5.Text.Trim(), tbxGifts_Qty_5.Text.Trim(), tbxGifts_Comment_5.Text.Trim());
                    }
                    flag = true;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                if (flag == true)
                {
                    SetSysMsg("資料修改成功！");

                    //if (Session["cType"] == "ContributeDataList")
                    //{
                        Response.Redirect(Util.RedirectByTime("../DonorMgr/Gift_Detail.aspx", "Gift_Id=" + HFD_Uid.Value));
                    //}
                    //else
                    //{
                        //Response.Redirect(Util.RedirectByTime("../DonorMgr/Gift_Detail.aspx"));
                    //}
                }

            }
    }
    protected void btnDel_Click(object sender, EventArgs e)
    {
        //刪除Gift表的資料
        string strSql2 = "delete from Gift where Gift_Id=@Gift_Id";
        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        dict2.Add("Gift_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql2, dict2);

        //刪除GiftData表的資料
        string strSql4 = "delete from GiftData where Gift_Id=@Gift_Id";
        Dictionary<string, object> dict4 = new Dictionary<string, object>();
        dict4.Add("Gift_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql4, dict4);


        SetSysMsg("公關贈品紀錄刪除成功！");
        if (Session["cType"] == "ContributeList")
        {
            Response.Redirect(Util.RedirectByTime("ContributeList.aspx"));
        }
        else if (Session["cType"] == "ContributeDataList")
        {
            Response.Redirect(Util.RedirectByTime("../ContributeMgr/ContributeDataList.aspx", "Donor_Id=" + HFD_Donor_Id.Value));
        }
        else
        {
            Response.Redirect(Util.RedirectByTime("../ContributeMgr/ContributeDataList.aspx"));
        }
    }
    protected void btnExit_Click(object sender, EventArgs e)
    {
        if (Session["cType"] == "ContributeDataList")
        {
            Response.Redirect(Util.RedirectByTime("../DonorMgr/Gift_Detail.aspx", "Gift_Id=" + HFD_Uid.Value));
        }
        else
        {
            Response.Redirect(Util.RedirectByTime("../DonorMgr/Gift_Detail.aspx"));
        }
    }
    public void Gift_Edit()
    {
        //****變數宣告****//
        Dictionary<string, object> dict = new Dictionary<string, object>();

        //****設定SQL指令****//
        string strSql = " update Gift set ";

        strSql += "  Gift_Date = @Gift_Date";
        strSql += ", Comment = @Comment";
        strSql += " where Gift_Id = @Gift_Id";

        dict.Add("Gift_Date", tbxGift_Date.Text.Trim());
        dict.Add("Comment", tbxComment.Text.Trim());

        dict.Add("Gift_Id", HFD_Uid.Value);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public void GiftData_AddNew(string Donor_Id, string Goods_Id, string Goods_Name, string Goods_Qty, string Goods_Comment)
    {
        string strSql = "insert into GiftData\n";
        strSql += "( Gift_Id, Donor_Id, Goods_Id, Goods_Name, Goods_Qty, \n";
        strSql += "  Goods_Comment, Create_Date, Create_User, \n";
        strSql += " Create_IP) values\n";

        strSql += "( @Gift_Id,@Donor_Id,@Goods_Id,@Goods_Name,@Goods_Qty, \n";
        strSql += " @Goods_Comment,@Create_Date, @Create_User, \n";
        strSql += " @Create_IP)";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Gift_Id", HFD_Uid.Value);
        dict.Add("Donor_Id", Donor_Id);
        dict.Add("Goods_Id", Goods_Id);
        dict.Add("Goods_Name", Goods_Name);
        if (string.IsNullOrEmpty(Goods_Qty) || Goods_Qty == "")
            dict.Add("Goods_Qty", "1");
        else
            dict.Add("Goods_Qty", Goods_Qty);
        //dict.Add("Goods_Unit", "");
        //dict.Add("Goods_Amt", "");
        //dict.Add("Goods_DueDate", null);
        dict.Add("Goods_Comment", Goods_Comment);
        //dict.Add("Contribute_IsStock", "N");
        //dict.Add("Export", "N");
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public bool Get_Goods(string strGiftName)
    {
        string strSql = "Select * From Linked2 Where Linked2_Name=@Goods_Name ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Goods_Name", strGiftName);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count == 0)
        {
            return false;
        }
        else
        {
            return true;
        }
    }
}