using System;
using System.Collections.Generic;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonorGroupMgr_GroupItemQry : BasePage
{
    GroupAuthrity ga = null;

    #region NpoGridView 處理換頁相關程式碼
    Button btnNextPage, btnPreviousPage, btnGoPage;
    HiddenField HFD_CurrentPage;

    override protected void OnInit(EventArgs e)
    {
        CreatePageControl();
        base.OnInit(e);
    }
    private void CreatePageControl()
    {
        // Create dynamic controls here.
        btnNextPage = new Button();
        btnNextPage.ID = "btnNextPage";
        Form1.Controls.Add(btnNextPage);
        btnNextPage.Click += new System.EventHandler(btnNextPage_Click);

        btnPreviousPage = new Button();
        btnPreviousPage.ID = "btnPreviousPage";
        Form1.Controls.Add(btnPreviousPage);
        btnPreviousPage.Click += new System.EventHandler(btnPreviousPage_Click);

        btnGoPage = new Button();
        btnGoPage.ID = "btnGoPage";
        Form1.Controls.Add(btnGoPage);
        btnGoPage.Click += new System.EventHandler(btnGoPage_Click);

        HFD_CurrentPage = new HiddenField();
        HFD_CurrentPage.Value = "1";
        HFD_CurrentPage.ID = "HFD_CurrentPage";
        Form1.Controls.Add(HFD_CurrentPage);
    }
    protected void btnPreviousPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.MinusStringNumber(HFD_CurrentPage.Value);
        LoadFormData();
    }
    protected void btnNextPage_Click(object sender, EventArgs e)
    {
        HFD_CurrentPage.Value = Util.AddStringNumber(HFD_CurrentPage.Value);
        LoadFormData();
    }
    protected void btnGoPage_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    #endregion NpoGridView 處理換頁相關程式碼
    //---------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "GroupItem";
        //權控處理
        AuthrityControl();

        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
        if (!IsPostBack)
        {
            LoadDropDownListData();
            if (String.IsNullOrEmpty(Request.Params["GroupItemUid"]))
            {
                lblGridList.Text = "<div align='center'>** 請先輸入查詢條件 **</div>";
                btnExcel.Enabled = false;
            }
            else
            {
                txtGroupItemUid.Text = Request.Params["GroupItemUid"].ToString();
                LoadFormData();
                txtGroupItemUid.Text = "";
            }
        }
    }
    //----------------------------------------------------------------------
    public void AuthrityControl()
    {
        ga = Authrity.GetGroupRight();
        if (Authrity.CheckPageRight(ga.Focus) == false)
        {
            return;
        }
        btnAdd.Visible = ga.AddNew;
        btnQuery.Visible = ga.Query;
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        string strSql = @"
                          select uid, GroupClassName
                          from GroupClass gc
                          where DeleteDate is null
                         ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        Util.FillDropDownList(ddlGroupClass, dt, "GroupClassName", "uid", true);
    }
    //---------------------------------------------------------------------------
    public void LoadFormData()
    {
        // 增加查詢條件顯示
        string strSQLWhere = "";

        string strSql = "";
        strSql = @"
                   select ROW_NUMBER() OVER(ORDER BY gi.uid) as [序號] ,gi.uid,
                   gc.GroupClassName as 群組類別,
                   gi.GroupItemName as 群組代表,
                   gi.Supplement as 備註
                   from GroupItem gi
                   left join GroupClass gc on gi.GroupClassUid=gc.uid
                   where gi.DeleteDate is null
                  ";

        if (ddlGroupClass.SelectedValue != "")
        {
            strSql += "and gi.GroupClassUID=@GroupClassUID\n";
            strSQLWhere += "群組類別：" + ddlGroupClass.SelectedItem.ToString();
        }
        if (txtGroupItemName.Text != "")
        {
            strSql += "and gi.GroupItemName like @GroupItemName\n";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "群組代表：" + txtGroupItemName.Text.Trim();
        }
        if (txtGroupItemUid.Text != "")
        {
            strSql += "and gi.uid = @GroupItemUid\n";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "群組代表編號：" + txtGroupItemUid.Text;
        }
        if (txtSupplement.Text != "")
        {
            strSql += "and gi.Supplement like @Supplement\n";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "備註：" + txtSupplement.Text.Trim();
        }
        if (txtDonorID.Text.Trim() != "")
        {
            strSql += "and gi.uid in (SELECT GroupItemUid FROM GroupMapping where DeleteDate is null and DonorID = @DonorID) \n";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "捐款人編號：" + txtDonorID.Text.Trim();
        }

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("GroupClassUID", ddlGroupClass.SelectedValue);
        dict.Add("GroupItemName", "%" + txtGroupItemName.Text.Trim() + "%");
        dict.Add("Supplement", "%" + txtSupplement.Text.Trim() + "%");
        dict.Add("DonorID", txtDonorID.Text.Trim());
        dict.Add("GroupItemUid", txtGroupItemUid.Text);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        /*
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("uid");
        npoGridView.DisableColumn.Add("uid");
        npoGridView.CurrentPage = Util.String2Number(HFD_CurrentPage.Value);
        npoGridView.EditLink = Util.RedirectByTime("GroupItem_Edit.aspx", "GroupItemUID=");
        lblGridList.Text = npoGridView.Render();
        */

        lblGridList.Text = "";

        if (dt.Rows.Count > 0)
        {

            string strBody = "<table class='table_h' width='100%'>";
            strBody += @"<TR><TH noWrap><SPAN>序號</SPAN></TH>
                            <TH noWrap><SPAN>群組類別</SPAN></TH>
                            <TH noWrap><SPAN>群組代表</SPAN></TH>
                            <TH noWrap><SPAN>備註</SPAN></TH>
                            <TH noWrap><SPAN>成員列表：捐款人編號(捐款人名稱)</SPAN></TH>
                        </TR>";

            foreach (DataRow dr in dt.Rows)
            {

                string strItemUid = dr["uid"].ToString().Trim();

                string strMember = "";
                string strSql2 = @"
					select Donor_Id,Donor_Name,Sex,Title,Donor_Type,Cellular_Phone,Tel_Office,Tel_Office_Loc,Tel_Office_Ext,Email
					,(Case When D.IsAbroad = 'N' then B.Name + C.Name + D.[Address] Else D.[Address] End) as [Address] 
					,Invoice_Type,Remark
                    from GroupMapping GM1 
                    join Donor D on GM1.DonorId = D.Donor_Id and D.DeleteDate is null
			        left Join CODECITY As B on D.City = B.ZipCode 
				    left Join CODECITY As C on D.Area = C.ZipCode
	    			where GM1.DeleteDate is null and GM1.GroupItemUid = @ItemUid
                    order by Donor_Id
                        ";

                Dictionary<string, object> dict2 = new Dictionary<string, object>();
                dict2.Add("ItemUid", strItemUid);
                DataTable dt2 = NpoDB.GetDataTableS(strSql2, dict2);
                if (dt2.Rows.Count > 0)
                {

                    DataRow dr2;
                    strMember = "";
                    for (int j = 0; j < dt2.Rows.Count; j++)
                    {
                        dr2 = dt2.Rows[j];
                        string strMemberTitle = "";
                        string strTel = dr2["Tel_Office_Loc"].ToString().Trim() == "" ? "" : "(" + dr2["Tel_Office_Loc"].ToString().Trim() + ")";
                        strTel += dr2["Tel_Office"].ToString().Trim();
                        strTel += strTel + dr2["Tel_Office_Ext"].ToString().Trim() == "" ? "" : "#" + dr2["Tel_Office_Ext"].ToString().Trim();
                        strMemberTitle += "性別：" + dr2["Sex"].ToString().Trim() + "\n";
                        strMemberTitle += "稱謂：" + dr2["Title"].ToString().Trim() + "\n";
                        strMemberTitle += "身分別：" + dr2["Donor_Type"].ToString().Trim() + "\n";
                        strMemberTitle += "手機：" + dr2["Cellular_Phone"].ToString().Trim() + "\n";
                        strMemberTitle += "電話(日)：" + strTel + "\n";
                        strMemberTitle += "E-Mail：" + dr2["Email"].ToString().Trim() + "\n";
                        strMemberTitle += "通訊地址：" + dr2["Address"].ToString().Trim() + "\n";
                        strMemberTitle += "收據開立：" + dr2["Invoice_Type"].ToString().Trim() + "\n";
                        strMemberTitle += "捐款人備註：\n" + dr2["Remark"].ToString().Trim() + "\n";
                        strMember += (strMember == "" ? "" : ", ") + "<SPAN onclick=\"window.open('" +
                            Util.RedirectByTime("/DonorMgr/DonorInfo_Edit.aspx", "Donor_Id=" + dr2["Donor_Id"].ToString().Trim()) +
                            "','_self','')\" title='" + strMemberTitle + "'><font color='blue'>" + dr2["Donor_Id"].ToString().Trim() + "</font>(" + dr2["Donor_Name"].ToString().Trim() + ")</SPAN>";


                    }

                }

                strBody += "<TR style=\"cursor: pointer; cursor: hand;\" onclick =\"window.open('" + Util.RedirectByTime("GroupItem_Edit.aspx", "GroupItemUID=" + strItemUid) + "','_self','')\">";
                strBody += "<TD noWrap align='right'><SPAN>" + dr["序號"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["群組類別"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["群組代表"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["備註"].ToString() + "</SPAN></TD>";
                strBody += "<TD onclick='window.event.cancelBubble=true;'>" + strMember + "&nbsp;</TD>";
                strBody += "</TR>";
            }

            strBody += "</table>";
            lblGridList.Text = "<font color='blue'>【查詢條件】</font>" + (strSQLWhere == "" ? "全部" : strSQLWhere) + strBody;
            Session["GridList"] = lblGridList.Text;
	    btnExcel.Enabled = true;

        }
        else
        {
            lblGridList.Text = "<div align='center'>** 沒有符合條件的資料 **</div>";
	    btnExcel.Enabled = false;
        }

    }
    //----------------------------------------------------------------------
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Response.Redirect(Util.RedirectByTime("GroupItem_Edit.aspx?Mode=ADD&", ""));
    }
    //----------------------------------------------------------------------
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        LoadFormData();
    }
    //----------------------------------------------------------------------
    protected void btnExcel_Click(object sender, EventArgs e)
    {
        Response.Redirect("GroupItemQry_Excel.aspx");
    }
    //----------------------------------------------------------------------

}
