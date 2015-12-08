using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;

public partial class DonorGroupMgr_FieldSetup : BasePage 
{
    GroupAuthrity ga = null;
    //------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "FieldSetup";
        //權控處理
        AuthrityControl();

        ShowSysMsg();

        if (!IsPostBack)
        {
            HFD_Uid.Value = Util.GetQueryString("Uid");

            LoadDropDownListData();
            LoadFormData();
        }

        SetButton();
    }
    //----------------------------------------------------------------------
    private void SetButton()
    {
    }
    //-------------------------------------------------------------------------
    public void AuthrityControl()
    {
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //欄位定義
        string strSql = "";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        strSql = @"
                    select distinct ConfigName
                    from FieldDetail
                  ";
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        Util.FillDropDownList(ddlConfigName, dt, "ConfigName", "ConfigName", true, "", "");
    }
    //----------------------------------------------------------------------
    private void LoadAllField()
    {
        string strSql = @"
                    select uid, FieldName
                    from FieldMaster
                    where 1=1
                 ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        Util.FillListBox(lstAllField, dt, "FieldName", "uid", false);
    }
    //----------------------------------------------------------------------
    private void LoadDisplayField()
    {
        //已設定的 field
        string strSql = "";
        if (ddlConfigName.SelectedValue != "")
        {
            strSql = @"
                    select 
                    0 as uid, '會籍號碼' as FieldName, 0 as OrderNo 
                    union                    
                    select 
                    0 as uid, '中文姓名' as FieldName, 0 as OrderNo 
                    union
                    select uid, FieldName, OrderNo
                    from FieldDetail
                    where 1=1
                    and ConfigName=@ConfigName
                    order by OrderNo
                 ";
        }
        else
        {
            strSql = @"
                    select 
                    0 as uid, '會籍號碼' as FieldName 
                    union                    
                    select 
                    0 as uid, '中文姓名' as FieldName 
                 ";
        }
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("ConfigName", ddlConfigName.SelectedValue);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        Util.FillListBox(lstDisplayField, dt, "FieldName", "uid", false);
    }
    //----------------------------------------------------------------------
    public void LoadFormData()
    {
        LoadDisplayField();
        LoadAllField();
    }
    //----------------------------------------------------------------------
    protected void btnSelect_Click(object sender, EventArgs e)
    {
        foreach (ListItem item in lstAllField.Items)
        {
            if (item.Selected == true)
            {
                lstDisplayField.Items.Add(item);
            }
        }
    }
    //----------------------------------------------------------------------
    protected void btnRemove_Click(object sender, EventArgs e)
    {
        for (int i = lstDisplayField.Items.Count - 1; i >= 0; i--)
        {
            ListItem item = lstDisplayField.Items[i];
            if (item.Selected == true)
            {
                if (item.Text == "會籍號碼" || item.Text == "中文姓名")
                {
                    continue;
                }
                lstDisplayField.Items.Remove(item);
            }
        }
    }
    //----------------------------------------------------------------------
    protected void ddlConfigName_SelectedIndexChanged(object sender, EventArgs e)
    {
         LoadFormData();
    }
    //----------------------------------------------------------------------
    protected void btnSaveAs_Click(object sender, EventArgs e)
    {
        if (txtNewConfigName.Text.Trim() == "")
        {
            ShowSysMsg("請輸入新名稱!");
            return;
        }
        bool Exist = Util.Exist("FieldDetail", "ConfigName", txtNewConfigName.Text);
        if (Exist == true)
        {
            ShowSysMsg("名稱重複!");
            return;
        }
        if (lstDisplayField.Items.Count == 2)
        {
            ShowSysMsg("欄位尚未設定!");
            return;
        }

        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();

        objNpoDB.BeginTrans();
        try
        {
            int Count = 1;
            foreach (ListItem item in lstDisplayField.Items)
            {
                if (item.Text == "會籍號碼" || item.Text == "中文姓名")
                {
                    continue;
                }
                Save(item, txtNewConfigName.Text, Count);
                Count++;
            }
            objNpoDB.CommitTrans();
        }
        catch (Exception ex)
        {
            objNpoDB.RollbackTrans();
            throw ex; 
        }
        SetSysMsg("另存新名稱完成!");
        Response.Redirect(Util.RedirectByTime("FieldSetup.aspx"));
    }
    //----------------------------------------------------------------------
    private void Save(ListItem item, string ConfigName, int Count)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();
        SetColumnData(list, item, ConfigName, Count);
        string strSql = Util.CreateInsertCommand("FieldDetail", list, dict);
        objNpoDB.ExecuteSQL(strSql, dict);
    }
    //----------------------------------------------------------------------
    private void SetColumnData(List<ColumnData> list, ListItem item, string ConfigName, int Count)
    {
        list.Add(new ColumnData("ConfigName", ConfigName, true, false, false));
        list.Add(new ColumnData("FieldName", item.Text, true, false, false));
        list.Add(new ColumnData("OrderNo", Count, true, true, false));
    }
    //----------------------------------------------------------------------
    protected void btnDelete_Click(object sender, EventArgs e)
    {
        if (ddlConfigName.SelectedValue == "")
        {
            ShowSysMsg("請先選擇命名!");
            return;
        }
        Delete(ddlConfigName.SelectedValue, true);
        SetSysMsg("命名 " + ddlConfigName.SelectedValue + " 刪除資料完成!");
        Response.Redirect(Util.RedirectByTime("FieldSetup.aspx"));
    }
    //----------------------------------------------------------------------
    private void Delete(string ConfigName, bool UseNpoDB)
    {
        string strSql = @"
                         delete from  FieldDetail
                         where ConfigName=@ConfigName
                        ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("ConfigName", ConfigName);
        if (UseNpoDB == true)
        {
            NpoDB.ExecuteSQLS(strSql, dict);
        }
        else
        {
            objNpoDB.ExecuteSQL(strSql, dict);
        }
    }
    //----------------------------------------------------------------------
    protected void btnSave_Click(object sender, EventArgs e)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();

        objNpoDB.BeginTrans();
        try
        {
            int Count = 1;
            Delete(ddlConfigName.SelectedValue, false);
            foreach (ListItem item in lstDisplayField.Items)
            {
                if (item.Text == "會籍號碼" || item.Text == "中文姓名")
                {
                    continue;
                }
                Save(item, ddlConfigName.SelectedValue, Count);
                Count++;
            }
            objNpoDB.CommitTrans();
        }
        catch (Exception ex)
        {
            objNpoDB.RollbackTrans();
            throw ex;
        }
        SetSysMsg("儲存名稱完成!");
        Response.Redirect(Util.RedirectByTime("FieldSetup.aspx"));
    }
    //----------------------------------------------------------------------
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Write(@"<script>self.opener.ReloadData();window.close();</script>");
    }
    //----------------------------------------------------------------------
    protected void btnUp_Click(object sender, EventArgs e)
    {
        int index = lstDisplayField.SelectedIndex;
        if (index == -1)
        {
            return;
        }
        int[] Indices = lstDisplayField.GetSelectedIndices();
        int length = Indices.Length;
        string text;
        string value;
        //如果选择的最小索引是０，表示是最上面的项
        if (Indices[0] == 0)
        {
            return;
        }

        //判断选择多项时是否是连续的项
        if (Indices.Length != 1 && Indices[0] + length - 1 != Indices[length - 1])
        {

            //MessageBox.Show("Page", "请选择连续的项！");
            return;
        }
        //将选中的上面一项未选中的值赋予临时变量
        text = lstDisplayField.Items[Indices[0] - 1].Text;
        value = lstDisplayField.Items[Indices[0] - 1].Value;

        for (int i = 0; i < length; i++)
        {
            lstDisplayField.Items[Indices[i] - 1].Text = lstDisplayField.Items[Indices[i]].Text;
            lstDisplayField.Items[Indices[i] - 1].Value = lstDisplayField.Items[Indices[i]].Value;
            //保证被选中状态
            lstDisplayField.Items[Indices[i] - 1].Selected = true;
            lstDisplayField.Items[Indices[i]].Selected = false;
        }
        //将选中的上面第一条未选中的值赋予到下面
        lstDisplayField.Items[Indices[0] + length - 1].Text = text;
        lstDisplayField.Items[Indices[0] + length - 1].Value = value;
    }
    //----------------------------------------------------------------------
    protected void Button1_Click(object sender, EventArgs e)
    {
        if (lstDisplayField.SelectedIndex == -1)
        {
            return;
        }
        //获得连续选中的项索引
        int[] Indices = lstDisplayField.GetSelectedIndices();
        int length = Indices.Length;

        string text;
        string value;
        //如果选择的是最底下的项

        if (Indices[length - 1] == lstDisplayField.Items.Count - 1)
        {
            return;
        }

        //判断选择多项时是否是连续的项

        if (Indices.Length != 1 && Indices[0] + length - 1 != Indices[length - 1])
        {
            //MessageBox.Show("Page", "请选择连续的项！");
            return;
        }
        //将选中的下面一项未选中的值赋予临时变量
        text = lstDisplayField.Items[Indices[length - 1] + 1].Text;
        value = lstDisplayField.Items[Indices[length - 1] + 1].Value;
        for (int i = length; i > 0; i--)
        {
            lstDisplayField.Items[Indices[i - 1] + 1].Text = lstDisplayField.Items[Indices[i - 1]].Text;
            lstDisplayField.Items[Indices[i - 1] + 1].Value = lstDisplayField.Items[Indices[i - 1]].Value;
            lstDisplayField.Items[Indices[i - 1] + 1].Selected = true;
            lstDisplayField.Items[Indices[i - 1]].Selected = false;
        }
        //将下面第一条未选中的项的值赋予到上
        lstDisplayField.Items[Indices[0]].Text = text;
        lstDisplayField.Items[Indices[0]].Value = value;
    }
    //----------------------------------------------------------------------
}
