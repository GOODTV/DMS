using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;

public partial class DonorMgr_EditDonorGroup : BasePage 
{
    GroupAuthrity ga = null;

    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "GroupMapping";
        //權控處理
        AuthrityControl();

        ShowSysMsg();

        HFD_Uid.Value = Util.GetQueryString("Uid");

        if (!IsPostBack)
        {
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
        //群組類別及項目名稱
        string strSql = @"
                          select uid, GroupClassName
                          from GroupClass gc
                          where DeleteDate is null
                         ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        Util.FillDropDownList(ddlGroupClass, dt, "GroupClassName", "uid", false);
        //------------------------------------------------------------------------
    }
    //----------------------------------------------------------------------
    public void LoadGroupItem(string GroupClassUID)
    {
        //已屬於某群組的 user 表
        string strSql = @"
                    select gi.uid, gi.GroupItemName
                    from GroupItem gi
                    where gi.GroupClassUID=@GroupClassUID and DeleteDate is null
                 ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("GroupClassUID", GroupClassUID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        Util.FillDropDownList(ddlGroupItem, dt, "GroupItemName", "uid", false);
    }
    //----------------------------------------------------------------------
    public void LoadGroupItemUser(string GroupItemUID)
    {
        //已屬於某群組的 user 表
        string strSql = @"
                    select gm.uid, d.Donor_Name
                    from GroupMapping gm
                    left join Donor d on gm.DonorID=d.Donor_Id
                    where gm.GroupItemUID=@GroupItemUID
                 ";
        strSql += " and d.Donor_Id=@DonorID ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("GroupItemUID", GroupItemUID);
        dict.Add("DonorID", HFD_Uid.Value);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        Util.FillListBox(lstGroupItemUser, dt, "Donor_Name", "uid", false);
    }
    //----------------------------------------------------------------------
    public void LoadAvailableUser()
    {
        //已屬於某群組的 user 表
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = @"
                    select d.Donor_Id, d.Donor_Name
                    from Donor d
                    where 1=1
                 ";
        if (txtDonorName.Text != "")
        {
            strSql += " and d.Donor_Name like @DonorName\n";
            dict.Add("DonorName", "%" + txtDonorName.Text + "%");
        }
        strSql += @"
                    and d.Donor_Id not in
                    (
                        select d.Donor_Id
                        from GroupMapping gm
                        left join Donor d on gm.DonorID=d.Donor_Id
                        where gm.GroupItemUID=@GroupItemUID
                    )
                   ";
        strSql += " and d.Donor_Id=@DonorID ";
        dict.Add("GroupItemUID", ddlGroupItem.SelectedValue);
        dict.Add("DonorID", HFD_Uid.Value);
                    
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        Util.FillListBox(lstAvailableUser, dt, "Donor_Name", "Donor_Id", false);
    }
    //----------------------------------------------------------------------
    public void LoadFormData()
    {
        LoadGroupItem(ddlGroupClass.SelectedValue);
        LoadGroupItemUser(ddlGroupItem.SelectedValue);
        //會用到 GroupItem, 所以要放最後面
        LoadAvailableUser();
        //------------------------------------------------------------------------
    }
    //----------------------------------------------------------------------
    protected void ddlGroupClass_SelectedIndexChanged(object sender, EventArgs e)
    {
        LoadGroupItem(ddlGroupClass.SelectedValue);
        LoadGroupItemUser(ddlGroupItem.SelectedValue);
        //會用到 GroupItem, 所以要放最後面
        LoadAvailableUser();
    }
    //----------------------------------------------------------------------
    protected void btnSelect_Click(object sender, EventArgs e)
    {
        string GroupItemUID = ddlGroupItem.SelectedValue;

        foreach (ListItem item in lstAvailableUser.Items)
        {
            if (item.Selected == true)
            {
                AddUser(item.Value, GroupItemUID);
            }
        }
        LoadGroupItemUser(ddlGroupItem.SelectedValue);
        LoadAvailableUser();

        // 2014/7/15 以表單呈現方式，取消這欄位的回傳
        //ScriptManager.RegisterStartupScript(this.Page, this.Page.GetType(), "AddDonor", "self.opener.document.getElementById('lblDonorGroup').value='" + CaseUtil.GetDonorGroup(HFD_Uid.Value) + "';", true);
    }
    //----------------------------------------------------------------------
    private void AddUser(string DonorID, string GroupItemUID)
    {
        List<ColumnData> list = new List<ColumnData>();
        SetColumnData(list, DonorID);
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = Util.CreateInsertCommand("GroupMapping", list, dict);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    //----------------------------------------------------------------------
    private void SetColumnData(List<ColumnData> list, string DonorID)
    {
        list.Add(new ColumnData("GroupItemUid", ddlGroupItem.SelectedValue, true, false, false));
        list.Add(new ColumnData("DonorID", DonorID, true, false, false));
    }
    //----------------------------------------------------------------------
    protected void btnRemove_Click(object sender, EventArgs e)
    {
        foreach (ListItem item in lstGroupItemUser.Items)
        {
            if (item.Selected == true)
            {
                RemoveUser(item.Value);
            }
        }
        LoadAvailableUser();
        LoadGroupItemUser(ddlGroupItem.SelectedValue);

        ScriptManager.RegisterStartupScript(this.Page, this.Page.GetType(), "AddDonor", "self.opener.document.getElementById('lblDonorGroup').value='" + CaseUtil.GetDonorGroup(HFD_Uid.Value) + "';", true);
    }
    //----------------------------------------------------------------------
    private void RemoveUser(string GroupMappingUID)
    {
        string strSql = @"
                          delete GroupMapping
                          where uid=@GroupMappingUID
                        ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("GroupMappingUID", GroupMappingUID);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    //----------------------------------------------------------------------
    protected void btnSearch_Click(object sender, EventArgs e)
    {
        LoadAvailableUser();
    }
    //----------------------------------------------------------------------
    protected void ddlGroupItem_SelectedIndexChanged(object sender, EventArgs e)
    {
        LoadAvailableUser();
        LoadGroupItemUser(ddlGroupItem.SelectedValue);
    }
    //----------------------------------------------------------------------
}
