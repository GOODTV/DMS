﻿using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class DonorGroupMgr_GroupClass_Edit : BasePage 
{
    GroupAuthrity ga = null;

    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "GroupClass";
        //權控處理
        AuthrityControl();

        //取得模式, 新增("ADD")或修改("")
        HFD_Mode.Value = Util.GetQueryString("Mode");

        if (HFD_Mode.Value == "ADD")
        {
            litTitle.Text = "新增群組類別資料";
        }
        else
        {
            litTitle.Text = "修改群組類別資料";
        }

        if (!IsPostBack)
        {
            HFD_GroupClassUID.Value = Util.GetQueryString("GroupClassUID");
            LoadDropDownListData();
            LoadFormData();
        }
        SetButton();
    }
    //----------------------------------------------------------------------
    private void SetButton()
    {
        //如果是新增,要做一些初始畫的動作
        if (HFD_Mode.Value == "ADD")
        {
            //新增時, 修改及刪除鍵沒動作
            btnUpdate.Visible = false;
            btnDelete.Visible = false;
        }
        else
        {
            //修改時, 新增鍵沒動作
            btnAdd.Visible = false;
        }
    }
    //-------------------------------------------------------------------------
    public void AuthrityControl()
    {
        ga = Authrity.GetGroupRight();
        if (Authrity.CheckPageRight(ga.Focus) == false)
        {
            return;
        }
        btnAdd.Visible = ga.AddNew;
        btnUpdate.Visible = ga.Update;
        btnDelete.Visible = ga.Delete;
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
    }
    //----------------------------------------------------------------------
    public void LoadFormData()
    {
        string strSql = "";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = null;
        DataRow dr = null;

        if (HFD_Mode.Value != "ADD")
        {
            strSql = @"
                        select * from
                        GroupClass
                        where uid=@uid
                      ";
            dict.Add("uid", HFD_GroupClassUID.Value);
            dt = NpoDB.GetDataTableS(strSql, dict);
            //資料異常
            if (dt.Rows.Count == 0)
            {
                Response.Redirect("GroupClassQry.aspx");
            }
            dr = dt.Rows[0];
        }

        //群組類別名稱
        txtGroupClassName.Text = (HFD_Mode.Value == "ADD") ? "" : dr["GroupClassName"].ToString();
        //備註
        txtSupplement.Text = (HFD_Mode.Value == "ADD") ? "" : dr["Supplement"].ToString();
    }
    //----------------------------------------------------------------------
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();
        SetColumnData(list);
        string strSql = Util.CreateInsertCommand("GroupClass", list, dict);
        NpoDB.ExecuteSQLS(strSql, dict);
        Session["Msg"] = "新增資料成功";
        Response.Redirect(Util.RedirectByTime("GroupClassQry.aspx"));
    }
    //-------------------------------------------------------------------------------------------------------------
    protected void btnUpdate_Click(object sender, EventArgs e)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        List<ColumnData> list = new List<ColumnData>();
        SetColumnData(list);
        string strSql = Util.CreateUpdateCommand("GroupClass", list, dict);
        NpoDB.ExecuteSQLS(strSql, dict);
        Session["Msg"] = "修改資料成功";
        Response.Redirect(Util.RedirectByTime("GroupClassQry.aspx"));
    }
    //----------------------------------------------------------------------
    private void SetColumnData(List<ColumnData> list)
    {
        //修改時作為 where 的條件值
        list.Add(new ColumnData("Uid", HFD_GroupClassUID.Value, false, false, true));

        list.Add(new ColumnData("GroupClassName", txtGroupClassName.Text, true, true, false));
        list.Add(new ColumnData("Supplement", txtSupplement.Text, true, true, false));

        if (HFD_Mode.Value == "ADD")
        {
            //新增日期
            list.Add(new ColumnData("InsertDate", Util.GetToday(DateType.yyyyMMddHHmmss), true, true, false));
            list.Add(new ColumnData("InsertID", SessionInfo.UserID, true, true, false));
            //修改日期
            list.Add(new ColumnData("UpdateDate", Util.GetToday(DateType.yyyyMMddHHmmss), true, true, false));
            list.Add(new ColumnData("UpdateID", SessionInfo.UserID, true, true, false));
        }
        else
        {
            //修改日期
            list.Add(new ColumnData("UpdateDate", Util.GetToday(DateType.yyyyMMddHHmmss), true, true, false));
            list.Add(new ColumnData("UpdateID", SessionInfo.UserID, true, true, false));
        }
    }
    //-------------------------------------------------------------------------------------------------------------
    protected void btnExit_Click(object sender, EventArgs e)
    {
        Response.Redirect("GroupClassQry.aspx?rand=" + DateTime.Now.Ticks.ToString());
    }
    //----------------------------------------------------------------------
    protected void btnDelete_Click(object sender, EventArgs e)
    {
        //TODO:是否要檢查
        //if (ClassHasItem() == true)
        //{
        //    ShowSysMsg("已有群組項目設定, 無法刪除!");
        //    return;
        //}
        //刪除群組類別
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = @"
                           update GroupClass set
                           DeleteID=@DeleteID,
                           DeleteDate=GetDate()
                           where uid = @uid
                         ";
        dict.Add("uid", HFD_GroupClassUID.Value);
        dict.Add("DeleteID", SessionInfo.UserID);
        NpoDB.ExecuteSQLS(strSql, dict);
        //20140516 新增 by Ian_Kao 刪除群組項目
        dict.Clear();
        strSql = @"
                    update GroupItem set
                    DeleteID=@DeleteID,
                    DeleteDate=GetDate()
                    where GroupClassUid = @GroupClassUid
                    select Uid 
                    from GroupItem 
                    where GroupClassUid = @GroupClassUid and DeleteID=@DeleteID and DeleteDate=GetDate()
                  ";
        dict.Add("GroupClassUid", HFD_GroupClassUID.Value);
        dict.Add("DeleteID", SessionInfo.UserID);
        string GroupItemUid = NpoDB.GetScalarS(strSql, dict);
        //20140516 新增 by Ian_Kao 刪除群組配對奉獻人
        dict.Clear();
        strSql = @"
                    update GroupMapping set
                    DeleteID=@DeleteID,
                    DeleteDate=GetDate()
                    where GroupItemUid = @GroupItemUid
                  ";
        dict.Add("GroupItemUid", GroupItemUid);
        dict.Add("DeleteID", SessionInfo.UserID);
        NpoDB.ExecuteSQLS(strSql, dict);

        Session["Msg"] = "刪除資料成功";
        Response.Redirect(Util.RedirectByTime("GroupClassQry.aspx"));
    }
    //----------------------------------------------------------------------
    private bool ClassHasItem()
    {
//        string strSql = @"
//                          select top 1 uid
//                          from GroupItem gi
//                          where gi.uid=@uid";
//        Dictionary<string, object> dict = new Dictionary<string, object>();
//        dict.Add("uid", HFD_GroupClassUID.Value);
//        objNpoDB.ExecuteSQL(strSql, dict);

//        strSql = "delete from  AdminRight where GroupID=@uid";
//        objNpoDB.ExecuteSQL(strSql, dict);

        return true;
    }
    //----------------------------------------------------------------------
}
