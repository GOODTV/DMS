using System;
using System.Data;
using System.Configuration;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Text;
/*
 * 群組名稱設定
 * 資料庫: CaseCode
 * List頁面: CaseCode.aspx
 * 新增頁面: CaseCode.aspx 
 * 修改頁面: CaseCode_Edit.aspx
 */
public partial class SysMgr_CaseCode : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "CaseCode";
        //Util.ShowSysMsgWithScript("");
        //權控處理
        AuthrityControl();
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            Inital();
            Case_DataBind();
            CaseItem_DataBind(HFD_GroupName.Value);
            tbxNewCaseID.Text = GetNewMaxLinked2_Seq(HFD_Uid.Value, tbxCaseFolder.Text);
        }
    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_AddNew", btnNewCaseItems);
        Authrity.CheckButtonRight("_Delete", btnDelCaseFolder);
    }
    //---------------------------------------------------------------------------
    protected void btnDelCaseFolder_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            //****變數宣告****//
            Dictionary<string, object> dict = new Dictionary<string, object>();

            //****設定SQL指令****//
            string strSql = " delete From CaseCode ";
            strSql += " where GroupName=@GroupName";

            dict.Add("GroupName", tbxCaseFolder.Text);

            //****執行語法****//
            NpoDB.ExecuteSQLS(strSql, dict);

            flag = true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        if (flag == true)
        {
            SetSysMsg("主檔代碼刪除成功!");
            Response.Redirect(Util.RedirectByTime("CaseCode.aspx"));
        }
    }
    //---------------------------------------------------------------------------
    protected void btnNewCaseItems_Click(object sender, EventArgs e)
    {
        string strRet = CheckFieldMustFillBasic();
        strRet += CheckData();
        if (strRet != "")
        {
            ShowSysMsg(strRet);
            return;
        }
        else
        {
            bool flag = false;
            try
            {
                string strSql = "insert into  CaseCode\n";
                strSql += "( CaseID,CaseName, GroupName,CodeType) values\n";
                strSql += "(@CaseID,@CaseName,@GroupName,@CodeType)";
                strSql += "\n";
                strSql += "select @@IDENTITY";

                Dictionary<string, object> dict = new Dictionary<string, object>();

                dict.Add("CaseID", tbxNewCaseID.Text);
                dict.Add("CaseName", tbxNewCaseItems.Text);
                dict.Add("GroupName", HFD_GroupName.Value);
                dict.Add("CodeType", tbxCodeType.Text);
                NpoDB.ExecuteSQLS(strSql, dict);

                flag = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (flag == true)
            {
                SetSysMsg("新增代碼選項成功！");
                Response.Redirect(Util.RedirectByTime("CaseCode.aspx", "Uid=" + HFD_Uid.Value));
            }
        }
    }
    //---------------------------------------------------------------------------
    public void Inital()
    {
        //tbxCaseFolder
        HFD_Uid.Value = Util.GetQueryString("Uid");
        if (HFD_Uid.Value == "")
        {
            HFD_Uid.Value = GetMin_CaseUid();
        }
        string strSql = @"select CodeType as [代碼], GroupName as [代碼名稱] from CaseCode where Uid ='" + HFD_Uid.Value + "'";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);
        DataRow dr = dt.Rows[0];
        tbxCodeType.Text = dr["代碼"].ToString();
        tbxCaseFolder.Text = dr["代碼名稱"].ToString();
        HFD_GroupName.Value = dr["代碼名稱"].ToString();
    }
    //---------------------------------------------------------------------------
    public void Case_DataBind()
    {
        string HLink, LinkParam, LinkTarget, strSql = "";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;

        strSql = @"Select Uid, CodeType as 代碼, GroupName as 代碼名稱 From CaseCode Where Uid In (Select Min(Uid) From CaseCode Group By GroupName) order by Uid";
        dt = NpoDB.GetDataTableS(strSql, dict);

        HLink = "CaseCode.aspx?Uid=";
        LinkParam = "Uid";
        LinkTarget = "_self";

        Grid_List1.Text = DataLinkGrids(dt, HLink, LinkParam, LinkTarget);
    }
    //---------------------------------------------------------------------------
    public void CaseItem_DataBind(string strGroupName)
    {
        string HLink, LinkParam, LinkParam2, DataLink, DataParam, DataTarget, strSql = "";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;

        strSql = @"select Uid,CaseID,CaseName as 代碼選項,CaseID as 排序 from CaseCode where GroupName ='" + strGroupName + "' order by 排序";
        dt = NpoDB.GetDataTableS(strSql, dict);

        HLink = "CaseCode_Edit.aspx?Uid=";
        LinkParam = "Uid";
        LinkParam2 = "Uid";
        DataLink = "CaseCode_Delete.aspx?Uid=";
        DataParam = "Uid";
        DataTarget = "_self";

        Grid_List2.Text = DataLinkGrids2(dt, HLink, LinkParam, LinkParam2, DataLink, DataParam, DataTarget);
    }
    //---------------------------------------------------------------------------
    //多出選項欄可以填入btn動作
    public string DataLinkGrids(DataTable dt, string HLink, string LinkParam, string LinkTarget)
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
                if (i == 1)
                {
                    sb.AppendLine(@"<td><img border='0' src='../images/folder_min.gif' width='16' height='16' align='absmiddle'><a href='" + HLink + dr[LinkParam] + @"' target='" + LinkTarget + @"'>" + dr[dc.ColumnName] + "</a></td>");
                }
                else
                {
                    //顯示資料
                    sb.AppendLine("<td>" + dr[dc.ColumnName] + "</td>");
                }
                i++;
            }
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</table>");
        return sb.ToString();
    }
    public string DataLinkGrids2(DataTable dt, string HLink, string LinkParam, string LinkParam2, string DataLink, string DataParam, string DataTarget)
    {
        int i;
        StringBuilder sb = new StringBuilder();
        sb.AppendLine(@"<table width='100%' align='center' class='table_h'>");
        sb.AppendLine(@"<tr>");
        i = 0;
        foreach (DataColumn dc in dt.Columns)
        {
            if (i < 2)
            {
                i++;
                continue;

            }
            sb.AppendLine(@"<th>" + dc.ColumnName + "</th>");
        }
        sb.AppendLine(@"<th>選項</th>");
        sb.AppendLine("</tr>");

        foreach (DataRow dr in dt.Rows)
        {
            i = 0;
            sb.AppendLine(@"<tr>");
            foreach (DataColumn dc in dt.Columns)
            {
                if (i < 2)
                {
                    i++;
                    continue;
                }
                if (i == 2)
                {
                    sb.AppendLine(@"<td><a href onclick=""window.open('" + HLink + dr[LinkParam] + @"','CaseCode_Edit','scrollbars=no,status=no,toolbar=no,top=100,left=120,width=580,height=200')"">" + dr[dc.ColumnName] + "</a></td>");
                }
                else
                {
                    //顯示資料
                    sb.AppendLine("<td>" + dr[dc.ColumnName] + "</td>");
                }
                i++;
            }
            sb.AppendLine(@"<td><a href='JavaScript:if(confirm(""是否確定要刪除 ?"")){window.location.href=""" + DataLink + dr[DataParam]  + @""";}' target='" + DataTarget + @"'><img src='../images/delete2.gif' align='absmiddle' border=0 alt='刪除'></a></td>");
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</table>");
        return sb.ToString();
    }
    //---------------------------------------------------------------------------
    //檢核基本資料必填欄位
    private string CheckFieldMustFillBasic()
    {
        string strRet = "";

        if (tbxNewCaseItems.Text.Trim() == "")
        {
            strRet += "選項 ";
        }
        if (tbxNewCaseID.Text.Trim() == "")
        {
            strRet += "排序 ";
        }
        if (tbxNewCaseItems.Text.Trim() == "" || tbxNewCaseID.Text.Trim() == "")
        {
            strRet += "欄位不可為空白！";
        }
        return strRet;
    }
    //檢核欄位有沒有重複
    private string CheckData()
    {
        string strRet = "";
        string strSql = "select * from CaseCode where CaseName = @CaseName and GroupName = @GroupName";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("CaseName",tbxNewCaseItems.Text);
        dict.Add("GroupName",HFD_GroupName.Value);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count > 0)
        {
            strRet += "選項 ";
        }
        strSql = "select * from CaseCode where CaseID = @CaseID and GroupName = @GroupName";
        dict.Clear();
        dict.Add("CaseID", tbxNewCaseID.Text);
        dict.Add("GroupName", HFD_GroupName.Value);
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count > 0)
        {
            strRet += "排序 ";
        }
        if (strRet!="")
        {
            strRet += "欄位不可重複！";
        }
        return strRet;
    }
    //抓取CaseCode中該分類的最大排序
    public string GetNewMaxLinked2_Seq(string strCaseID,string strCasefolder)
    {
        string strRet = "";

        string strSql = @"select isnull(MAX(CaseID),'') as CaseID from CaseCode ";
        strSql += " where GroupName='" + strCasefolder + "'";
        DataTable dt = NpoDB.GetDataTableS(strSql, null);
        if (dt.Rows.Count > 0)
        {
            string CaseID = dt.Rows[0]["CaseID"].ToString();
            if (CaseID == "")
            {
                strRet = "1";
            }
            else
            {
                strRet = (Convert.ToInt32(CaseID) + 1).ToString("0");
            }
        }

        return strRet;
    }
    public string GetMin_CaseUid()
    {
        string strRet = "";

        string strSql = @"select isnull(Min(Uid),'') as Uid from CaseCode ";
        DataTable dt = NpoDB.GetDataTableS(strSql, null);
        if (dt.Rows.Count > 0)
        {
            string Uid = dt.Rows[0]["Uid"].ToString();
            strRet = Convert.ToInt32(Uid).ToString();
        }
        else
        {
            Response.Redirect("Default.aspx");
        }
        return strRet;
    }
}