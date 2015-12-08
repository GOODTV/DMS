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
 * 資料庫: Linked/Linked2
 * 相關程式:
 * List頁面: LinkedList.aspx
 * 新增頁面: LinkedList_Add.aspx 
 * 修改頁面: LinkedList_Edit.aspx/LinkedItem_Edit.aspx
 */
public partial class Contribute_LinkedList : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "LinkeddList";
        Util.ShowSysMsgWithScript("$('#btnPreviousPage').hide();$('#btnNextPage').hide();$('#btnGoPage').hide();");
        //權控處理
        AuthrityControl();
        if (!IsPostBack)
        {
            Inital();
            LinkList_DataBind();
            LinkItem_DataBind();
            tbxNewLinkedItems_Seq.Text = CaseUtil.GetNewMaxLinked2_Seq(HFD_Linked_Id.Value);
        }
    }
    //---------------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        Authrity.CheckButtonRight("_AddNew", btnNewLinkedFolder);
        Authrity.CheckButtonRight("_AddNew", btnNewLinkedItems);
        Authrity.CheckButtonRight("_Delete", btnDelLinkedFolder);
    }
    //---------------------------------------------------------------------------
    protected void btnDelLinkedFolder_Click(object sender, EventArgs e)
    {
        bool flag = false;
        try
        {
            //****變數宣告****//
            Dictionary<string, object> dict = new Dictionary<string, object>();

            //****設定SQL指令****//
            string strSql = " delete From Linked ";
            strSql += " where Linked_Id = @Linked_Id";

            dict.Add("Linked_Id", HFD_Linked_Id.Value);

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
            SetSysMsg("公關贈品品項刪除成功!");
            Response.Redirect(Util.RedirectByTime("LinkedList.aspx"));
        }
    }
    //---------------------------------------------------------------------------
    protected void btnNewLinkedItems_Click(object sender, EventArgs e)
    {
        string strRet = CheckFieldMustFillBasic();
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
                string strSql = "insert into  Linked2\n";
                strSql += "( Linked_Id,Linked2_Name, Linked2_Seq ) values\n";
                strSql += "(@Linked_Id,@Linked2_Name,@Linked2_Seq)";
                strSql += "\n";
                strSql += "select @@IDENTITY";

                Dictionary<string, object> dict = new Dictionary<string, object>();

                dict.Add("Linked_Id", HFD_Linked_Id.Value);
                dict.Add("Linked2_Name", tbxNewLinkedItems.Text.Trim());
                dict.Add("Linked2_Seq", tbxNewLinkedItems_Seq.Text.Trim());
                NpoDB.ExecuteSQLS(strSql, dict);

                flag = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (flag == true)
            {
                SetSysMsg("新增物品類別次分類成功！");
                Response.Redirect(Util.RedirectByTime("LinkedList.aspx", "Linked_Id=" + HFD_Linked_Id.Value));
            }
        }
    }
    //---------------------------------------------------------------------------
    public void Inital()
    {
        //tbxGiftFolder
        HFD_Linked_Id.Value = Util.GetQueryString("Linked_Id");
        if (HFD_Linked_Id.Value == "")
        {
            HFD_Linked_Id.Value = CaseUtil.GetOldMin_Good_Linked_Id();
        }
        string strSql = @"select Linked_Id,Linked_Name as 物品類別 from Linked where Linked_Type='contribute' and Linked_Id ='" + HFD_Linked_Id.Value + "'";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);
        DataRow dr = dt.Rows[0];
        tbxLinkedFolder.Text = dr["物品類別"].ToString();
    }
    //---------------------------------------------------------------------------
    public void LinkList_DataBind()
    {
        string HLink, LinkParam, LinkTarget, strSql = "";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;

        strSql = @"select Linked_Id,Linked_Name as 物品類別,Linked_Seq as 排序 from Linked where Linked_Type='contribute' order by 排序";
        dt = NpoDB.GetDataTableS(strSql, dict);

        HLink = "LinkedList.aspx?Linked_Id=";
        LinkParam = "Linked_Id";
        LinkTarget = "_self";

        Grid_List1.Text = DataLinkGrids(dt, HLink, LinkParam, LinkTarget);
    }
    //---------------------------------------------------------------------------
    public void LinkItem_DataBind()
    {
        string HLink, LinkParam, LinkParam2, DataLink, DataParam, DataTarget, strSql = "";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;

        strSql = @"select Linked_Id,Ser_No,Linked2_Name as 次分類,Linked2_Seq as 排序 from Linked2 where 1=1 and Linked_Id ='" + HFD_Linked_Id.Value + "' order by 排序";
        dt = NpoDB.GetDataTableS(strSql, dict);

        HLink = "LinkedItem_Edit.aspx?Ser_No=";
        LinkParam = "Ser_No";
        LinkParam2 = "Linked_Id";
        DataLink = "LinkedItem_Delete.aspx?Ser_No=";
        DataParam = "Ser_No";
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
        sb.AppendLine(@"<th>選項</th>");
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
            sb.AppendLine(@"<td><a href onclick=""window.open('LinkedList_Edit.aspx?Linked_Id=" + dr[0] + @"','LinkedList_Edit','scrollbars=no,status=no,toolbar=no,top=100,left=120,width=580,height=200')""><img src='../images/save.gif' border=0 align='absmiddle' alt='修改'>修改</a></td>");
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
                    sb.AppendLine(@"<td><a href onclick=""window.open('" + HLink + dr[LinkParam] + "&" + LinkParam2 + "=" + dr[LinkParam2] + @"','LinkedItem_Edit','scrollbars=no,status=no,toolbar=no,top=100,left=120,width=580,height=200')"">" + dr[dc.ColumnName] + "</a></td>");
                }
                else
                {
                    //顯示資料
                    sb.AppendLine("<td>" + dr[dc.ColumnName] + "</td>");
                }
                i++;
            }
            sb.AppendLine(@"<td><a href='JavaScript:if(confirm(""是否確定要刪除 ?"")){window.location.href=""" + DataLink + dr[DataParam] + "&Linked_Id=" + HFD_Linked_Id.Value + @""";}' target='" + DataTarget + @"'><img src='../images/delete2.gif' align='absmiddle' border=0 alt='刪除'></a></td>");
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

        if (tbxNewLinkedItems.Text.Trim() == "")
        {
            strRet += "次分類名稱 ";
        }
        if (tbxNewLinkedItems_Seq.Text.Trim() == "")
        {
            strRet += "排序 ";
        }
        if (tbxNewLinkedItems.Text.Trim() == "" || tbxNewLinkedItems_Seq.Text.Trim() == "")
        {
            strRet += "欄位不可為空白！";
        }
        return strRet;
    }
}