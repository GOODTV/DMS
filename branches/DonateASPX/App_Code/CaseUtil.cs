using System;
using System.Collections.Generic;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.IO;
using System.Web;
using System.Web.SessionState;


/// <summary>
/// Summary description for CaseUtil
/// </summary>

public class CaseUtil
{
    public static HttpSessionState Session { get { return HttpContext.Current.Session; } }
    public static HttpRequest Request { get { return HttpContext.Current.Request; } }
    public static HttpResponse Response { get { return HttpContext.Current.Response; } }
    public static HttpServerUtility Server { get { return HttpContext.Current.Server; } }
    public static Page page { get { return HttpContext.Current.CurrentHandler as Page; } }

    static List<CButton> ButtonList1 = new List<CButton>();
    static List<CButton> ButtonList2 = new List<CButton>();
    //-------------------------------------------------------------------------------------------------------------
    //public static bool CheckDate(string Date, int LimitDay)
    //{
    //    //日期不能空白
    //    if (Date == "")
    //    {
    //        return false;
    //    }
    //    DateTime LimitStartDate = Util.GetStartDateOfMonth(DateTime.Now.AddMonths(-1));
    //    DateTime LimitEndDate = Util.GetEndDateOfMonth(DateTime.Now.AddMonths(-1));
    //    DateTime ServiceDate;
    //    bool flag = false;
    //    try
    //    {
    //        ServiceDate = Convert.ToDateTime(Date);
    //        //日期不能大於今天
    //        if (ServiceDate > DateTime.Now)
    //        {
    //            return false;
    //        }
    //    }
    //    catch (Exception ex)
    //    {
    //        return false;
    //    }

    //    if (DateTime.Now.Day <= LimitDay)
    //    {
    //        //5號之前, 可以能輸入上個月的資料
    //        if (ServiceDate >= LimitStartDate)
    //        {
    //            return true;
    //        }
    //        else
    //        {
    //            return false;
    //        }
    //    }
    //    else
    //    {
    //        //超過5號, 則不能輸入上個月的資料
    //        if (ServiceDate < LimitEndDate)
    //        {
    //            return false;
    //        }
    //        else
    //        {
    //            return true;
    //        }
    //    }
    //}
    //-------------------------------------------------------------------------------------------------------------
    //產生上傳檔案的 JavaScript 碼
    public static string GetUploadJavaScriptCode()
    {
        string Script = "";
        Script += "function openUploadWindow(Ap_Name) {\n";
        Script += "var PersonProfile_Uid = document.getElementById('HFD_Uid').value;\n";
        Script += "window.open('../Common/UpLoad.aspx?PersonProfile_Uid=' + PersonProfile_Uid + '&Ap_Name=' + Ap_Name + '&AllowType=img', 'upload', 'scrollbars=no,status=no,toolbar=no,top=100,left=120,width=560,height=120');\n";
        Script += "}\n";
        return Script;
    }
    //-------------------------------------------------------------------------------------------------------------
    public static List<CButton> GetButtonList(int MenuNo)
    {
        if (MenuNo == 1)
        {
            return ButtonList1;
        }
        else if (MenuNo == 2)
        {
            return ButtonList2;
        }
        return null;
    }
    //-------------------------------------------------------------------------------------------------------------
    public static string MakeMenu(string StreetUID, int SelectedTab)
    {
        ButtonList1.Clear();
        CButton Btn = new CButton();
        Btn.TabNo = 1;
        Btn.Text = "寒士開案評估表";
        Btn.Style = "style='width:35mm'";  //需要不同寬度時,可以在此設定新值
        Btn.ImgSrc = "../images/toolbar_exit.gif";
        Btn.OnClick = "Street_Edit.aspx?StreetUID=" + StreetUID;
        ButtonList1.Add(Btn);

        Btn = new CButton();
        Btn.TabNo = 2;
        Btn.Text = "特殊事件處理";
        Btn.Style = "style='width:30mm'";  //需要不同寬度時,可以在此設定新值
        Btn.ImgSrc = "../images/toolbar_modify.gif";
        Btn.OnClick = "SpecialEvent.aspx?StreetUID=" + StreetUID;
        ButtonList1.Add(Btn);

        Btn = new CButton();
        Btn.TabNo = 3;
        Btn.Text = "夜宿申請表";
        Btn.Style = "style='width:25mm'";  //需要不同寬度時,可以在此設定新值
        Btn.ImgSrc = "../images/toolbar_familytree.gif";
        Btn.OnClick = "NightSleep.aspx?StreetUID=" + StreetUID;
        ButtonList1.Add(Btn);

        Btn = new CButton();
        Btn.TabNo = 4;
        Btn.Text = "急難救助金暨醫療費申請表";
        Btn.Style = "style='width:50mm'";  //需要不同寬度時,可以在此設定新值
        Btn.ImgSrc = "../images/toolbar_new.gif";
        Btn.OnClick = "MoneyApply.aspx?StreetUID=" + StreetUID;
        ButtonList1.Add(Btn);

        return MenuList(ButtonList1, SelectedTab);
    }
    //-------------------------------------------------------------------------
    public static string MakeMenu2(string WomanUID, int SelectedTab)
    {
        ButtonList2.Clear();
        CButton Btn = new CButton();
        Btn.TabNo = 1;
        Btn.Text = "基本資料";
        Btn.Style = "style='width:35mm'";  //需要不同寬度時,可以在此設定新值
        Btn.ImgSrc = "../images/toolbar_exit.gif";
        Btn.OnClick = "Woman_Edit.aspx?WomanUID=" + WomanUID;
        ButtonList2.Add(Btn);

        Btn = new CButton();
        Btn.TabNo = 2;
        Btn.Text = "家系圖與附件";
        Btn.Style = "style='width:35mm'";  //需要不同寬度時,可以在此設定新值
        Btn.ImgSrc = "../images/toolbar_familytree.gif";
        Btn.OnClick = "WomanFamilyTree_Edit.aspx?WomanUID=" + WomanUID;
        ButtonList2.Add(Btn);

        Btn = new CButton();
        Btn.TabNo = 3;
        Btn.Text = "開案評估表";
        Btn.Style = "style='width:30mm'";  //需要不同寬度時,可以在此設定新值
        Btn.ImgSrc = "../images/toolbar_new.gif";
        Btn.OnClick = "Woman_Edit2.aspx?WomanUID=" + WomanUID;
        ButtonList2.Add(Btn);

        return MenuList(ButtonList2, SelectedTab);
    }
    //-------------------------------------------------------------------------
    public static string MenuList(List<CButton> ButtonList, int SelectedTab)
    {
        string s = "";
        foreach (CButton btn in ButtonList)
        {
            if (btn.TabNo == SelectedTab)
            {
                btn.CssClass = "tabSelected";
            }
            else
            {
                btn.CssClass = "tabNormal";
            }

            string disabled = "";
            if (btn.Disabled == true)
            {
                disabled = "disabled='disabled'";
            }
            //s += "<button type='button' " + btn.Style + " class='" + btn.CssClass + "' onclick=\"javascript:location.href='" + btn.OnClick + "' \"><img src='" + btn.ImgSrc + "' width='" + btn.Width + "' height='" + btn.Height + "' align='" + btn.Align + "' />" + btn.Text + "</button>\n";
            s += "<button type='button' title='" + btn.Title + "' " + disabled + " class='" + btn.CssClass + "' " + btn.Style + " onclick=\"javascript:location.href='" + btn.OnClick + "' \"><img src='" + btn.ImgSrc + "' width='" + btn.Width + "' height='" + btn.Height + "' align='" + btn.Align + "' />" + btn.Text + "</button>\n";
            if (btn.ShowBR == true)
            {
                s += "<br/>";
            }
        }
        return s;
    }
    //-------------------------------------------------------------------------
    public static string GetCaseInfo(string StreetUID)
    {
        string strSql = @"
                         select 
                         d.DeptName, s.CaseNo, s.CaseName, CaseStatus
                         from StreetData s
                         inner join Dept d on s.DeptID=d.DeptID
                         where s.uid=@uid
                         and isnull(s.IsDelete, 'N') = 'N'

                         ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("uid", StreetUID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count != 0)
        {
            return "站別：" + dt.Rows[0]["DeptName"].ToString() +
                   "　　案號：" + dt.Rows[0]["CaseNo"].ToString() + 
                   "　　案主姓名：" + dt.Rows[0]["CaseName"].ToString() +
                   "　　個案狀態：" + dt.Rows[0]["CaseStatus"].ToString();
        }
        return "";
    }
    //-------------------------------------------------------------------------------------------------------------} //end of class CaseUtil
    public static string GetWomanInfo(string WomanUID)
    {
        string strSql = @"
                         select 
                         d.DeptName, w.CaseNo, w.CaseName, CaseStatus
                         from Woman w
                         inner join Dept d on w.DeptID=d.DeptID
                         where w.uid=@uid
                         and isnull(w.IsDelete, 'N') = 'N'
                         ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("uid", WomanUID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count != 0)
        {
            return "站別：" + dt.Rows[0]["DeptName"].ToString() +
                   "　　案號：" + dt.Rows[0]["CaseNo"].ToString() +
                   "　　案主姓名：" + dt.Rows[0]["CaseName"].ToString() +
                   "　　個案狀態：" + dt.Rows[0]["CaseStatus"].ToString();
        }
        return "";
    }
    //-------------------------------------------------------------------------------------------------------------} //end of class CaseUtil
    public static string GetUploadPic(string ObjectID, string ApName, Button UploadButton, Button DelButton, string Width, string Height)
    {
        if (Width == "" && Height == "")
        {
            Width = "width:180px";
            Height = "height:180px";
        }
        else if (Width == "0" && Height == "0")
        {
            Width = "";
            Height = "";
        }
        else
        {
            Width = "width:" + Width;
            Height = "height:" + Height;
        }

        Dictionary<string, object> dict = new Dictionary<string, object>();
        //帶出相關圖檔資料
        string strSql = " select ApName, UploadFileURL from Upload\n";
        strSql += " where ObjectID=@ObjectID and ApName=@ApName\n";
        dict.Add("ObjectID", ObjectID);
        dict.Add("ApName", ApName);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        string PictURL = "";
        if (dt.Rows.Count != 0)
        {
            PictURL = dt.Rows[0]["UploadFileURL"].ToString(); ;
        }
        //如果沒有URL時(代表無上傳圖檔)，關閉圖檔刪除按鈕
        if (PictURL == "")
        {
            UploadButton.Visible = true;
            DelButton.Visible = false;
        }
        else
        {
            UploadButton.Visible = false;
            DelButton.Visible = true;
        }
        //以此URL顯示圖片
        string RetStr;
        if (PictURL == "")
        {
            RetStr = "";
        }
        else
        {
            //RetStr = "<img src=\".." + PictURL + "\" border=\"0\" style=\"width:180pt;height:180pt;cursor:hand \" onclick=\"var x=window.open('','','height=480, width=640, toolbar=no, menubar=no, scrollbars=1, resizable=yes, location=no, status=no');x.document.write('<img src=.." + PictURL + " border=0>');\" alt=\"點選看放大圖\">";
            RetStr = "<img src='.." + PictURL + "' border='0' style='" + Width + ";" + Height + ";cursor:hand' onclick=\"var x=window.open('','','height=480, width=640, toolbar=no, menubar=no, scrollbars=1, resizable=yes, location=no, status=no');x.document.write('<img src=.." + PictURL + " border=0>');\" alt=\"點選看放大圖\">";
        }
        return RetStr;
    }
    //----------------------------------------------------------------------
    public static void DeletePic(string ObjectID, string ApName, string URL)
    {
        string strSql = @"
                  select uid, UploadFileURL from Upload
                  where ObjectID=@ObjectID and ApName=@ApName
              ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("ObjectID", ObjectID);
        dict.Add("ApName", ApName); //以此名稱區分是哪一個程式上傳此檔
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count != 0)
        {
            strSql = "delete from Upload where uid=@uid";
            dict.Add("uid", dt.Rows[0]["uid"].ToString());
            NpoDB.ExecuteSQLS(strSql, dict);
            //檔案一併刪除
            string FilePath = Server.MapPath(".." + dt.Rows[0]["UploadFileURL"].ToString());
            File.Delete(FilePath);
        }
        Session["Msg"] = "圖檔刪除成功!";
        Response.Redirect(URL);
    }
    //----------------------------------------------------------------------
    public static void FillDept(DropDownList ddl)
    {
        string strSql = @"
                         select DeptID, DeptName
                         from Dept where DeptType='1'
                         order by DeptID
                       ";
        Util.FillDropDownList(ddl, strSql, "DeptName", "DeptID", true);
    }
    //----------------------------------------------------------------------
    public static bool IsAllDept()
    {
        if ((page as BasePage).SessionInfo.GroupArea == "全區")
        {
            return true;
        }
        return false;
    }
    //----------------------------------------------------------------------
    //列印時傳回 img 碼
    public static string GetUploadPic(string ObjectID, string ApName, string Width, string Height)
    {
        if (Width == "" && Height == "")
        {
            Width = "width:180px";
            Height = "height:180px";
        }
        else if (Width == "0" && Height == "0")
        {
            Width = "";
            Height = "";
        }
        else
        {
            Width = "width:" + Width;
            Height = "height:" + Height;
        }

        Dictionary<string, object> dict = new Dictionary<string, object>();
        //帶出相關圖檔資料
        string strSql = " select ApName, UploadFileURL from Upload\n";
        strSql += " where ObjectID=@ObjectID and ApName=@ApName\n";
        dict.Add("ObjectID", ObjectID);
        dict.Add("ApName", ApName);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        string PictURL = "";
        if (dt.Rows.Count != 0)
        {
            PictURL = dt.Rows[0]["UploadFileURL"].ToString(); ;
        }
        //以此URL顯示圖片
        string RetStr;
        if (PictURL == "")
        {
            RetStr = "";
        }
        else
        {
            RetStr = "<img src='.." + PictURL + "' border='0' style='" + Width + ";" + Height + "'";
        }
        return RetStr;
    }
    //----------------------------------------------------------------------
    public static string GetDeptNameByDeptID(string DeptID)
    {
        string DeptName = "";
        string strSql = @"
                          select DeptName
                          from Dept
                          where DeptID=@DeptID
                         ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("DeptID", DeptID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count != 0)
        {
            DeptName = dt.Rows[0]["DeptName"].ToString();
        }
        return DeptName;
    }
    //-------------------------------------------------------------------------------------------------------------
} //end of CaseUtil

public class CButton
{
    private int _TabNo;
    private string _CssClass;
    private string _OnClick;
    private string _ImgSrc;
    private string _Width = "20";
    private string _Height = "20";
    private string _Align = "absbottom";
    private string _Title = "";
    private string _Text;
    private string _Style;
    private bool _ShowBR = false;
    private bool _Disabled = false;

    public int TabNo
    {
        get { return _TabNo; }
        set { _TabNo = value; }
    }
    public string CssClass
    {
        get { return _CssClass; }
        set { _CssClass = value; }
    }
    public string OnClick
    {
        get { return _OnClick; }
        set { _OnClick = value; }
    }
    public string ImgSrc
    {
        get { return _ImgSrc; }
        set { _ImgSrc = value; }
    }
    public string Width
    {
        get { return _Width; }
        set { _Width = value; }
    }
    public string Height
    {
        get { return _Height; }
        set { _Height = value; }
    }
    public string Align
    {
        get { return _Align; }
        set { _Align = value; }
    }
    public string Title
    {
        get { return _Title; }
        set { _Title = value; }
    }
    public string Text
    {
        get { return _Text; }
        set { _Text = value; }
    }
    public string Style
    {
        get { return _Style; }
        set { _Style = value; }
    }
    public bool ShowBR
    {
        get { return _ShowBR; }
        set { _ShowBR = value; }
    }
    public bool Disabled
    {
        get { return _Disabled; }
        set { _Disabled = value; }
    }
}
