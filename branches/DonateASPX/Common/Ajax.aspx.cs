using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Threading;

public partial class Ajax : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string Type = Request.Form["Type"];
        string result = "";
        switch (Type)
        {
            case "1":  //以 CityID 取得其鄉鎮列表
                result = GetAreaList();
                break;
            case "2":  //同個案戶籍地址
                result = GetHouseholdAddress();
                break;
            case "3":  //新聞
                result = GetNewsDetail();
                break;
            case "4":  //檢核Email是否存在
                result = EmailIsExist();
                break;
        }
        Response.Write(result);
    }
    //-------------------------------------------------------------------------
    private string GetNewsDetail()
    {
        //string NewsID = Request.Form["NewsID"];
        //string strSql = "select u.User_name, o.Org_name, news.*\n";
        //strSql += "from News\n";
        //strSql += "left join Admin_User u on News.PostUser_id=u.User_id\n";
        //strSql += "left join Organization o on u.Org_id=o.Org_id\n";
        //strSql += "where mbr_status=1 and mbr_num=@mbr_num";
        //Dictionary<string, object> dict = new Dictionary<string, object>();
        //dict.Add("mbr_num", NewsID);
        ////dict.Add("MBR_CHKREAD", Session["user_uid"].ToString());
        //DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        //if (dt.Rows.Count == 0)   6
        //{
        //    return "無此新聞";
        //}
        //DataRow dr = dt.Rows[0];
        //string strTemp = "";
        //HtmlTable table = new HtmlTable();
        //HtmlTableRow row;
        //HtmlTableCell cell;
        //CssStyleCollection css;

        //table.Border = 1;
        //table.CellPadding = 0;
        //table.CellSpacing = 0;
        //css = table.Style;
        //css.Add("font-size", "24px");
        //css.Add("font-family", "標楷體");
        //css.Add("width", "100%");

        //string strFontSize = "14px";

        //string strTitle = "系統公告";
        //row = new HtmlTableRow();
        //cell = new HtmlTableCell();
        //cell.InnerHtml = strTitle;
        //css = cell.Style;
        //css.Add("text-align", "center");
        //css.Add("font-size", "24");
        //css.Add("background-color", "#9CC2E7");
        //cell.ColSpan = 2;
        //row.Cells.Add(cell);
        //table.Rows.Add(row);

        //row = new HtmlTableRow();
        //cell = new HtmlTableCell();
        //cell.InnerHtml = "時間";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //css.Add("width", "20%");
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = GSSNPO.TransCDate(DBFunction.DB_DateTimeToShortString(dr["Mbr_Date"]));
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //table.Rows.Add(row);
        ////--------------------------------------------
        //row = new HtmlTableRow();

        //cell = new HtmlTableCell();
        //cell.InnerHtml = "標題";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = dr["Mbr_Title"].ToString();
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);

        //row.Cells.Add(cell);
        //table.Rows.Add(row);
        ////--------------------------------------------
        //row = new HtmlTableRow();

        //cell = new HtmlTableCell();
        //cell.InnerHtml = "發佈單位";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = "OrgName";
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);
        //table.Rows.Add(row);
        ////--------------------------------------------
        //row = new HtmlTableRow();

        //cell = new HtmlTableCell();
        //cell.InnerHtml = "發佈者";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = "Username";
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);
        //table.Rows.Add(row);
        ////--------------------------------------------
        //row = new HtmlTableRow();

        //cell = new HtmlTableCell();
        //cell.InnerHtml = "內容";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = "<textarea name='a' rows='20' cols='15' id='a' readonly='readonly' class='font9' style='width:100%;'>" + dr["Mbr_Content"].ToString() + "</textarea>";
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);
        //table.Rows.Add(row);
        ////--------------------------------------------
        //row = new HtmlTableRow();

        //cell = new HtmlTableCell();
        //cell.InnerHtml = "附件";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = dr["Mbr_File"].ToString();
        //string RealFilename = strTemp.Substring(strTemp.LastIndexOf('_') + 1);
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);
        //table.Rows.Add(row);
        ////--------------------------------------------

        ////轉成 html 碼
        //StringWriter sw = new StringWriter();
        //HtmlTextWriter htw = new HtmlTextWriter(sw);
        //table.RenderControl(htw);

        //return htw.InnerWriter.ToString();
        return "";
    }
    //-------------------------------------------------------------------------
    private string GetHouseholdAddress()
    {
        //string CaseData_Uid = Request.Form["CaseData_Uid"];

        //string strSql = "";
        //if (type == 1)
        //{
        //    strSql += "select HouseholdCity as City, HouseholdVillage as Village, HouseholdZip as Zip, HouseholdAddress as Address\n";
        //}
        //else
        //{
        //    strSql += "select ContactCity as City, ContactVillage as Village, ContactZip as Zip, ContactAddress as Address\n";
        //}
        //strSql += "from ServiceAdvisory sa\n";
        //strSql += "inner join CaseData cd on sa.uid=cd.ServiceAdvisory_Uid\n";
        //strSql += "where cd.uid=@uid\n";

        //Dictionary<string, object> dict = new Dictionary<string, object>();
        //dict.Add("uid", CaseData_Uid);
        //DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        //if (dt.Rows.Count != 0)
        //{
        //    return dt.Rows[0]["City"].ToString().Trim() + "|" +
        //           dt.Rows[0]["Village"].ToString().Trim() + "|" +
        //           dt.Rows[0]["Zip"].ToString().Trim() + "|" +
        //           dt.Rows[0]["Address"].ToString().Trim();
        // }
        return "";
    }
    //-------------------------------------------------------------------------
    private string GetAreaList()
    {
        string CityID = Request.Form["CityID"];

        string strSql = "select ZipCode, Name from CodeCityNew\n";
        strSql += "where ParentCityID=@CityID\n";
        strSql += "order by Sort\n";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("CityID", CityID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        string retStr = "";
        retStr += "<option value=''></option>";
        foreach (DataRow dr in dt.Rows)
        {
            retStr += "<option value='" + dr["ZipCode"].ToString() + "'>" + dr["Name"].ToString() + "</option>";
        }
        return retStr;
    }
    //-------------------------------------------------------------------------
    private string EmailIsExist()
    {
        //Thread.Sleep(20000);
        string strRet = "N"; 
        string Email = Request.Form["Email"];

        string strSql = "select Email from Donor\n";
        strSql += "where Email=@Email\n";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Email", Email);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            strRet = "Y";
        }
        return strRet;
    }
    //-------------------------------------------------------------------------
}