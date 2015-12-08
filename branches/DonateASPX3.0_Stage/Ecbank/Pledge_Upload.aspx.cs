using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using System.Web.Security;
using System.IO;
using System.Text;
using System.Data;

public partial class Ecbank_Pledge_Upload : BasePage
{
    Microsoft.Office.Interop.Excel.Application xlApp = null;
    Workbook wb = null;
    Worksheet ws = null;
    Range aRange = null;
    //*******************************/
    //要上傳Excel檔的Server端 檔案總管目錄
    string upload_excel_Dir = HttpContext.Current.Server.MapPath("~/UpLoad/");
    
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "Pledge_Upload";
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
        }
    }
    protected void btnUpload_Click(object sender, EventArgs e)
    {
        if (!FileUpload.HasFile)
        {
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript'>alert('請選擇檔案');</script>");
            return;
        }
        if (FileNameValid() == false)
            return;
        if (!AllowSave())
            return;

        string filePath = "";
        string file_name = System.IO.Path.GetFileNameWithoutExtension(FileUpload.PostedFile.FileName);
        string file_extension = System.IO.Path.GetExtension(FileUpload.PostedFile.FileName);
        try
        {
            filePath = SaveFileAndReturnPath(file_name, file_extension);//先上傳檔案給Server

            //寫入資料表UpLoad中，以便下載用
            string strSql = "insert into  UpLoad\n";
            strSql += "( ObjectID, ApName, AttachType, UploadFileURL, UploadDate) values\n";
            strSql += "( @ObjectID,@ApName,@AttachType,@UploadFileURL,@UploadDate) ";
            strSql += "\n";
            strSql += "select @@IDENTITY";

            Dictionary<string, object> dict = new Dictionary<string, object>();
            dict.Add("ObjectID", "0");
            dict.Add("ApName", "pledge");
            dict.Add("AttachType", file_extension);
            dict.Add("UploadFileURL", upload_excel_Dir + file_name + file_extension);
            dict.Add("UploadDate", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));

            NpoDB.ExecuteSQLS(strSql, dict);
        }
        catch (Exception ex)
        {
            throw ex;
        }
        ShowSysMsg("轉帳授權書上傳成功！");
    }
    protected void btnDownload_Click(object sender, EventArgs e)
    {
        string strSql = @"Select TOP 1 * From UpLoad Where ObjectID='0' And ApName='pledge' And AttachType='.doc' Order By uid Desc";
        System.Data.DataTable dt;
        dt = NpoDB.QueryGetTable(strSql);
        //沒上傳檔案就下載，下載範例
        if (dt.Rows.Count == 0)
        {
            if (File.Exists(upload_excel_Dir+"定額捐款授權4合一表.doc"))
            {
                FileInfo xpath_file = new FileInfo(upload_excel_Dir + "定額捐款授權4合一表.doc");  //要 using System.IO;
                // 將傳入的檔名以 FileInfo 來進行解析（只以字串無法做）
                System.Web.HttpContext.Current.Response.Clear(); //清除buffer
                System.Web.HttpContext.Current.Response.ClearHeaders(); //清除buffer 表頭
                System.Web.HttpContext.Current.Response.Buffer = false;
                System.Web.HttpContext.Current.Response.ContentType = "application/octet-stream";
                // 檔案類型還有下列幾種"application/pdf"、"application/vnd.ms-excel"、"text/xml"、"text/HTML"、"image/JPEG"、"image/GIF"
                System.Web.HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment;filename=" + System.Web.HttpUtility.UrlEncode("定額捐款授權4合一表") + ".doc");
                // 考慮 utf-8 檔名問題，以 out_file 設定另存的檔名
                System.Web.HttpContext.Current.Response.AppendHeader("Content-Length", xpath_file.Length.ToString()); //表頭加入檔案大小
                System.Web.HttpContext.Current.Response.WriteFile(xpath_file.FullName);
                // 將檔案輸出
                System.Web.HttpContext.Current.Response.Flush();
                // 強制 Flush buffer 內容
                System.Web.HttpContext.Current.Response.End();
            }
        }
        else
        {
            DataRow dr = dt.Rows[0];
            {
                if (File.Exists(dr["UploadFileURL"].ToString()))
                {
                    FileInfo xpath_file = new FileInfo(dr["UploadFileURL"].ToString());  //要 using System.IO;
                    // 將傳入的檔名以 FileInfo 來進行解析（只以字串無法做）
                    System.Web.HttpContext.Current.Response.Clear(); //清除buffer
                    System.Web.HttpContext.Current.Response.ClearHeaders(); //清除buffer 表頭
                    System.Web.HttpContext.Current.Response.Buffer = false;
                    System.Web.HttpContext.Current.Response.ContentType = "application/octet-stream";
                    // 檔案類型還有下列幾種"application/pdf"、"application/vnd.ms-excel"、"text/xml"、"text/HTML"、"image/JPEG"、"image/GIF"
                    System.Web.HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment;filename=" + System.Web.HttpUtility.UrlEncode(dr["ApName"].ToString(), System.Text.Encoding.UTF8) + ".doc");
                    // 考慮 utf-8 檔名問題，以 out_file 設定另存的檔名
                    System.Web.HttpContext.Current.Response.AppendHeader("Content-Length", xpath_file.Length.ToString()); //表頭加入檔案大小
                    System.Web.HttpContext.Current.Response.WriteFile(xpath_file.FullName);
                    // 將檔案輸出
                    System.Web.HttpContext.Current.Response.Flush();
                    // 強制 Flush buffer 內容
                    System.Web.HttpContext.Current.Response.End();
                }
                else
                {
                    Page.ClientScript.RegisterStartupScript(this.GetType(), "Prompt", "Confirm();", true);
                }
            }
        }
    }
    //---------------------------------------------------------------------------
    public bool FileNameValid()
    {
        string strfilePath = FileUpload.FileName;
        string strfileName = strfilePath.Substring(strfilePath.LastIndexOf(@"\") + 1);
        if (strfileName.Contains(" ") || strfileName.Contains("　"))
        {
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('檔案名稱不能有空白');</script>");
            return false;
        }
        return true;
    }
    //檢查及限定上傳檔案類型與大小
    public bool AllowSave()
    {
        bool flag = false;
        //判斷副檔名
        switch (System.IO.Path.GetExtension(FileUpload.PostedFile.FileName))
        {
            case ".xls":
            case ".doc":
            case ".pdf":
                flag = true;
                break;
            default:
                this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('轉帳授權書 檔案名稱錯誤，副檔名必須為.doc或.pdf或.xls');</script>");
                flag = false;
                break;
        }
        return flag;
    }
    //儲存EXCEL檔案給Server
    private string SaveFileAndReturnPath(string file_name, string file_extension)
    {
        string return_file_path = "";//上傳的Excel檔在Server上的位置
        if (FileUpload.FileName != "")
        {
            return_file_path = System.IO.Path.Combine(this.upload_excel_Dir, file_name + file_extension);

            FileUpload.SaveAs(return_file_path);
        }
        return return_file_path;
    }
}