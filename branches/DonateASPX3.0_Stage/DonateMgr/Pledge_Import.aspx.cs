using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Security;
using System.IO;
using System.Text;
using System.Data;


public partial class DonateMgr_Pledge_Import : BasePage
{
    //*******************************/
    //要上傳TXT檔的Server端 資料夾目錄
    string upload_excel_Dir = @"\UpLoad\";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
        }
    }
    //---------------------------------------------------------------------------
    protected void btnInput_Click(object sender, EventArgs e)
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

        string txt_filePath = "";

        txt_filePath = SaveFileAndReturnPath();//先上傳TXT檔案給Server
        SaveOrInsertDB();//讀取TXT中的資料寫入DB
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
        int DenyMbSize = 30;//限制大小30MB
        bool flag = false;
        //判斷副檔名
        switch (System.IO.Path.GetExtension(FileUpload.PostedFile.FileName).ToUpper())
        {
            case ".TXT":
                flag = true;
                break;
            case ".01R":
                flag = true;
                break;
            case ".02R":
                flag = true;
                break;
            case ".03R":
                flag = true;
                break;
            default:
                //this.RegisterStartupScript("js", @"<script language='javascript' >alert('此檔案類型不允許上傳');</script>"); 
                this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('此檔案類型不允許上傳');</script>");
                flag = false;
                break;
        }
        if (Util.GetFileSizeMB(FileUpload.FileBytes.Length) > DenyMbSize)
        {
            //this.RegisterStartupScript("js", @"<script language='javascript' >alert('檔案大小不得超過" + DenyMbSize + @"MB');</script>");
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('檔案大小不得超過" + DenyMbSize + @"MB');</script>");
            flag = false;
        }
        return flag;
    }
    //儲存TXT檔案給Server
    private string SaveFileAndReturnPath()
    {
        string return_file_path = "";//上傳的Excel檔在Server上的位置
        if (FileUpload.FileName != "")
        {
            String appPath = Request.PhysicalApplicationPath;
            return_file_path = System.IO.Path.Combine(appPath + this.upload_excel_Dir, Guid.NewGuid().ToString() + ".txt");

            FileUpload.SaveAs(return_file_path);
        }
        return return_file_path;
    }
    //把TXT資料Insert into Table
    private void SaveOrInsertDB()
    {
        //先將PLEDGE_SEND_RETURN中資料刪除
        string strSql = "Delete from Pledge_Send_Return where Pledge_Id>0";
        NpoDB.ExecuteSQLS(strSql, null);
        //讀取TXT檔
        bool flag = false;
        System.IO.StreamReader red = new System.IO.StreamReader(this.FileUpload.PostedFile.InputStream, System.Text.Encoding.Default);
        while(red.Peek()>0)
        {
            string strContent = red.ReadLine();
            //基本資料
            string SourcePath = "";
            string SourceName = "";
            string UploadName = "";
            string UploadSize = "";
            string ExtName = "";
            if (strContent.Trim() != "")
            {
                UploadName = Guid.NewGuid().ToString() + ".txt";
                SourcePath = FileUpload.PostedFile.FileName.Replace(UploadName, "");
                SourceName = FileUpload.FileName;
                UploadSize = string.Format((FileUpload.PostedFile.ContentLength / 1024).ToString(), "0.00");
                ExtName = System.IO.Path.GetExtension(FileUpload.PostedFile.FileName).ToUpper();
            }
            //Data
            string Account_No = "";
            string Donate_Amt = "";
            string Return_Status = "";
            string Pledge_Id = "";
            if (strContent.IndexOf("T", 0) == -1)
            {
                Account_No = (Mid(strContent, 19, 19).Trim());
                Donate_Amt = Trim0(Mid(strContent, 45, 10));
                Return_Status = (Mid(strContent, 57, 2).Trim());
                Pledge_Id = Trim0(Mid(strContent, 76, 8));

                //insert Data
                strSql = "insert into  Pledge_Send_Return\n";
                strSql += "( Pledge_Id, Account_No, Return_Status, Donate_Amt,Return_Status_No ) values\n";
                strSql += "( @Pledge_Id,@Account_No,@Return_Status,@Donate_Amt,@Return_Status_No ) ";
                strSql += "\n";
                strSql += "select @@IDENTITY";

                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Pledge_Id", Pledge_Id);
                dict.Add("Account_No", Account_No);
                if (Return_Status == "00")
                {
                    dict.Add("Return_Status", "Y");
                }
                else
                {
                    dict.Add("Return_Status", "N");
                }
                dict.Add("Donate_Amt", Donate_Amt);
                dict.Add("Return_Status_No", Return_Status);
                NpoDB.ExecuteSQLS(strSql, dict);
                flag = true;
                
                this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript'>ReturnOpener();</script>");
            }
            
        }
        if (flag == true)
        {
            strSql = "Select Count(*) AS Count,REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,SUM(Donate_Amt)),1),'.00','') AS Sum_Amt From dbo.PLEDGE_SEND_RETURN Where Return_Status='Y'";
            DataTable dt = NpoDB.GetDataTableS(strSql, null);
            DataRow dr = dt.Rows[0];
            ShowSysMsg("回覆檔資料匯入成功，共" + dr["Count"] + "筆，金額" + dr["Sum_Amt"] + "元！");
        }
        else
        {
            ShowSysMsg("無符合的資料！");
        }
    }
    // 取得起始位置至結束之字串
    public static string Mid(string sSource, int iStart, int iLength)
    {
        if (sSource.Trim().Length > 0)
        {
            int iStartPoint = iStart > sSource.Length ? sSource.Length : iStart;
            return sSource.Substring(iStartPoint, iStartPoint + iLength > sSource.Length ? sSource.Length - iStartPoint : iLength);
        }
        else
            return "";
    }
    public static string Trim0(string sSource)
    {
        int r = 0;
        for (int i = 0; i < sSource.Length; i++)
        {
            if (sSource[i].ToString().Equals("0"))
            {
                r += 1;
            }
            else
            {
                break;
            }
        }
        return sSource.Substring(r);
    }  
}