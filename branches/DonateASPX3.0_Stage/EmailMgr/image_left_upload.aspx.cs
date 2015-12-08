using System;
using System.Data;
using System.Configuration;
using System.Collections.Generic ;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO ;
/******************************************************
程式的HEADSCRIPT要放這一段
 * <script type="text/javascript"  language="javascript">
     function openUploadWindow()
     {
        window.open("FileUpLoad.aspx?Ap_Name=UpLoad","upload","scrollbars=no,status=no,toolbar=no,top=100,left=120,width=580,height=200");
     }
</script>
程式的原始碼裡面要放這一段
<asp:Button runat="server" Text=" 上傳新檔案 " ID="btnNewFile" CssClass="cbutton" OnClientClick="openUploadWindow();return false;" />
******************************************************/

public partial class Function_image_left_upload : BasePage 
{
    //folderPath 讀取Global.asax 中的 Session["FolderPath"]，以網址型式
    private string folderPath;//記錄檔案存檔目錄
    private string fileName;//記錄檔案名稱
    private string Upload_FileSize;//記錄檔案大小
    private string old_fileName;//記錄有縮圖原檔檔案名稱
    private string Upload_FileSize_Old;//記錄有縮圖原檔檔案大小

    protected void Page_Load(object sender, EventArgs e)
    {
        folderPath = Util.GetAppSetting("UploadPath");
        if (!IsPostBack)
        {
            Inital();
        }
    }
    //---------------------------------------------------------------------------
    public void Inital()
    {
        HFD_item.Value = Util.GetQueryString("item");
        HFD_ser_no.Value = Util.GetQueryString("ser_no");
        HFD_subject.Value = Util.GetQueryString("subject");
        HFD_img_width.Value = Util.GetQueryString("img_width");
        HFD_img_height.Value = Util.GetQueryString("img_height");
        HFD_code_id.Value = Util.GetQueryString("code_id");
        HFD_id.Value = Util.GetQueryString("id");
    }
    //---------------------------------------------------------------------------
    protected void btnSave_Click(object sender, EventArgs e)
    {
        if (!FileUpload1.HasFile)
        {
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript'>alert('請選擇檔案 ！');</script>");
            return;
        }
        else if (Util.GetControlValue("resize") == "Y" && Util.GetControlValue("img_width").Trim() == "" && Util.GetControlValue("img_height").Trim() == "")
        {
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript'>alert('寬度或高度 不可為空白！');</script>");
            return;
        }
        if (FileNameValid() == false)
            return;
        if (!AllowSave())
            return;
        //****上傳資料處理****//
         FileUpLoad();

        //***********************************
        // Insert To DataBase
        //***********************************
        //****變數宣告****//
        string strSql;
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****變數設定****//
        string filePath = FileUpload1.FileName;
        string strImageFile = DateTime.Now.Year.ToString() + @"/" + fileName;
                	
		List<ColumnData> list = new List<ColumnData>();
        list.Add(new ColumnData("ObjectID", Util.GetControlValue("HFD_ser_no"), true, false, false));
        list.Add(new ColumnData("ApName", Util.GetControlValue("HFD_item"), true, false, false));
        list.Add(new ColumnData("AttachType", "img", true, false, false));
        list.Add(new ColumnData("UploadFileURL", fileName, true, false, false));
        list.Add(new ColumnData("UploadDate", DateTime.Now, true, false, false));
        strSql = Util.CreateInsertCommand("UPLOAD", list, dict);
        NpoDB.ExecuteSQLS(strSql, dict);    
        
        Session["Msg"]="圖檔上傳成功 ！";
        //Util.Sys_log(SessionInfo.UserID, SessionInfo.UserName, SessionInfo.DeptID, Util.GetControlValue("HFD_item"), "圖檔上傳：" + Util.GetControlValue("HFD_ser_no"));
        this.Response.Write(@"<script> window.opener.location.href='image_left_list.aspx?dept_id=" + SessionInfo.DeptID + "&item=" + Util.GetControlValue("HFD_item") + "&code_id=" + Util.GetControlValue("HFD_code_id") + "&ser_no=" + Util.GetControlValue("HFD_ser_no") + "&id=" + Util.GetControlValue("HFD_id") + "';window.close();</script>");                        	            
    }
    //---------------------------------------------------------------------------
    //檢查及限定上傳檔案類型與大小
    public bool AllowSave()
    {
        int DenyMbSize = 10;//限制大小MB
        bool flag = false;
        //判斷副檔名
        switch (System.IO.Path.GetExtension(FileUpload1.PostedFile.FileName).ToUpper())
        {
            case ".JPG":
            case ".JPEG":
            case ".BMP":
            case ".PNG":
            //case ".PDF":
            case ".GIF":
                flag = true;
                break;
            default :
                //this.RegisterStartupScript("js", @"<script language='javascript' >alert('此檔案類型不允許上傳');</script>"); 
                this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('此檔案類型不允許上傳');</script>"); 
                flag= false;
                break;
        }
        if (GetFileSizeMB(FileUpload1.FileBytes.Length ) > DenyMbSize)
        {
            //this.RegisterStartupScript("js", @"<script language='javascript' >alert('檔案大小不得超過" + DenyMbSize + @"MB');</script>");
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('檔案大小不得超過" + DenyMbSize + @"MB');</script>");
            flag = false;
        }
        return flag;
    }
    //---------------------------------------------------------------------------
    //取得單位為KB的檔案Size
    public double GetFileSizeKB(int intLength)
    {
        int intKbSize = 1024;
        return intLength / intKbSize;
    }
    //---------------------------------------------------------------------------
    //取得單位為MB的檔案Size
    public double GetFileSizeMB(int intLength)
    {
        int intMbSize = 1048576;
        return intLength / intMbSize;
    }
    //---------------------------------------------------------------------------
    //處理上傳檔案
    public void FileUpLoad()
    {
        string UploadFileFolderPath, UploadFilePath;
        string filePath;
        filePath = FileUpload1.FileName;
        //UploadFileFolderPath = Server.MapPath("~" + folderPath);
        UploadFileFolderPath = Server.MapPath("~" + folderPath) + @"/";
        if (!Directory.Exists(UploadFileFolderPath))
        {
            Directory.CreateDirectory(UploadFileFolderPath);
        }
        this.fileName = "{" + DateTime.Now.Ticks.ToString() + "}_" + filePath.Substring(filePath.LastIndexOf(@"\") + 1);
        this.old_fileName = "{" + DateTime.Now.Ticks.ToString() + "}_" + filePath.Substring(filePath.LastIndexOf(@"\") + 1);
        //this.fileSize = FileUpload1.FileBytes.Length;
        UploadFilePath = UploadFileFolderPath + this.fileName;
        //上傳原相片
        FileUpload1.SaveAs(UploadFilePath);

        if (Util.GetControlValue("resize") == "Y")
        {
            //縮圖
            string strThumNailPath = UploadFileFolderPath + "S_{" + DateTime.Now.Ticks.ToString() + "}_" + filePath.Substring(filePath.LastIndexOf(@"\") + 1);
            if (Util.GetControlValue("img_width").Trim() != "" && Util.GetControlValue("img_height").Trim() != "")//指定高寬縮放,可能變形
            {
                MakeThumNail(UploadFilePath, strThumNailPath, Convert.ToUInt16(Util.GetControlValue("img_width").Trim()), Convert.ToUInt16(Util.GetControlValue("img_height").Trim()), "HW");
            }
            else if (Util.GetControlValue("img_width").Trim() != "" && Util.GetControlValue("img_height").Trim() == "")//指定寬度,高度按照比例縮放
            {
                MakeThumNail(UploadFilePath, strThumNailPath, Convert.ToUInt16(Util.GetControlValue("img_width").Trim()), 0, "W");
            }
            else if (Util.GetControlValue("img_width").Trim() == "" && Util.GetControlValue("img_height").Trim() != "")//指定高度,寬度按照等比例縮放
            {
                MakeThumNail(UploadFilePath, strThumNailPath, 0, Convert.ToUInt16(Util.GetControlValue("img_height").Trim()), "H");
            }            
            //刪除原相片
            File.Delete(UploadFilePath);
            this.fileName = strThumNailPath.Replace(UploadFileFolderPath, "");
            FileInfo fInfo = new FileInfo(strThumNailPath);
            Upload_FileSize_Old = fInfo.Length.ToString();
        }
    }
    //---------------------------------------------------------------------------
    public bool FileNameValid()
    {
        string strfilePath = FileUpload1.FileName;
        string strfileName =  strfilePath.Substring(strfilePath.LastIndexOf(@"\") + 1);
        if (strfileName.Contains(" ") || strfileName.Contains("　"))
        {
            //this.RegisterStartupScript("js", @"<script language='javascript' >alert('檔案名稱不能有空白');</script>");
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('檔案名稱不能有空白');</script>");
            return false;
        }
        return true;
    }
    //---------------------------------------------------------------------------
    public static void MakeThumNail(string originalImagePath, string thumNailPath, int width, int height, string model)
    {
        System.Drawing.Image originalImage = System.Drawing.Image.FromFile(originalImagePath);

        int thumWidth = width;      //縮略圖的寬度
        int thumHeight = height;    //縮略圖的高度

        int x = 0;
        int y = 0;

        int originalWidth = originalImage.Width;    //原始圖片的寬度
        int originalHeight = originalImage.Height;  //原始圖片的高度

        switch (model)
        {
            case "HW":      //指定高寬縮放,可能變形
                break;
            case "W":       //指定寬度,高度按照比例縮放
                thumHeight = originalImage.Height * width / originalImage.Width;
                break;
            case "H":       //指定高度,寬度按照等比例縮放
                thumWidth = originalImage.Width * height / originalImage.Height;
                break;
            case "Cut":
                if ((double)originalImage.Width / (double)originalImage.Height > (double)thumWidth / (double)thumHeight)
                {
                    originalHeight = originalImage.Height;
                    originalWidth = originalImage.Height * thumWidth / thumHeight;
                    y = 0;
                    x = (originalImage.Width - originalWidth) / 2;
                }
                else
                {
                    originalWidth = originalImage.Width;
                    originalHeight = originalWidth * height / thumWidth;
                    x = 0;
                    y = (originalImage.Height - originalHeight) / 2;
                }
                break;
            default:
                break;
        }

        //新建一個bmp圖片
        System.Drawing.Image bitmap = new System.Drawing.Bitmap(thumWidth, thumHeight);

        //新建一個畫板
        System.Drawing.Graphics graphic = System.Drawing.Graphics.FromImage(bitmap);

        //設置高質量查值法
        graphic.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;

        //設置高質量，低速度呈現平滑程度
        graphic.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

        //清空畫布並以透明背景色填充
        graphic.Clear(System.Drawing.Color.Transparent);

        //在指定位置並且按指定大小繪製原圖片的指定部分
        graphic.DrawImage(originalImage, new System.Drawing.Rectangle(0, 0, thumWidth, thumHeight), new System.Drawing.Rectangle(x, y, originalWidth, originalHeight), System.Drawing.GraphicsUnit.Pixel);

        try
        {
            bitmap.Save(thumNailPath, System.Drawing.Imaging.ImageFormat.Jpeg);
        }
        catch (Exception ex)
        {

            throw ex;
        }
        finally
        {
            originalImage.Dispose();
            bitmap.Dispose();
            graphic.Dispose();
        }        
    }
    //---------------------------------------------------------------------------
}
