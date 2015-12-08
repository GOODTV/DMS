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
�{����HEADSCRIPT�n��o�@�q
 * <script type="text/javascript"  language="javascript">
     function openUploadWindow()
     {
        window.open("FileUpLoad.aspx?Ap_Name=UpLoad","upload","scrollbars=no,status=no,toolbar=no,top=100,left=120,width=580,height=200");
     }
</script>
�{������l�X�̭��n��o�@�q
<asp:Button runat="server" Text=" �W�Ƿs�ɮ� " ID="btnNewFile" CssClass="cbutton" OnClientClick="openUploadWindow();return false;" />
******************************************************/

public partial class Function_image_left_upload : BasePage 
{
    //folderPath Ū��Global.asax ���� Session["FolderPath"]�A�H���}����
    private string folderPath;//�O���ɮצs�ɥؿ�
    private string fileName;//�O���ɮצW��
    private string Upload_FileSize;//�O���ɮפj�p
    private string old_fileName;//�O�����Y�ϭ����ɮצW��
    private string Upload_FileSize_Old;//�O�����Y�ϭ����ɮפj�p

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
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript'>alert('�п���ɮ� �I');</script>");
            return;
        }
        else if (Util.GetControlValue("resize") == "Y" && Util.GetControlValue("img_width").Trim() == "" && Util.GetControlValue("img_height").Trim() == "")
        {
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript'>alert('�e�שΰ��� ���i���ťաI');</script>");
            return;
        }
        if (FileNameValid() == false)
            return;
        if (!AllowSave())
            return;
        //****�W�Ǹ�ƳB�z****//
         FileUpLoad();

        //***********************************
        // Insert To DataBase
        //***********************************
        //****�ܼƫŧi****//
        string strSql;
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****�ܼƳ]�w****//
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
        
        Session["Msg"]="���ɤW�Ǧ��\ �I";
        //Util.Sys_log(SessionInfo.UserID, SessionInfo.UserName, SessionInfo.DeptID, Util.GetControlValue("HFD_item"), "���ɤW�ǡG" + Util.GetControlValue("HFD_ser_no"));
        this.Response.Write(@"<script> window.opener.location.href='image_left_list.aspx?dept_id=" + SessionInfo.DeptID + "&item=" + Util.GetControlValue("HFD_item") + "&code_id=" + Util.GetControlValue("HFD_code_id") + "&ser_no=" + Util.GetControlValue("HFD_ser_no") + "&id=" + Util.GetControlValue("HFD_id") + "';window.close();</script>");                        	            
    }
    //---------------------------------------------------------------------------
    //�ˬd�έ��w�W���ɮ������P�j�p
    public bool AllowSave()
    {
        int DenyMbSize = 10;//����j�pMB
        bool flag = false;
        //�P�_���ɦW
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
                //this.RegisterStartupScript("js", @"<script language='javascript' >alert('���ɮ����������\�W��');</script>"); 
                this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('���ɮ����������\�W��');</script>"); 
                flag= false;
                break;
        }
        if (GetFileSizeMB(FileUpload1.FileBytes.Length ) > DenyMbSize)
        {
            //this.RegisterStartupScript("js", @"<script language='javascript' >alert('�ɮפj�p���o�W�L" + DenyMbSize + @"MB');</script>");
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('�ɮפj�p���o�W�L" + DenyMbSize + @"MB');</script>");
            flag = false;
        }
        return flag;
    }
    //---------------------------------------------------------------------------
    //���o��쬰KB���ɮ�Size
    public double GetFileSizeKB(int intLength)
    {
        int intKbSize = 1024;
        return intLength / intKbSize;
    }
    //---------------------------------------------------------------------------
    //���o��쬰MB���ɮ�Size
    public double GetFileSizeMB(int intLength)
    {
        int intMbSize = 1048576;
        return intLength / intMbSize;
    }
    //---------------------------------------------------------------------------
    //�B�z�W���ɮ�
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
        //�W�ǭ�ۤ�
        FileUpload1.SaveAs(UploadFilePath);

        if (Util.GetControlValue("resize") == "Y")
        {
            //�Y��
            string strThumNailPath = UploadFileFolderPath + "S_{" + DateTime.Now.Ticks.ToString() + "}_" + filePath.Substring(filePath.LastIndexOf(@"\") + 1);
            if (Util.GetControlValue("img_width").Trim() != "" && Util.GetControlValue("img_height").Trim() != "")//���w���e�Y��,�i���ܧ�
            {
                MakeThumNail(UploadFilePath, strThumNailPath, Convert.ToUInt16(Util.GetControlValue("img_width").Trim()), Convert.ToUInt16(Util.GetControlValue("img_height").Trim()), "HW");
            }
            else if (Util.GetControlValue("img_width").Trim() != "" && Util.GetControlValue("img_height").Trim() == "")//���w�e��,���׫��Ӥ���Y��
            {
                MakeThumNail(UploadFilePath, strThumNailPath, Convert.ToUInt16(Util.GetControlValue("img_width").Trim()), 0, "W");
            }
            else if (Util.GetControlValue("img_width").Trim() == "" && Util.GetControlValue("img_height").Trim() != "")//���w����,�e�׫��ӵ�����Y��
            {
                MakeThumNail(UploadFilePath, strThumNailPath, 0, Convert.ToUInt16(Util.GetControlValue("img_height").Trim()), "H");
            }            
            //�R����ۤ�
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
        if (strfileName.Contains(" ") || strfileName.Contains("�@"))
        {
            //this.RegisterStartupScript("js", @"<script language='javascript' >alert('�ɮצW�٤��঳�ť�');</script>");
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('�ɮצW�٤��঳�ť�');</script>");
            return false;
        }
        return true;
    }
    //---------------------------------------------------------------------------
    public static void MakeThumNail(string originalImagePath, string thumNailPath, int width, int height, string model)
    {
        System.Drawing.Image originalImage = System.Drawing.Image.FromFile(originalImagePath);

        int thumWidth = width;      //�Y���Ϫ��e��
        int thumHeight = height;    //�Y���Ϫ�����

        int x = 0;
        int y = 0;

        int originalWidth = originalImage.Width;    //��l�Ϥ����e��
        int originalHeight = originalImage.Height;  //��l�Ϥ�������

        switch (model)
        {
            case "HW":      //���w���e�Y��,�i���ܧ�
                break;
            case "W":       //���w�e��,���׫��Ӥ���Y��
                thumHeight = originalImage.Height * width / originalImage.Width;
                break;
            case "H":       //���w����,�e�׫��ӵ�����Y��
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

        //�s�ؤ@��bmp�Ϥ�
        System.Drawing.Image bitmap = new System.Drawing.Bitmap(thumWidth, thumHeight);

        //�s�ؤ@�ӵe�O
        System.Drawing.Graphics graphic = System.Drawing.Graphics.FromImage(bitmap);

        //�]�m����q�d�Ȫk
        graphic.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;

        //�]�m����q�A�C�t�קe�{���Ƶ{��
        graphic.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

        //�M�ŵe���åH�z���I�����R
        graphic.Clear(System.Drawing.Color.Transparent);

        //�b���w��m�åB�����w�j�pø�s��Ϥ������w����
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
