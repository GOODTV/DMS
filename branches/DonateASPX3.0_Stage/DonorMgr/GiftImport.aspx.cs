using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using System.Drawing;

public partial class DonorMgr_GiftImport : BasePage
{
    HSSFWorkbook workbook = null;
    HSSFSheet u_sheet = null;
    //要上傳Excel檔的Server端 檔案總管目錄
    //string upload_excel_Dir = @"D:\";
    
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "GiftImport";
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            LoadDropDownListData();
        }
        lblGridList.Font.Size = 16;
        lblGridList.ForeColor = Color.Blue;
    }
    //---------------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構
        Util.FillDropDownList(ddlDept, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept.SelectedIndex = 0;
    }
        
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

        string excel_filePath = "";
        try
        {
            excel_filePath = SaveFileAndReturnPath();//先上傳EXCEL檔案給Server

            this.workbook = new HSSFWorkbook(FileUpload.FileContent);
            this.u_sheet = (HSSFSheet)workbook.GetSheetAt(0);  //取得第0個Sheet
            //不同於Microsoft Object Model，NPOI都是從索引0開始算起

            //從第一個Worksheet讀資料&Insert into GiftTemp  
            SaveOrInsertSheet(this.u_sheet);

            //ClientScript.RegisterClientScriptBlock(typeof(System.Web.UI.Page), "匯入完成", "alert('匯入完成');", true);

        }
        catch (Exception ex)
        {
            ClientScript.RegisterClientScriptBlock(typeof(System.Web.UI.Page), "匯入失敗！", "alert('匯入失敗！');", true);
            ShowSysMsg(ex.Message.ToString());
        }
        finally
        {

            //釋放 NPOI的資源
            if (this.workbook != null) this.workbook = null;
            if (this.u_sheet != null) this.u_sheet = null;


            //是否刪除Server上的Excel檔(預設true)
            bool isDeleteFileFromServer = true;
            if (isDeleteFileFromServer)
            {
                System.IO.File.Delete(excel_filePath);
            }


            GC.Collect();
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
        int DenyMbSize = 10;//限制大小10MB
        bool flag = false;
        //判斷副檔名
        switch (System.IO.Path.GetExtension(FileUpload.PostedFile.FileName).ToUpper())
        {
            case ".XLS":
            case ".XLSX":
                flag = true;
                break;
            default:
                this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('此檔案類型不允許上傳');</script>");
                flag = false;
                break;
        }
        if (Util.GetFileSizeMB(FileUpload.FileBytes.Length) > DenyMbSize)
        {
            this.ClientScript.RegisterStartupScript(this.GetType(), "js", @"<script language='javascript' >alert('檔案大小不得超過" + DenyMbSize + @"MB');</script>");
            flag = false;
        }
        return flag;
    }
    //儲存EXCEL檔案給Server
    private string SaveFileAndReturnPath()
    {
        string return_file_path = "";//上傳的Excel檔在Server上的位置
        string ImportFileName = string.Empty;
        string ImportSource = string.Empty;
        string appPath = Request.PhysicalApplicationPath;
        string saveDir = "\\UpLoad\\";

        if (FileUpload.FileName != "")
        {
            //return_file_path = System.IO.Path.Combine(this.upload_excel_Dir, DateTime.Now.ToString("yyyyMMddHHmmss") + "公關贈品批次.xls");
            ImportFileName = this.FileUpload.FileName;
            //改存檔檔名為日期+上傳檔名
            return_file_path = appPath + saveDir + DateTime.Now.ToString("yyyyMMddHHmmss") + "公關贈品批次.xls";

            FileUpload.SaveAs(return_file_path);
        }
        return return_file_path;

    }
    private void SaveOrInsertSheet(HSSFSheet u_sheet)
    {
        bool flag = false;
        //先刪除暫存table資料後 再新增
        string strSqlDel = "delete from GiftTemp";
        NpoDB.ExecuteSQLS(strSqlDel, null);
        int i ;
        string Donor_Id, Gift_Date, Goods_Name, Goods_Qty = "";
        //因為要讀取的資料列不包含標頭，所以i從u_sheet.FirstRowNum + 1開始讀
        for (i = u_sheet.FirstRowNum + 1; i <= u_sheet.LastRowNum; i++)/*一列一列地讀取資料*/
        {
            HSSFRow row = (HSSFRow)u_sheet.GetRow(i);//取得目前的資料列
            if (row.Cells.Count != 4)//資料列有缺漏
            {
                flag = false;
                //欄位有問題刪除之前新增的
                strSqlDel = "delete from GiftTemp";
                NpoDB.ExecuteSQLS(strSqlDel, null);
                //ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"請再次檢查匯入資料是否有誤！\");</script>");
                lblGridList.Text = "請再次檢查匯入資料是否有誤！";
            }
            else
            {
                //再對各Cell處理完商業邏輯後，Insert into Table
                //目前資料列第1格的值
                Donor_Id = row.GetCell(0).ToString() != null ? row.GetCell(0).ToString() : "";
                //目前資料列第2格的值
                Gift_Date = row.GetCell(1).ToString() != null ? row.GetCell(1).ToString() : "";
                //目前資料列第3格的值
                Goods_Name = row.GetCell(2).ToString() != null ? row.GetCell(2).ToString() : "";
                //目前資料列第4格的值
                Goods_Qty = row.GetCell(3).ToString() != null ? row.GetCell(3).ToString() : "";
                string strSql = "select  Ser_No from Linked2 where Linked2_Name= '" + Goods_Name + "'";
                System.Data.DataTable dt = NpoDB.QueryGetTable(strSql);
                if (dt.Rows.Count == 0)
                {
                    //ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"請先新增公關贈品品項！\");</script>");
                    lblGridList.Text = "請先新增公關贈品品項！";
                }
                else //Insert To GiftTemp
                {
                    string strSqlTemp = "insert into  GiftTemp\n";
                    strSqlTemp += "( Donor_Id, Gift_Date, Goods_Name ,Goods_Qty) values \n";
                    strSqlTemp += "( @Donor_Id,@Gift_Date,@Goods_Name,@Goods_Qty)";

                    Dictionary<string, object> dict = new Dictionary<string, object>();
                    dict.Add("Donor_Id", Donor_Id);
                    dict.Add("Gift_Date", Gift_Date);
                    dict.Add("Goods_Name", Goods_Name);
                    dict.Add("Goods_Qty", Goods_Qty);
                    NpoDB.ExecuteSQLS(strSqlTemp, dict);
                    flag = true;
                }
            }
        }
        if (flag == true)
        {
            //ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"公關贈品資料匯入共" + String.Format("{0:#,0}", i-1) + "筆。\");</script>");
            lblGridList.Text = "公關贈品資料匯入共 " + String.Format("{0:#,0}", i - 1) + " 筆。";
            //Insert Gift from GiftTemp
            InsertToGiftdata();
        }
    }
    //private void InsertToGiftTemp(string excel_filename)
    //{
    //    string ProviderName = "Microsoft.Jet.OLEDB.4.0;";
    //    string ExtendedString = "'Excel 8.0;";
    //    string Hdr = "Yes;";
    //    string IMEX = "1';";

    //    string cs = "Data Source=" + excel_filename + ";" +
    //                "Provider=" + ProviderName +
    //                "Extended Properties=" + ExtendedString +
    //                "HDR=" + Hdr +
    //                "IMEX=" + IMEX;

    //    // 用來存放上傳MOD檔案所需要的資料內容
    //    ArrayList ModDataList = new ArrayList();

    //    // 用來檢核Planning Title 第二碼是英文或數字  如果是數字抓七碼 如果不是數字抓八碼 (節目名稱)
    //    string strValue = @"^\d+(\.)?\d*$"; //數字
    //    System.Text.RegularExpressions.Regex r = new System.Text.RegularExpressions.Regex(strValue);
        
    //    using (OleDbConnection cn = new OleDbConnection(cs))
    //    {
    //        cn.Open();
    //        string qs = "Select * From [Sheet1$]";
    //        //先刪除暫存table資料後 再新增
    //        string strSqlDel = "delete from GiftTemp";
    //        NpoDB.ExecuteSQLS(strSqlDel, null);
    //        try
    //        {
    //            using (OleDbCommand cmd = new OleDbCommand(qs, cn))
    //            {
    //                using (OleDbDataReader dr = cmd.ExecuteReader())
    //                {
    //                    int i = 0;
    //                    while (dr.Read())
    //                    {
    //                        //int Col = dr.FieldCount;
    //                        i++;

    //                        //Donor_Id
    //                        string Donor_Id = dr["捐款人編號"] != null ? dr["捐款人編號"].ToString() : "";
    //                        //Gift_Date
    //                        string Gift_Date = dr["贈與日期"] != null ? dr["贈與日期"].ToString() : "";
    //                        //Goods_Name	
    //                        string Goods_Name = dr["品項"] != null ? dr["品項"].ToString() : "";
    //                        //Goods_Qty
    //                        string Goods_Qty = dr["數量"] != null ? dr["數量"].ToString() : "";

    //                        string strSql = "select  Ser_No from Linked2 where Linked2_Name= '" + Goods_Name + "'";
    //                        System.Data.DataTable dt = NpoDB.QueryGetTable(strSql);
    //                        if (dt.Rows.Count == 0)
    //                        {
    //                            ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"請先新增公關贈品品項！\");</script>");
    //                        }
    //                        else //Insert To GiftTemp
    //                        {
    //                            string strSqlTemp = "insert into  GiftTemp\n";
    //                            strSqlTemp += "( Donor_Id, Gift_Date, Goods_Name ,Goods_Qty) values \n";
    //                            strSqlTemp += "( @Donor_Id,@Gift_Date,@Goods_Name,@Goods_Qty)";

    //                            Dictionary<string, object> dict = new Dictionary<string, object>();
    //                            dict.Add("Donor_Id", Donor_Id);
    //                            dict.Add("Gift_Date", Gift_Date);
    //                            dict.Add("Goods_Name", Goods_Name);
    //                            dict.Add("Goods_Qty", Goods_Qty);
    //                            NpoDB.ExecuteSQLS(strSqlTemp, dict);
    //                            flag = true;
    //                            //GiftTemp_AddNew(Donor_Id, Gift_Date, Goods_Name, Goods_Qty);
    //                        }
     
    //                    }
    //                    if (flag == true)
    //                    {
    //                        ClientScript.RegisterStartupScript(this.GetType(), "js", "<script>alert(\"公關贈品資料匯入共" + String.Format("{0:#,0}", i) + "筆。\");</script>");
    //                        //Insert Gift from GiftTemp
    //                        InsertToGiftdata();
    //                    }
    //                }
                    
    //            }
                

    //        }
    //        catch (Exception ex)
    //        {
    //            ShowSysMsg("公關贈品匯入暫存失敗！");
    //            ShowSysMsg(ex.Message.ToString());
    //            //this.logger.Error(ex.Message, ex);

    //        }

    //    }

    //}
    public void InsertToGiftdata()
    {
        string strSql, strSqlGift, strSqlGiftData = "";
        strSql = "select Donor_Id,CONVERT(VarChar,Gift_Date,111) as Gift_Date,Goods_Name,Goods_Qty from  GiftTemp";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt;
        dt = NpoDB.GetDataTableS(strSql, dict);
        if (dt.Rows.Count != 0)
        {
            int intCount = Convert.ToInt32(dt.Rows.Count);
            foreach (DataRow dr in dt.Rows)
            {
                string Donor_Id = dr["Donor_Id"].ToString();
                string Gift_Date = dr["Gift_Date"].ToString();
                string Goods_Name = dr["Goods_Name"].ToString();
                string Goods_Qty = dr["Goods_Qty"].ToString();
                try
                {
                    //新增至Gift
                    strSqlGift = "insert into  Gift\n";
                    strSqlGift += "( Donor_Id, Gift_Date, Gift_Payment, Dept_Id \n";
                    strSqlGift += " , Act_id, Comment \n";
                    strSqlGift += " , Create_Date, Create_User, Create_IP) values\n";
                    strSqlGift += "( @Donor_Id,@Gift_Date,@Gift_Payment,@Dept_Id \n";
                    strSqlGift += " ,@Act_id,@Comment \n";
                    strSqlGift += " ,@Create_Date,@Create_User,@Create_IP)";
                    strSqlGift += "\n";
                    strSqlGift += "select @@IDENTITY";

                    Dictionary<string, object> dictGift = new Dictionary<string, object>();
                    dictGift.Add("Donor_Id", Donor_Id);
                    dictGift.Add("Gift_Date", Gift_Date);
                    dictGift.Add("Gift_Payment", "公關贈品");
                    dictGift.Add("Dept_Id", SessionInfo.DeptID);
                    dictGift.Add("Act_id", "");
                    dictGift.Add("Comment", "批次匯入");
                    dictGift.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    dictGift.Add("Create_User", SessionInfo.UserName);
                    dictGift.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
                    NpoDB.ExecuteSQLS(strSqlGift, dictGift);
                    
                    //找出對應的Gift_Id，再新增進GiftData
                    string strSqlGiftId = "Select Max(Gift_Id) Gift_Id From GIFT Where Donor_Id = '" + Donor_Id + "' And Gift_Date = '" + Gift_Date + "'";
                    DataTable dt2 = NpoDB.QueryGetTable(strSqlGiftId);
                    //找對應Goods_Name的Goods_Id，再新增進GiftData
                    string strSql_GoodsId = "select  Ser_No from Linked2 where Linked2_Name= '" + Goods_Name + "'";
                    DataTable dt3 = NpoDB.QueryGetTable(strSql_GoodsId);

                    strSqlGiftData = "insert into GiftData\n";
                    strSqlGiftData += "( Gift_Id, Donor_Id, Goods_Id, Goods_Name, Goods_Qty, \n";
                    strSqlGiftData += " Create_Date, Create_User,Create_IP)values \n";
                    strSqlGiftData += "( @Gift_Id,@Donor_Id,@Goods_Id,@Goods_Name,@Goods_Qty, \n";
                    strSqlGiftData += " @Create_Date, @Create_User,@Create_IP) \n";
                    strSqlGiftData += "select @@IDENTITY";

                    Dictionary<string, object> dictGiftData = new Dictionary<string, object>();
                    dictGiftData.Add("Gift_Id", dt2.Rows[0]["Gift_Id"].ToString());
                    dictGiftData.Add("Donor_Id", Donor_Id);
                    dictGiftData.Add("Goods_Id", dt3.Rows[0]["Ser_No"].ToString());
                    dictGiftData.Add("Goods_Name", Goods_Name);
                    if (string.IsNullOrEmpty(Goods_Qty) || Goods_Qty == "")
                        dictGiftData.Add("Goods_Qty", "1");
                    else
                        dictGiftData.Add("Goods_Qty", Goods_Qty);
                    dictGiftData.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    dictGiftData.Add("Create_User", SessionInfo.UserName);
                    dictGiftData.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
                    NpoDB.ExecuteSQLS(strSqlGiftData, dictGiftData);
                    
                }
                catch (Exception ex)
                {
                    ShowSysMsg("公關贈品匯入失敗！");
                    ShowSysMsg(ex.Message.ToString());
                }
            }
        }
    }
}