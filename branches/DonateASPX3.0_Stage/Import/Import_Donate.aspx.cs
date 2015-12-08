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

public partial class Import_Import_Donate : BasePage
{
    Microsoft.Office.Interop.Excel.Application xlApp = null;
    Workbook wb = null;
    Worksheet ws = null;
    Range aRange = null;
    //*******************************/
    //要上傳Excel檔的Server端 檔案總管目錄
    string upload_excel_Dir = @"D:\";

    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "Import_Donate";
        if (!IsPostBack)
        {
            Util.ShowSysMsgWithScript("");
            LoadDropDownListData();
        }
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

            if (this.xlApp == null)
            {
                this.xlApp = new Microsoft.Office.Interop.Excel.Application();
            }
            //打開Server上的Excel檔案
            this.xlApp.Workbooks.Open(excel_filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            this.wb = xlApp.Workbooks[1];//第一個Workbook
            this.wb.Save();

            //引用活頁簿類別
            int count = wb.Worksheets.Count;
            bool getSheet = false;
            for (int i = 1; i <= count;i++ )
            {
                ws = (Worksheet)xlApp.Worksheets[i];
                if (ws.Name == tbxSheet_Name.Text)
                {
                    SaveOrInsertSheet(excel_filePath, (Worksheet)xlApp.Worksheets[i]);
                    getSheet = true;
                }
            }
            if (getSheet == false)
            {
                ShowSysMsg("沒有符合的工作表名稱");
                return;
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            xlApp.Workbooks.Close();
            xlApp.Quit();
            try
            {
                //刪除 Windows工作管理員中的Excel.exe 處理緒.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.xlApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.ws);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.aRange);
            }
            catch { }
            this.xlApp = null;
            this.wb = null;
            this.ws = null;
            this.aRange = null;

            //是否刪除Server上的Excel檔
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
        if (FileUpload.FileName != "")
        {
            return_file_path = System.IO.Path.Combine(this.upload_excel_Dir, Guid.NewGuid().ToString() + ".xls");

            FileUpload.SaveAs(return_file_path);
        }
        return return_file_path;
    }
    //把Excel資料Insert into Table
    private void SaveOrInsertSheet(string excel_filename, Worksheet ws)
    {

        //要開始讀取的起始列(微軟Worksheet是從1開始算)
        int rowIndex = 2;

        //計算筆數、總金額
        int count = 0;
        int Amount_Total = 0;
        //取得一列的範圍
        this.aRange = ws.get_Range("A" + rowIndex.ToString(), "M" + rowIndex.ToString());

        //判斷Row範圍裡第1格有值的話，迴圈就往下跑
        while (((object[,])this.aRange.Value2)[1, 1] != null)//用this.aRange.Cells[1, 1]來取值的方式似乎會造成無窮迴圈？
        {
            //範圍裡第1格的值
            string Donor_Name = ((object[,])this.aRange.Value2)[1, 1] != null ? ((object[,])this.aRange.Value2)[1, 1].ToString() : "";
            //範圍裡第2格的值
            string IDNo = ((object[,])this.aRange.Value2)[1, 2] != null ? ((object[,])this.aRange.Value2)[1, 2].ToString() : "";
            //範圍裡第3格的值
            string Cellular_Phone = ((object[,])this.aRange.Value2)[1, 3] != null ? ((object[,])this.aRange.Value2)[1, 3].ToString() : "";
            if (Cellular_Phone.Trim().Length == 9 && Cellular_Phone.Trim().Substring(0, 1) == "9")
            {
                Cellular_Phone = "0" + Cellular_Phone.Trim();
            }
            //範圍裡第4格的值
            string Tel_Office = ((object[,])this.aRange.Value2)[1, 4] != null ? ((object[,])this.aRange.Value2)[1, 4].ToString() : "";
            //範圍裡第5格的值
            string Email = ((object[,])this.aRange.Value2)[1, 5] != null ? ((object[,])this.aRange.Value2)[1, 5].ToString() : "";
            //範圍裡第6格的值
            string ZipCode = ((object[,])this.aRange.Value2)[1, 6] != null ? ((object[,])this.aRange.Value2)[1, 6].ToString() : "";
            //範圍裡第7格的值
            string Address_temp = ((object[,])this.aRange.Value2)[1, 7] != null ? ((object[,])this.aRange.Value2)[1, 7].ToString() : "";
            Address_temp = Address_temp.Replace("臺", "台");
            Address_temp = Address_temp.Replace("台北縣", "新北市");
            Address_temp = Address_temp.Replace("台中縣", "台中市");
            Address_temp = Address_temp.Replace("台南縣", "台南市");
            Address_temp = Address_temp.Replace("高雄縣", "高雄市");
            //抓出縣市*****************
            string City = Address_temp.Substring(0, 3);//縣市
            string Area = "";
            string Address_temp2 = Address_temp.Substring(3);//沒有了縣市的地址
            //特殊的三個鄉鎮市區
            if(Address_temp.IndexOf("新竹市")>0)
            {
                if (Address_temp2.IndexOf("北區") > 0)
                {
                    ZipCode = "3001";
                }
                else if (Address_temp2.IndexOf("香山區") > 0)
                {
                    ZipCode = "3002";
                }
            }
            else if (Address_temp.IndexOf("嘉義市") > 0)
            {
                if (Address_temp2.IndexOf("西區") > 0)
                {
                    ZipCode = "6001";
                }
            }

            if (ZipCode != "")
            {
                Area = Util.GetAreaName(ZipCode);
                //沒有縣市和鄉鎮市區的地址*****************
            }
            string Address = Address_temp2.Replace(Area, "");

            //範圍裡第8格的值
            string Donate_Date = "";
            string Donate_Date_temp = ((object[,])this.aRange.Value2)[1, 8] != null ? ((object[,])this.aRange.Value2)[1, 8].ToString() : "";
            if (Donate_Date_temp.Length == 5)
            {
                Donate_Date_temp = Util.getDateStr(Donate_Date_temp);
            }
            if (Donate_Date_temp.Substring(0, 1) == "1")
            {
                if (Donate_Date_temp.Trim().Length == 7 && Donate_Date_temp.Trim().Substring(3, 1) != "/")//ex:1020101
                {
                    Donate_Date = (int.Parse(Donate_Date_temp.Substring(0, 3)) + 1911).ToString() + "/" + Donate_Date_temp.Substring(3, 2) + "/" + Donate_Date_temp.Substring(5, 2);
                }
                else if (Donate_Date_temp.Trim().Length >= 7 && Donate_Date_temp.Trim().Substring(3, 1) == "/")//ex:102/1/1
                {
                    Donate_Date = Donate_Date_temp.Replace(Donate_Date_temp.Substring(0, 3), (int.Parse(Donate_Date_temp.Substring(0, 3)) + 1911).ToString());
                }
            }
            else if (Donate_Date_temp.Substring(0, 1) == "2")
            {
                if (Donate_Date_temp.Trim().Length == 8 && Donate_Date_temp.Trim().Substring(4, 1) != "/") //ex:20130101
                {
                    Donate_Date = Donate_Date_temp.Substring(0, 4) + "/" + Donate_Date_temp.Substring(4, 2) + "/" + Donate_Date_temp.Substring(6, 2);
                }
                else if (Donate_Date_temp.Trim().Length >= 8 && Donate_Date_temp.Trim().Substring(4, 1) == "/")//ex:2013/01/01
                {
                    Donate_Date = Donate_Date_temp.Trim();
                }
            }
            //範圍裡第9格的值
            string Donate_Payment = ((object[,])this.aRange.Value2)[1, 9] != null ? ((object[,])this.aRange.Value2)[1, 9].ToString() : "";
            //範圍裡第10格的值
            string Donate_Purpose = ((object[,])this.aRange.Value2)[1, 10] != null ? ((object[,])this.aRange.Value2)[1, 10].ToString() : "";
            //範圍裡第11格的值
            string Donate_Amt = ((object[,])this.aRange.Value2)[1, 11] != null ? ((object[,])this.aRange.Value2)[1, 11].ToString() : "";
            //範圍裡第12格的值
            string Invoice_Type = ((object[,])this.aRange.Value2)[1, 12] != null ? ((object[,])this.aRange.Value2)[1, 12].ToString() : "";
            //範圍裡第13格的值
            string Invoice_Title = ((object[,])this.aRange.Value2)[1, 13] != null ? ((object[,])this.aRange.Value2)[1, 13].ToString() : "";
            if (Invoice_Title.Trim() == "")
            {
                Invoice_Title = Donor_Name.Trim();
            }

            //捐款人資料(DONOR)
            string strSql = "Select Donor_Id From DONOR Where Donor_Name='" + Donor_Name + "' and DeleteDate is null ";
            string WhereSQL = "";
            if (IDNo != "")
            {
                if (WhereSQL == "")
                {
                    WhereSQL += "And (IDNo='" + IDNo + "' ";
                }
                else
                {
                    WhereSQL += "Or IDNo='" + IDNo + "' ";
                }
            }
            if (Cellular_Phone != "")
            {
                if (WhereSQL == "")
                {
                    WhereSQL += "And (Cellular_Phone='" + Cellular_Phone + "' ";
                }
                else
                {
                    WhereSQL += "Or Cellular_Phone='" + Cellular_Phone + "' ";
                }
            }
            if (Tel_Office != "")
            {
                if (WhereSQL == "")
                {
                    WhereSQL += "And (Tel_Office='" + Tel_Office + "' or Tel_Home='" + Tel_Office + "' ";
                }
                else
                {
                    WhereSQL += "Or Tel_Office='" + Tel_Office + "' or Tel_Home='" + Tel_Office + "' ";
                }
            }
            if (ZipCode != "")
            {
                if (WhereSQL == "")
                {
                    WhereSQL += "And (ZipCode='" + ZipCode + "' or Invoice_ZipCode='" + ZipCode + "' ";
                }
                else
                {
                    WhereSQL += "Or ZipCode='" + ZipCode + "' or Invoice_ZipCode='" + ZipCode + "' ";
                }
            }
            if (Address != "")
            {
                if (WhereSQL == "")
                {
                    WhereSQL += "And (Address='" + Address + "' or Invoice_Address='" + Address + "' ";
                }
                else
                {
                    WhereSQL += "Or Address='" + Address + "' or Invoice_Address='" + Address + "' ";
                }
            }
            if (Email != "")
            {
                if (WhereSQL == "")
                {
                    WhereSQL += "And (Email='" + Email + "' ";
                }
                else
                {
                    WhereSQL += "Or Email='" + Email + "' ";
                }
            }
            if (WhereSQL != "")
            {
                WhereSQL += ")";
                strSql += WhereSQL;
            }
            System.Data.DataTable dt = NpoDB.QueryGetTable(strSql);
            if (dt.Rows.Count == 0)//新增此捐款人的資料
            {
                Donor_AddNew(Donor_Name, IDNo, Cellular_Phone, Tel_Office, Email, ZipCode, City, Area, Address, Invoice_Type, Invoice_Title);
                dt = NpoDB.QueryGetTable(strSql);
            }

            DataRow dr = dt.Rows[0];
            HFD_Donor_Id.Value = dr["Donor_Id"].ToString();

            //新增此次捐款資料
            Donate_AddNew(Donate_Date, Donate_Payment, Donate_Purpose, Donate_Amt, Invoice_Type, Invoice_Title);
            //修改捐款人的捐款資料
            Donor_Edit(Donate_Amt);
            count+=1;
            Amount_Total += int.Parse(Donate_Amt);
            //往下抓一列Excel範圍
            rowIndex++;
            this.aRange = ws.get_Range("A" + rowIndex.ToString(), "M" + rowIndex.ToString());
        }
        string alert = "alert('捐款資料匯入成功 ！共計：" + count + "筆，金額：" + Amount_Total + "元');</script>";
        this.ClientScript.RegisterStartupScript(this.GetType(), "js", "<script language='javascript'>" + alert);
    }
    //---------------------------------------------------------------------------
    public void Donor_AddNew(string Donor_Name,string IDNo,string Cellular_Phone,string Tel_Office,string Email,string ZipCode,string City,string Area,string Address,string Invoice_Type,string Invoice_Title)
    {
        string strSql = "insert into  DONOR\n";
        strSql += "( Donor_Name, Category, Sex, Title, Donor_Type, IDNo, Birthday, Cellular_Phone, Tel_Office, Fax, Email,\n";
        strSql += " Contactor, JobTitle, IsAbroad, ZipCode, City, Area, Address, IsSendNews, IsSendEpaper, IsSendYNews, \n";
        strSql += " IsBirthday, IsXmas, Invoice_Type, IsAnonymous, NickName, Title2, Invoice_Title, \n";
        strSql += " Invoice_IDNo, IsAbroad_Invoice, Report_Type, Invoice_Title_Man, Remark , IsFdc, \n";
        strSql += " Street,Section,Lane,Alley,HouseNo,HouseNoSub,Floor,FloorSub,Room,OverseasCountry,\n";
        strSql += " OverseasAddress,Invoice_Street,Invoice_Section,Invoice_Lane,Invoice_Alley,Invoice_HouseNo,\n";
        strSql += " Invoice_HouseNoSub,Invoice_Floor,Invoice_FloorSub,Invoice_Room,Invoice_OverseasCountry,\n";
        strSql += " Invoice_OverseasAddress,Invoice_ZipCode,Invoice_City,Invoice_Area,Invoice_Address, Dept_Id, IsMember, Create_Date, Create_DateTime, Create_User, Create_IP) values\n";
        strSql += "(@Donor_Name,@Category,@Sex,@Title,@Donor_Type,@IDNo,@Birthday,@Cellular_Phone,@Tel_Office,@Fax,@Email,\n";
        strSql += "@Contactor,@JobTitle,@IsAbroad,@ZipCode,@City,@Area,@Address,@IsSendNews,@IsSendEpaper,@IsSendYNews, \n";
        strSql += "@IsBirthday,@IsXmas,@Invoice_Type,@IsAnonymous,@NickName,@Title2,@Invoice_Title,@Invoice_IDNo,@IsAbroad_Invoice, \n";
        strSql += "@Report_Type,@Invoice_Title_Man,@Remark,@IsFdc,@Street,@Section,@Lane,@Alley,@HouseNo,@HouseNoSub,@Floor, \n";
        strSql += "@FloorSub,@Room,@OverseasCountry,@OverseasAddress,@Invoice_Street,@Invoice_Section,@Invoice_Lane,@Invoice_Alley, \n";
        strSql += "@Invoice_HouseNo,@Invoice_HouseNoSub,@Invoice_Floor,@Invoice_FloorSub,@Invoice_Room,@Invoice_OverseasCountry,@Invoice_OverseasAddress, \n";
        strSql += "@Invoice_ZipCode,@Invoice_City,@Invoice_Area,@Invoice_Address,@Dept_Id,@IsMember,@Create_Date,@Create_DateTime,@Create_User,@Create_IP)";
        strSql += "\n";
        strSql += "select @@IDENTITY";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Name", Donor_Name.Trim());
        if (IDNo.Length == 8)
        {
            dict.Add("Category", "團體");
            dict.Add("Sex", "");
        }
        else if (IDNo.Length!=0)
        {
            dict.Add("Category", "個人");
            if (IDNo.Substring(1, 1) == "1")
            {
                dict.Add("Sex", "男");
            }
            else if (IDNo.Substring(1, 1) == "2")
            {
                dict.Add("Sex", "女");
            }
            else
            {
                dict.Add("Sex", "");
            }
        }
        else 
        {
            dict.Add("Category", "個人");
            dict.Add("Sex", "");
        }
        dict.Add("Title", "君");
        dict.Add("Donor_Type", "");
        dict.Add("IDNo", IDNo.Trim());
        dict.Add("Birthday", null);
        dict.Add("Cellular_Phone", Cellular_Phone.Trim());
        dict.Add("Tel_Office", Tel_Office.Trim());
        dict.Add("Fax", "");
        dict.Add("Email", Email.Trim());
        dict.Add("Contactor", "");
        dict.Add("JobTitle", "");
        dict.Add("ZipCode", ZipCode.Trim());
        dict.Add("City", Util.GetCityCode(City.Trim(), "0"));
        dict.Add("Area", Util.GetAreaCode(Area.Trim()));
        dict.Add("IsAbroad", "N");
        dict.Add("Street", "");
        dict.Add("Section", "");
        dict.Add("Lane", "");
        dict.Add("Alley", "");
        dict.Add("HouseNo", "");
        dict.Add("HouseNoSub", "");
        dict.Add("Floor", "");
        dict.Add("FloorSub", "");
        dict.Add("Room", "");
        dict.Add("Address", Address.Trim());
        dict.Add("OverseasCountry", "");
        dict.Add("OverseasAddress", "");
        dict.Add("IsSendNews", "N");
        dict.Add("IsSendEpaper", "N");
        dict.Add("IsSendYNews", "N");
        dict.Add("IsBirthday", "N");
        dict.Add("IsXmas", "N");
        dict.Add("IsMember", "N");
        dict.Add("Dept_Id", ddlDept.SelectedValue);
        if (Invoice_Type.Trim() == "不需要" || Invoice_Type.Trim() == "年度證明及收據" || Invoice_Type.Trim() == "國外格式(A)" || Invoice_Type.Trim() == "國外格式(B)")
        {
            dict.Add("Invoice_Type", Invoice_Type);
        }
        else
        {
            dict.Add("Invoice_Type", "");
        }
        dict.Add("IsAnonymous", "N");
        dict.Add("IsAbroad_Invoice", "N");
        dict.Add("Invoice_ZipCode", ZipCode.Trim());
        dict.Add("Invoice_City", Util.GetCityCode(City.Trim(), "0"));
        dict.Add("Invoice_Area", Util.GetAreaCode(Area.Trim()));
        dict.Add("Invoice_Street", "");
        dict.Add("Invoice_Section", "");
        dict.Add("Invoice_Lane", "");
        dict.Add("Invoice_Alley", "");
        dict.Add("Invoice_HouseNo", "");
        dict.Add("Invoice_HouseNoSub", "");
        dict.Add("Invoice_Floor", "");
        dict.Add("Invoice_FloorSub", "");
        dict.Add("Invoice_Room", "");
        dict.Add("Invoice_Address", Address.Trim());
        dict.Add("Invoice_OverseasCountry", "");
        dict.Add("Invoice_OverseasAddress", "");
        dict.Add("NickName", "");
        dict.Add("Title2", "");
        dict.Add("Invoice_Title", Invoice_Title.Trim());
        dict.Add("Invoice_IDNo", IDNo.Trim());
        dict.Add("Report_Type", "");
        dict.Add("Invoice_Title_Man", "");
        dict.Add("Remark", "");
        dict.Add("IsFdc", "N");
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        NpoDB.ExecuteSQLS(strSql, dict); 
    }
    public void Donate_AddNew(string Donate_Date, string Donate_Payment, string Donate_Purpose, string Donate_Amt, string Invoice_Type, string Invoice_Title)
    {
        string strSql = "insert into  Donate\n";
        strSql += "( Donor_Id, Donate_Date, Donate_Payment, Donate_Purpose_Type, Donate_Purpose, Invoice_Type, Donate_Amt, Donate_Fee, Donate_Accou,\n";
        strSql += " Donate_Forign, Dept_Id, Invoice_Title, Invoice_Pre, Invoice_No, Request_Date, Accoun_Bank, Accoun_Date,\n";
        strSql += " Donate_Type, Donation_NumberNo, Donation_SubPoenaNo, Accounting_Title, InvoceSend_Date, Act_Id,\n";
        strSql += " Comment, Invoice_PrintComment, Export, Create_Date, Create_DateTime, Create_User, Create_IP,\n";
        strSql += " Check_No, Check_ExpireDate, Card_Bank, Card_Type, Account_No, Valid_Date, Card_Owner, Owner_IDNo, Relation, \n";
        strSql += " Authorize, Post_Name, Post_IDNo, Post_SavingsNo, Post_AccountNo, Excel_Type) values\n";
        strSql += "( @Donor_Id,@Donate_Date,@Donate_Payment,@Donate_Purpose_Type,@Donate_Purpose,@Invoice_Type,@Donate_Amt,@Donate_Fee,@Donate_Accou,\n";
        strSql += "@Donate_Forign,@Dept_Id,@Invoice_Title,@Invoice_Pre,@Invoice_No,@Request_Date,@Accoun_Bank,@Accoun_Date,\n";
        strSql += "@Donate_Type,@Donation_NumberNo,@Donation_SubPoenaNo,@Accounting_Title,@InvoceSend_Date,@Act_Id,\n";
        strSql += "@Comment,@Invoice_PrintComment,@Export,@Create_Date,@Create_DateTime,@Create_User,@Create_IP,\n";
        strSql += "@Check_No,@Check_ExpireDate,@Card_Bank,@Card_Type,@Account_No,@Valid_Date,@Card_Owner,@Owner_IDNo,@Relation, \n";
        strSql += "@Authorize,@Post_Name,@Post_IDNo,@Post_SavingsNo,@Post_AccountNo,@Excel_Type) ";
        strSql += "\n";
        strSql += "select @@IDENTITY";


        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", HFD_Donor_Id.Value);
        dict.Add("Donate_Date", Donate_Date.Trim());
        dict.Add("Donate_Payment", Donate_Payment.Trim());
        dict.Add("Donate_Purpose_Type", "D");
        dict.Add("Donate_Purpose", Donate_Purpose.Trim());
        dict.Add("Invoice_Type", Invoice_Type.Trim());
        dict.Add("Donate_Amt", Donate_Amt.Trim());
        dict.Add("Donate_Fee", "0");
        dict.Add("Donate_Accou", Donate_Amt.Trim());
        dict.Add("Donate_Forign", "0");

        dict.Add("Check_No", "");
        dict.Add("Check_ExpireDate", null);

        dict.Add("Card_Bank", "");
        dict.Add("Card_Type", "");
        dict.Add("Account_No", "");
        dict.Add("Valid_Date", "");
        dict.Add("Card_Owner", "");
        dict.Add("Owner_IDNo", "");
        dict.Add("Relation", "");
        dict.Add("Authorize", "");

        dict.Add("Post_Name", "");
        dict.Add("Post_IDNo", "");
        dict.Add("Post_SavingsNo", "");
        dict.Add("Post_AccountNo", "");

        dict.Add("Dept_Id", ddlDept.SelectedValue);
        dict.Add("Invoice_Title", Invoice_Title.Trim());
        //收據編號
        dict.Add("Invoice_Pre", "A");
        //******設定收據編號******//
        string strSql2 = @"select isnull(MAX(Invoice_No),'') as Invoice_No from Donate
                    where Invoice_No like '%" + DateTime.Now.ToString("yyyyMMdd") + "%'";
        //****執行查詢流水號語法****//
        System.Data.DataTable dt2 = NpoDB.QueryGetTable(strSql2);
        string Invoice_No = "";
        if (dt2.Rows.Count > 0)
        {
            string Invoice_No_value = dt2.Rows[0]["Invoice_No"].ToString();

            if (Invoice_No_value == "")
            {
                Invoice_No = DateTime.Now.ToString("yyyyMMdd") + "0001";
            }
            else
            {
                Invoice_No = (Convert.ToInt64(Invoice_No_value) + 1).ToString();
            }
        }
        //************************//
        dict.Add("Invoice_No", Invoice_No);

        dict.Add("Request_Date", null);
        dict.Add("Accoun_Bank", "");
        dict.Add("Accoun_Date", null);
        dict.Add("Donate_Type", "單次捐款");
        dict.Add("Donation_NumberNo", "");
        dict.Add("Donation_SubPoenaNo", "");
        dict.Add("Accounting_Title", "");
        dict.Add("InvoceSend_Date", null);
        dict.Add("Act_Id", "");
        dict.Add("Comment", "");
        dict.Add("Invoice_PrintComment", "");
        dict.Add("Excel_Type", "import");
        dict.Add("Export", "N");
        dict.Add("Create_Date", DateTime.Now.ToString("yyyy-MM-dd"));
        dict.Add("Create_DateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
        dict.Add("Create_User", SessionInfo.UserName);
        dict.Add("Create_IP", Request.ServerVariables["REMOTE_ADDR"].ToString());
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    protected void Donor_Edit(string Donate_Amt)
    {
        //先抓出第一次/最近一次的捐款日期、捐款次數和捐款總額
        string Begin_DonateDate, Last_DonateDate, Donate_Total_S;
        int Donate_No;
        //****變數宣告****//
        string strSql, uid;
        System.Data.DataTable dt;
        //****變數設定****//
        uid = HFD_Donor_Id.Value;
        //****設定查詢****//
        strSql = " select *  from Donor  where Donor_Id='" + uid + "'";
        //****執行語法****//
        dt = NpoDB.QueryGetTable(strSql);
        DataRow dr = dt.Rows[0];

        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        string strSql2 = "";
        if (dr["Begin_DonateDate"].ToString() == "" || DateTime.Parse(dr["Begin_DonateDate"].ToString()).ToString("yyyy/MM/dd") == "1900/01/01")
        {
            //初次捐款
            strSql2 = " update Donor set ";
            strSql2 += "  Begin_DonateDate = @Begin_DonateDate";
            strSql2 += ", Last_DonateDate = @Last_DonateDate";
            strSql2 += ", Donate_No = @Donate_No";
            strSql2 += ", Donate_Total = @Donate_Total";
            strSql2 += " where Donor_Id = @Donor_Id";

            dict2.Add("Begin_DonateDate", DateTime.Now.ToString("yyyy-MM-dd"));
            dict2.Add("Last_DonateDate", DateTime.Now.ToString("yyyy-MM-dd"));
            dict2.Add("Donate_No", "1");
            dict2.Add("Donate_Total", Donate_Amt.Trim());
            dict2.Add("Donor_Id", HFD_Donor_Id.Value);
        }
        else
        {
            Begin_DonateDate = dr["Begin_DonateDate"].ToString();
            Last_DonateDate = dr["Last_DonateDate"].ToString();
            Donate_No = Int16.Parse(dr["Donate_No"].ToString());
            Donate_Total_S = (Convert.ToInt64(dr["Donate_Total"])).ToString();
            int Donate_Total = Int32.Parse(Donate_Total_S);
            //更新以上欄位

            strSql2 = " update Donor set ";
            strSql2 += "  Last_DonateDate = @Last_DonateDate";
            strSql2 += ", Donate_No = @Donate_No";
            strSql2 += ", Donate_Total = @Donate_Total";
            strSql2 += " where Donor_Id = @Donor_Id";

            dict2.Add("Last_DonateDate", DateTime.Now.ToString("yyyy-MM-dd"));
            dict2.Add("Donate_No", Donate_No + 1);
            dict2.Add("Donate_Total", Donate_Total + int.Parse(Donate_Amt.Trim()));
            dict2.Add("Donor_Id", HFD_Donor_Id.Value);
        }
        NpoDB.ExecuteSQLS(strSql2, dict2);
    }
}
