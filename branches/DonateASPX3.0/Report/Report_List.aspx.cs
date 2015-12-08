using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Report_Report_List : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["ProgID"] = "Report_List";
        //權控處理
        AuthrityControl();
        if (!IsPostBack)
        {
            LoadDropDownListData();
            LoadCheckBoxListData();
            //week3 預設欄位
            //tbxAccumulateDateS_week3.Text = DateTime.Now.Year + "/1/1";
            //tbxAccumulateDateE_week3.Text = DateTime.Now.Year + "/12/31";
            //week4 預設欄位
            tbxAccumulateDateS_week4.Text = DateTime.Now.Year + "/1/1";
            tbxAccumulateDateE_week4.Text = DateTime.Now.Year + "/12/31";
        }

        //Panel_week1.Style.Value = "display:none";
        Panel_week2.Style.Value = "display:none";
        //Panel_week3.Style.Value = "display:none";
        Panel_week4.Style.Value = "display:none";
        Panel_month1.Style.Value = "display:none";
        Panel_month2.Style.Value = "display:none";
        Panel_month3.Style.Value = "display:none";
        Panel_month4.Style.Value = "display:none";
        Panel_month5.Style.Value = "display:none";
        //Panel_season1.Style.Value = "display:none";
        Panel_season2.Style.Value = "display:none";
        Panel_season3.Style.Value = "display:none";
        Panel_season4.Style.Value = "display:none";
        Panel_year1.Style.Value = "display:none";
        //Panel_year2.Style.Value = "display:none";
        Panel_year3.Style.Value = "display:none";
        Panel_year4.Style.Value = "display:none";
        Panel_year5.Style.Value = "display:none";
        //Panel_year6.Style.Value = "display:none";
        Panel_year7.Style.Value = "display:none";
        //Panel_year8.Style.Value = "display:none";
        Panel_year9.Style.Value = "display:none";
        Panel_year10.Style.Value = "display:none";
        Panel_other1.Style.Value = "display:none";
        //Panel_other2.Style.Value = "display:none";
        Panel_other3.Style.Value = "display:none";
        Panel_other4.Style.Value = "display:none";
        Panel_other5.Style.Value = "display:none";
        Panel_other6.Style.Value = "display:none";
        Panel_other7.Style.Value = "display:none";
        Panel_other8.Style.Value = "display:none";
    }
    //-----------------------------------------------------------------------
    public void AuthrityControl()
    {
        if (Authrity.PageRight("_Focus") == false)
        {
            return;
        }
        //輸出網頁
        //Authrity.CheckButtonRight("_Print", btnPrint_week1);
        Authrity.CheckButtonRight("_Print", btnPrint_week2);
        //Authrity.CheckButtonRight("_Print", btnPrint_week3);
        Authrity.CheckButtonRight("_Print", btnPrint_week4);
        Authrity.CheckButtonRight("_Print", btnPrint_month1);
        Authrity.CheckButtonRight("_Print", btnPrint_month1);
        Authrity.CheckButtonRight("_Print", btnPrint_month2);
        Authrity.CheckButtonRight("_Print", btnPrint_month3);
        Authrity.CheckButtonRight("_Print", btnPrint_month4);
        Authrity.CheckButtonRight("_Print", btnPrint_month5);
        //Authrity.CheckButtonRight("_Print", btnPrint_season1);
        Authrity.CheckButtonRight("_Print", btnPrint_season2);
        Authrity.CheckButtonRight("_Print", btnPrint_season3);
        Authrity.CheckButtonRight("_Print", btnPrint_season4);
        Authrity.CheckButtonRight("_Print", btnPrint_year1);
        //Authrity.CheckButtonRight("_Print", btnPrint_year2);
        Authrity.CheckButtonRight("_Print", btnPrint_year3);
        Authrity.CheckButtonRight("_Print", btnPrint_year4);
        Authrity.CheckButtonRight("_Print", btnPrint_year5);
        //Authrity.CheckButtonRight("_Print", btnPrint_year6);
        Authrity.CheckButtonRight("_Print", btnPrint_year7);
        //Authrity.CheckButtonRight("_Print", btnPrint_year8);
        Authrity.CheckButtonRight("_Print", btnPrint_year9);
        Authrity.CheckButtonRight("_Print", btnPrint_year10);
        Authrity.CheckButtonRight("_Print", btnPrint_other1);
        //Authrity.CheckButtonRight("_Print", btnPrint_other2);
        Authrity.CheckButtonRight("_Print", btnPrint_other3);
        Authrity.CheckButtonRight("_Print", btnPrint_other4);
        Authrity.CheckButtonRight("_Print", btnPrint_other5);
        Authrity.CheckButtonRight("_Print", btnPrint_other6);
        Authrity.CheckButtonRight("_Print", btnPrint_other7);
        Authrity.CheckButtonRight("_Print", btnPrint_other8);
        //匯出xls
        //Authrity.CheckButtonRight("_Print", btnToxls_week1);
        Authrity.CheckButtonRight("_Print", btnToxls_week2);
        //Authrity.CheckButtonRight("_Print", btnToxls_week3);
        Authrity.CheckButtonRight("_Print", btnToxls_week4);
        Authrity.CheckButtonRight("_Print", btnToxls_month1);
        Authrity.CheckButtonRight("_Print", btnToxls_month1);
        Authrity.CheckButtonRight("_Print", btnToxls_month2);
        Authrity.CheckButtonRight("_Print", btnToxls_month3);
        Authrity.CheckButtonRight("_Print", btnToxls_month4);
        Authrity.CheckButtonRight("_Print", btnToxls_month5);
        //Authrity.CheckButtonRight("_Print", btnToxls_season1);
        Authrity.CheckButtonRight("_Print", btnToxls_season2);
        Authrity.CheckButtonRight("_Print", btnToxls_season3);
        Authrity.CheckButtonRight("_Print", btnToxls_season4);
        Authrity.CheckButtonRight("_Print", btnToxls_year1);
        //Authrity.CheckButtonRight("_Print", btnToxls_year2);
        Authrity.CheckButtonRight("_Print", btnToxls_year3);
        Authrity.CheckButtonRight("_Print", btnToxls_year4);
        Authrity.CheckButtonRight("_Print", btnToxls_year5);
        //Authrity.CheckButtonRight("_Print", btnToxls_year6);
        Authrity.CheckButtonRight("_Print", btnToxls_year7);
        //Authrity.CheckButtonRight("_Print", btnToxls_year8);
        Authrity.CheckButtonRight("_Print", btnToxls_year9);
        Authrity.CheckButtonRight("_Print", btnToxls_year10);
        Authrity.CheckButtonRight("_Print", btnToxls_other1);
        //Authrity.CheckButtonRight("_Print", btnToxls_other2);
        Authrity.CheckButtonRight("_Print", btnToxls_other3);
        Authrity.CheckButtonRight("_Print", btnToxls_other4);
        Authrity.CheckButtonRight("_Print", btnToxls_other5);
        Authrity.CheckButtonRight("_Print", btnToxls_other6);
        Authrity.CheckButtonRight("_Print", btnToxls_other7);
        Authrity.CheckButtonRight("_Print", btnToxls_other8);
    }
    //----------------------------------------------------------------------
    public void LoadDropDownListData()
    {
        //機構_week1
        //Util.FillDropDownList(ddlDept_week1, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        //ddlDept_week1.SelectedIndex = 0;

        //機構_week2
        Util.FillDropDownList(ddlDept_week2, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_week2.SelectedIndex = 0;

        //機構_week3
        //Util.FillDropDownList(ddlDept_week3, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        //ddlDept_week3.SelectedIndex = 0;

        //機構_week4
        Util.FillDropDownList(ddlDept_week4, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_week4.SelectedIndex = 0;

        //機構_month1
        Util.FillDropDownList(ddlDept_month1, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_month1.SelectedIndex = 0;

        //生日月份_month1
        Util.FillDropDownList(ddlBirthday_Month_month1, Util.GetDataTable("CaseCode", "GroupName", "生日月份", "", ""), "CaseName", "CaseID", true);
        ddlBirthday_Month_month1.SelectedIndex = 0;

        //機構_month2
        Util.FillDropDownList(ddlDept_month2, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_month2.SelectedIndex = 0;

        //授權終止年月_month2
        Util.FillDropDownList(ddlYear_month2, DateTime.Now.Year - 5, DateTime.Now.Year + 5, false, -1);
        ddlYear_month2.Text = DateTime.Now.Year.ToString();
        Util.FillDropDownList(ddlMonth_month2, 1, 12, false, -1);
        ddlMonth_month2.Text = DateTime.Now.Month.ToString();

        //機構_month3
        Util.FillDropDownList(ddlDept_month3, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_month3.SelectedIndex = 0;

        //機構_month4
        Util.FillDropDownList(ddlDept_month4, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_month4.SelectedIndex = 0;

        //機構_month5
        Util.FillDropDownList(ddlDept_month5, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_month5.SelectedIndex = 0;

        //統計年月_month5
        //Util.FillDropDownList(ddlYear_month5, DateTime.Now.Year - 10, DateTime.Now.Year, false, -1);
        //ddlYear_month5.Text = DateTime.Now.Year.ToString();
        //Util.FillDropDownList(ddlMonth_month5, 1, 12, false, -1);
        //ddlMonth_month5.Text = DateTime.Now.Month.ToString();

        //機構_season1
        //Util.FillDropDownList(ddlDept_season1, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        //ddlDept_season1.SelectedIndex = 0;

        //機構_season2
        Util.FillDropDownList(ddlDept_season2, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_season2.SelectedIndex = 0;

        //機構_season3
        Util.FillDropDownList(ddlDept_season3, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_season3.SelectedIndex = 0;

        //機構_season4
        Util.FillDropDownList(ddlDept_season4, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_season4.SelectedIndex = 0;

        //機構_year1
        Util.FillDropDownList(ddlDept_year1, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_year1.SelectedIndex = 0;

        //機構_year2
        //Util.FillDropDownList(ddlDept_year2, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        //ddlDept_year2.SelectedIndex = 0;

        //捐款年_year2
        //Util.FillDropDownList(ddlYear_year2, DateTime.Now.Year - 10, DateTime.Now.Year, false, -1);
        //ddlYear_year2.Text = DateTime.Now.Year.ToString();

        //機構_year3
        Util.FillDropDownList(ddlDept_year3, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_year3.SelectedIndex = 0;

        //機構_year4
        Util.FillDropDownList(ddlDept_year4, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_year4.SelectedIndex = 0;

        //機構_year5
        Util.FillDropDownList(ddlDept_year5, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_year5.SelectedIndex = 0;

        //捐款年_year5
        /*Util.FillDropDownList(ddlYearBegin_year5, DateTime.Now.Year - 10, DateTime.Now.Year, false, -1);
        ddlYearBegin_year5.Text = DateTime.Now.Year.ToString();
        Util.FillDropDownList(ddlYearEnd_year5, DateTime.Now.Year - 10, DateTime.Now.Year, false, -1);
        ddlYearEnd_year5.Text = DateTime.Now.Year.ToString();*/

        //機構_year6
        //Util.FillDropDownList(ddlDept_year6, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        //ddlDept_year6.SelectedIndex = 0;

        //捐款年_year6
        /*Util.FillDropDownList(ddlYearBegin_year6, DateTime.Now.Year - 10, DateTime.Now.Year, false, -1);
        ddlYearBegin_year6.Text = DateTime.Now.Year.ToString();
        Util.FillDropDownList(ddlYearEnd_year6, DateTime.Now.Year - 10, DateTime.Now.Year, false, -1);
        ddlYearEnd_year6.Text = DateTime.Now.Year.ToString();*/

        //機構_year7
        Util.FillDropDownList(ddlDept_year7, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_year7.SelectedIndex = 0;

        //捐款年
        Util.FillDropDownList(ddlYear_year7, DateTime.Now.Year - 10, DateTime.Now.Year, false, -1);
        ddlYear_year7.Text = DateTime.Now.Year.ToString();

        //首捐年_year7
        Util.FillDropDownList(ddlFirst_DonateYear_year7, DateTime.Now.Year - 10, DateTime.Now.Year, false, -1);
        ddlFirst_DonateYear_year7.Text = (DateTime.Now.Year - 10).ToString();

        //機構_year8
        //Util.FillDropDownList(ddlDept_year8, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        //ddlDept_year8.SelectedIndex = 0;

        //機構_year9
        Util.FillDropDownList(ddlDept_year9, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_year9.SelectedIndex = 0;

        //機構_year10
        Util.FillDropDownList(ddlDept_year10, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_year10.SelectedIndex = 0;

        //捐款年
        Util.FillDropDownList(ddlYear_year10, DateTime.Now.Year - 10, DateTime.Now.Year, false, -1);
        ddlYear_year10.Text = (DateTime.Now.Year -1).ToString();

        //機構_other1
        Util.FillDropDownList(ddlDept_other1, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_other1.SelectedIndex = 0;

        //機構_other2
        //Util.FillDropDownList(ddlDept_other2, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        //ddlDept_other2.SelectedIndex = 0;

        //物品名稱_other2
        //Util.FillDropDownList(ddlLinked2Name_other2, Util.GetDataTable2("Linked2", "Linked2_Name", "L2 INNER JOIN Linked L ON L.Linked_Id=L2.Linked_Id", "Linked_Name", "歷年天使DVD贈品", "Linked2_Seq", ""), "Linked2_Name", "Linked2_Name", false);
        //ddlLinked2Name_other2.SelectedIndex = 0;

        //機構_other3
        Util.FillDropDownList(ddlDept_other3, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_other3.SelectedIndex = 0;

        //機構_other4
        Util.FillDropDownList(ddlDept_other4, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_other4.SelectedIndex = 0;
        
        //機構_other5
        Util.FillDropDownList(ddlDept_other5, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_other5.SelectedIndex = 0;

        //機構_other6
        Util.FillDropDownList(ddlDept_other6, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_other6.SelectedIndex = 0;

        //機構_other7
        Util.FillDropDownList(ddlDept_other7, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_other7.SelectedIndex = 0;

        //機構_other8
        Util.FillDropDownList(ddlDept_other8, Util.GetDataTable("Dept", "1", "1", "", ""), "DeptShortName", "DeptId", false);
        ddlDept_other8.SelectedIndex = 0;
    }
    public void LoadCheckBoxListData()
    {
        //捐款用途_week2
        Util.FillCheckBoxList(cblDonate_Purpose_week2, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseName", false);
        cblDonate_Purpose_week2.Items.Insert(0, "會費");
        cblDonate_Purpose_week2.Items[1].Selected = false;

        //身分別_month1
        Util.FillCheckBoxList(cblDonor_Type_month1, Util.GetDataTable("CaseCode", "GroupName", "身份別", "CaseID", ""), "CaseName", "CaseName", false);
        cblDonor_Type_month1.Items[0].Selected = false;

        //授權方式_month2
        Util.FillCheckBoxList(cblDonate_Payment_month2, Util.GetDataTable("Pledge", "distinct Donate_Payment", "1", "1","",""), "Donate_Payment", "Donate_Payment", false);
        cblDonate_Payment_month2.Items[0].Selected = false;

        //捐款用途_month4
        Util.FillCheckBoxList(cblDonate_Purpose_month4, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseName", false);
        cblDonate_Purpose_month4.Items.Insert(0, "會費");
        cblDonate_Purpose_month4.Items[1].Selected = false;

        //捐款用途_season3
        Util.FillCheckBoxList(cblDonate_Purpose_season3, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseName", false);
        cblDonate_Purpose_season3.Items.Insert(0, "會費");
        cblDonate_Purpose_season3.Items[1].Selected = false;

        //捐款用途_year1
        Util.FillCheckBoxList(cblDonate_Purpose_year1, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseName", false);
        cblDonate_Purpose_year1.Items.Insert(0, "會費");
        cblDonate_Purpose_year1.Items[1].Selected = false;

        //縣市_year4
        Util.FillCheckBoxList(cblCity_year4, Util.GetDataTable("CodeCity", "ParentCityID", "0", "Sort", ""), "Name", "Name", false);
        cblCity_year4.Items[0].Selected = false;

        //身分別_year5
        Util.FillCheckBoxList(cblDonor_Type_year5, Util.GetDataTable("CaseCode", "GroupName", "身份別", "CaseID", ""), "CaseName", "CaseName", false);
        cblDonor_Type_year5.Items[0].Selected = false;

        //捐款用途_year7
        Util.FillCheckBoxList(cblDonate_Purpose_year7, Util.GetDataTable("CaseCode", "GroupName", "款項用途", "", ""), "CaseName", "CaseName", false);
        cblDonate_Purpose_year7.Items.Insert(0, "會費");
        cblDonate_Purpose_year7.Items[1].Selected = false;
    }
    //----------------------------------------------------------------------
    protected void btnPrint_Click(object sender, EventArgs e)
    {
        Button btnTmp = (Button)sender;
        switch (btnTmp.ID)
        {
            //Week
            //case "btnPrint_week1":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Week_Report1_Print.aspx?Donate_Purpose=" + ddlDonate_Purpose_week1.SelectedItem.Text + "&DonateDateS=" + tbxDonateDateS_week1.Text + "&DonateDateE=" + tbxDonateDateE_week1.Text +"&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_week1.Style.Value = "display:block";
            //    break;
            case "btnPrint_week2":
                string Donate_Purpose_Items_week2 = "";
                for (int i = 0; i < cblDonate_Purpose_week2.Items.Count; i++)
                {
                    if (cblDonate_Purpose_week2.Items[i].Selected)
                    {
                        Donate_Purpose_Items_week2 += "," + cblDonate_Purpose_week2.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Week_Report2_Print.aspx?DonateDateS=" + tbxDonateDateS_week2.Text + "&DonateDateE=" + tbxDonateDateE_week2.Text + "&Donate_Total_Begin=" + tbxDonate_Total_Begin_week2.Text + "&Donate_Total_End=" + tbxDonate_Total_End_week2.Text + "&Donate_Purpose=" + Donate_Purpose_Items_week2 + "&Is_Abroad=" + cbxIs_Abroad_week2.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_week2.Checked + "&Sex=" + cbxSex_week2.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_week2.Style.Value = "display:block";
                break;
            //case "btnPrint_week3":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Week_Report3_Print.aspx?Donate_Amt=" + tbxDonate_Amt_week3.Text + "&DonateDateS=" + tbxDonateDateS_week3.Text + "&DonateDateE=" + tbxDonateDateE_week3.Text + "&AccumulateDateS=" + tbxAccumulateDateS_week3.Text + "&AccumulateDateE=" + tbxAccumulateDateE_week3.Text + "&Is_Abroad=" + cbxIs_Abroad_week3.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_week3.Checked + "&Sex=" + cbxSex_week3.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_week3.Style.Value = "display:block";
            //    break;
            case "btnPrint_week4":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Week_Report4_Print.aspx?Donate_Amt=" + tbxDonate_Amt_week4.Text + "&DonateDateS=" + tbxDonateDateS_week4.Text + "&DonateDateE=" + tbxDonateDateE_week4.Text + "&AccumulateDateS=" + tbxAccumulateDateS_week4.Text + "&AccumulateDateE=" + tbxAccumulateDateE_week4.Text + "&Is_Abroad=" + cbxIs_Abroad_week4.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_week4.Checked + "&Sex=" + cbxSex_week4.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_week4.Style.Value = "display:block";
                break;
            //Month
            case "btnPrint_month1":
                string Donor_Type_Items_month1 = "";
                for (int i = 0; i < cblDonor_Type_month1.Items.Count; i++)
                {
                    if (cblDonor_Type_month1.Items[i].Selected)
                    {
                        Donor_Type_Items_month1 += "," + cblDonor_Type_month1.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Month_Report1_Print.aspx?Donate_Amt=" + tbxDonate_Amt_month1.Text + "&Donate_Total_Begin=" + tbxDonate_Total_Begin_month1.Text + "&Donate_Total_End=" + tbxDonate_Total_End_month1.Text + "&Donor_Type=" + Donor_Type_Items_month1 + "&Birthday_Month=" + ddlBirthday_Month_month1.SelectedValue + "&Is_Abroad=" + cbxIs_Abroad_month1.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_month1.Checked + "&Sex=" + cbxSex_month1.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_month1.Style.Value = "display:block";
                break;
            case "btnPrint_month2":
                string Donate_Payment_Items_month2 = "";
                for (int i = 0; i < cblDonate_Payment_month2.Items.Count; i++)
                {
                    if (cblDonate_Payment_month2.Items[i].Selected)
                    {
                        Donate_Payment_Items_month2 += "," + cblDonate_Payment_month2.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Month_Report2_Print.aspx?Year=" + ddlYear_month2.SelectedItem.Text + "&Month=" + ddlMonth_month2.SelectedItem.Text + "&Donate_Payment=" + Donate_Payment_Items_month2 + "&Is_Abroad=" + cbxIs_Abroad_month2.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_month2.Checked + "&Sex=" + cbxSex_month2.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_month2.Style.Value = "display:block";
                break;
            case "btnPrint_month3":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Month_Report3_Print.aspx?IsMember=" + rblDonor_Type_month3.SelectedValue + "&Is_Abroad=" + cbxIs_Abroad_month3.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_month3.Checked + "&Sex=" + cbxSex_month3.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_month3.Style.Value = "display:block";
                break;
            case "btnPrint_month4":
                string Donate_Purpose_Items_month4 = "";
                for (int i = 0; i < cblDonate_Purpose_month4.Items.Count; i++)
                {
                    if (cblDonate_Purpose_month4.Items[i].Selected)
                    {
                        Donate_Purpose_Items_month4 += "," + cblDonate_Purpose_month4.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Month_Report4_Print.aspx?Donate_Amt=" + tbxDonate_Amt_month4.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_month4.Text + "&Donate_Purpose=" + Donate_Purpose_Items_month4 + "&Is_Abroad=" + cbxIs_Abroad_month4.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_month4.Checked + "&Sex=" + cbxSex_month4.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_month4.Style.Value = "display:block";
                break;
            case "btnPrint_month5":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Month_Report5_Print.aspx?DonateDateS=" + tbxDonateDateS_month5.Text + "&DonateDateE=" + tbxDonateDateE_month5.Text + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_month5.Style.Value = "display:block";
                break;
            //Season
            //case "btnPrint_season1":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Season_Report1_Print.aspx?DonateDateS=" + tbxDonateDateS_season1.Text + "&DonateDateE=" + tbxDonateDateE_season1.Text + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_season1.Style.Value = "display:block";
            //    break;
            case "btnPrint_season2":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Season_Report2_Print.aspx?Donate_Amt=" + tbxDonate_Amt_season2.Text + "&DonateDateS=" + tbxDonateDateS_season2.Text + "&DonateDateE=" + tbxDonateDateE_season2.Text + "&Is_Abroad=" + cbxIs_Abroad_season2.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_season2.Checked + "&Sex=" + cbxSex_season2.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_season2.Style.Value = "display:block";
                break;
            case "btnPrint_season3":
                string Donate_Purpose_Items_season3 = "";
                for (int i = 0; i < cblDonate_Purpose_season3.Items.Count; i++)
                {
                    if (cblDonate_Purpose_season3.Items[i].Selected)
                    {
                        Donate_Purpose_Items_season3 += "," + cblDonate_Purpose_season3.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Season_Report3_Print.aspx?Donate_Amt=" + tbxDonate_Amt_season3.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_season3.Text + "&Donate_Purpose=" + Donate_Purpose_Items_season3 + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_season3.Style.Value = "display:block";
                break;
            case "btnPrint_season4":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Season_Report4_Print.aspx?Donate_Amt=" + tbxDonate_Amt_season4.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_season4.Text + "&DonateDateS=" + tbxDonateDateS_season4.Text + "&DonateDateE=" + tbxDonateDateE_season4.Text + "&Is_Abroad=" + cbxIs_Abroad_season4.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_season4.Checked + "&Sex=" + cbxSex_season4.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_season4.Style.Value = "display:block";
                break;
            //Year
            case "btnPrint_year1":
                string Donate_Purpose_Items_year1 = "";
                for (int i = 0; i < cblDonate_Purpose_year1.Items.Count; i++)
                {
                    if (cblDonate_Purpose_year1.Items[i].Selected)
                    {
                        Donate_Purpose_Items_year1 += "," + cblDonate_Purpose_year1.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report1_Print.aspx?Donate_Amt=" + tbxDonate_Amt_year1.Text + "&Donate_Total_Begin=" + tbxDonate_Total_Begin_year1.Text + "&Donate_Total_End=" + tbxDonate_Total_End_year1.Text + "&DonateDateS=" + tbxDonateDateS_year1.Text + "&DonateDateE=" + tbxDonateDateE_year1.Text + "&Donate_Purpose=" + Donate_Purpose_Items_year1 + "&Is_Abroad=" + cbxIs_Abroad_year1.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year1.Checked + "&Sex=" + cbxSex_year1.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year1.Style.Value = "display:block";
                break;
            //case "btnPrint_year2":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report2_Print.aspx?Year=" + ddlYear_year2.SelectedItem.Text + "&Is_Abroad=" + cbxIs_Abroad_year2.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year2.Checked + "&Sex=" + cbxSex_year2.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_year2.Style.Value = "display:block";
            //    break;
            case "btnPrint_year3":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report3_Print.aspx?Donate_Amt=" + tbxDonate_Amt_year3.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_year3.Text + "&DonateDateS=" + tbxDonateDateS_year3.Text + "&DonateDateE=" + tbxDonateDateE_year3.Text + "&Is_ErrAddress=" + cbxIs_ErrAddress_year3.Checked + "&Sex=" + cbxSex_year3.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year3.Style.Value = "display:block";
                break;
            case "btnPrint_year4":
                string City_Items_year4 = "";
                for (int i = 0; i < cblCity_year4.Items.Count; i++)
                {
                    if (cblCity_year4.Items[i].Selected)
                    {
                        City_Items_year4 += "," + cblCity_year4.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report4_Print.aspx?Donate_Amt=" + tbxDonate_Amt_year4.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_year4.Text + "&DonateDateS=" + tbxDonateDateS_year4.Text + "&DonateDateE=" + tbxDonateDateE_year4.Text + "&City=" + City_Items_year4 + "&Is_ErrAddress=" + cbxIs_ErrAddress_year4.Checked + "&Sex=" + cbxSex_year4.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year4.Style.Value = "display:block";
                break;
            case "btnPrint_year5":
                string Donor_Type_Items_year5 = "";
                for (int i = 0; i < cblDonor_Type_year5.Items.Count; i++)
                {
                    if (cblDonor_Type_year5.Items[i].Selected)
                    {
                        Donor_Type_Items_year5 += "," + cblDonor_Type_year5.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report5_Print.aspx?Donate_Amt=" + tbxDonate_Amt_year5.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_year5.Text + "&Donor_Type=" + Donor_Type_Items_year5 + "&DonateDateS=" + tbxDonateDateS_year5.Text + "&DonateDateE=" + tbxDonateDateE_year5.Text + "&Is_Abroad=" + cbxIs_Abroad_year5.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year5.Checked + "&Sex=" + cbxSex_year5.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year5.Style.Value = "display:block";
                break;
            //case "btnPrint_year6":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report6_Print.aspx?Donate_Amt=" + tbxDonate_Amt_year6.Text + "&DonateDateS=" + tbxDonateDateS_year6.Text + "&DonateDateE=" + tbxDonateDateE_year6.Text + "&Is_Abroad=" + cbxIs_Abroad_year6.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year6.Checked + "&Sex=" + cbxSex_year6.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_year6.Style.Value = "display:block";
            //    break;
            case "btnPrint_year7":
                string Donate_Purpose_Items_year7 = "";
                for (int i = 0; i < cblDonate_Purpose_year7.Items.Count; i++)
                {
                    if (cblDonate_Purpose_year7.Items[i].Selected)
                    {
                        Donate_Purpose_Items_year7 += "," + cblDonate_Purpose_year7.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report7_Print.aspx?Donate_Amt=" + tbxDonate_Amt_year7.Text + "&Donate_Purpose=" + Donate_Purpose_Items_year7 + "&Year=" + ddlYear_year7.SelectedItem.Text + "&First_DonateYear=" + ddlFirst_DonateYear_year7.SelectedItem.Text + "&Last_Donate_Date=" + tbxLast_Donate_Date_year7.Text + "&Is_Abroad=" + cbxIs_Abroad_year7.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year7.Checked + "&Sex=" + cbxSex_year7.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year7.Style.Value = "display:block";
                break;
            //case "btnPrint_year8":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report8_Print.aspx?Donate_Amt=" + tbxDonate_Amt_year8.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_year8.Text + "&DonateDateS=" + tbxDonateDateS_year8.Text + "&DonateDateE=" + tbxDonateDateE_year8.Text + "&Is_Abroad=" + cbxIs_Abroad_year8.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year8.Checked + "&Sex=" + cbxSex_year8.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_year8.Style.Value = "display:block";
            //    break;
            case "btnPrint_year9":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report9_Print.aspx?Donate_Total_Amt=" + tbxDonate_Total_Amt_year9.Text + "&DonateDateS=" + tbxDonateDateS_year9.Text + "&DonateDateE=" + tbxDonateDateE_year9.Text + "&Is_ErrAddress=" + cbxIs_ErrAddress_year9.Checked + "&Sex=" + cbxSex_year9.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year9.Style.Value = "display:block";
                break;
            case "btnPrint_year10":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report10_Print.aspx?Year=" + ddlYear_year10.SelectedItem.Text + "&DonorId=" + tbxDonorId_year10.Text + "&Is_Abroad=" + cbxIs_Abroad_year10.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year10.Checked + "&Sex=" + cbxSex_year10.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year10.Style.Value = "display:block";
                break;
            //Other
            case "btnPrint_other1":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report1_Print.aspx?DonateDateS=" + tbxDonateDateS_other1.Text + "&DonateDateE=" + tbxDonateDateE_other1.Text + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other1.Style.Value = "display:block";
                break;
            //case "btnPrint_other2":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report2_Print.aspx?Linked2_Name=" + ddlLinked2Name_other2.SelectedItem.Text + "&ContributeDateS=" + tbxContributeDateS_other2.Text + "&ContributeDateE=" + tbxContributeE_other2.Text + "&Is_Abroad=" + cbxIs_Abroad_other2.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_other2.Checked + "&Sex=" + cbxSex_other2.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_other2.Style.Value = "display:block";
            //    break;
            case "btnPrint_other3":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report3_Print.aspx?Is_SameName=" + cbxIs_SameName_other3.Checked + "&Is_SameAddress=" + cbxIs_SameAddress_other3.Checked + "&Address=" + rblAddress_other3.SelectedValue + "&Is_SameTel=" + cbxIs_SameTel_other3.Checked + "&Tel=" + rblTel_other3.SelectedValue + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other3.Style.Value = "display:block";
                break;
            case "btnPrint_other4":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report4_Print.aspx?IsMember=" + rblDonor_Type_other4.SelectedValue + "&DonateDateS=" + tbxDonateDateS_other4.Text + "&DonateDateE=" + tbxDonateDateE_other4.Text + "&Is_Abroad=" + cbxIs_Abroad_other4.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_other4.Checked + "&Sex=" + cbxSex_other4.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other4.Style.Value = "display:block";
                break;
            case "btnPrint_other5":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report5_Print.aspx?IsMember=" + rblDonor_Type_other5.SelectedValue + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other5.Style.Value = "display:block";
                break;
            case "btnPrint_other6":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report6_Print.aspx?Type=" + rbl_Type_other6.SelectedValue + "&DonateDateS=" + tbxDonateDateS_other6.Text + "&DonateDateE=" + tbxDonateDateE_other6.Text + "&Is_Abroad=" + cbxIs_Abroad_other6.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_other6.Checked + "&Sex=" + cbxSex_other6.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other6.Style.Value = "display:block";
                break;
            case "btnPrint_other7":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report7_Print.aspx?DonateDateS=" + tbxDonateDateS_other7.Text + "&DonateDateE=" + tbxDonateDateE_other7.Text + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other7.Style.Value = "display:block";
                break;
            case "btnPrint_other8":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report8_Print.aspx?DonorName=" + tbxDonorName_other8.Text + "&DonateMotive1=" + DonateMotive1.Checked + "&DonateMotive2=" + DonateMotive2.Checked + "&DonateMotive3=" + DonateMotive3.Checked + "&DonateMotive4=" + DonateMotive4.Checked + "&DonateMotive5=" + DonateMotive5.Checked + "&DonateMotive6=" + DonateMotive6.Checked + "&WatchMode1=" + WatchMode1.Checked + "&WatchMode2=" + WatchMode2.Checked + "&WatchMode3=" + WatchMode3.Checked + "&WatchMode4=" + WatchMode4.Checked + "&WatchMode5=" + WatchMode5.Checked + "&WatchMode6=" + WatchMode6.Checked + "&WatchMode7=" + WatchMode7.Checked + "&WatchMode8=" + WatchMode8.Checked + "&WatchMode9=" + WatchMode9.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other8.Style.Value = "display:block";
                break;
        }
    }
    protected void btnToxls_Click(object sender, EventArgs e)
    {
        Button btnTmp = (Button)sender;
        switch (btnTmp.ID)
        {
            //Week
            //case "btnToxls_week1":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Week_Report1_Print_Excel.aspx?Donate_Purpose=" + ddlDonate_Purpose_week1.SelectedItem.Text + "&DonateDateS=" + tbxDonateDateS_week1.Text + "&DonateDateE=" + tbxDonateDateE_week1.Text + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_week1.Style.Value = "display:block";
            //    break;
            case "btnToxls_week2":
                string Donate_Purpose_Items_week2 = "";
                for (int i = 0; i < cblDonate_Purpose_week2.Items.Count; i++)
                {
                    if (cblDonate_Purpose_week2.Items[i].Selected)
                    {
                        Donate_Purpose_Items_week2 += "," + cblDonate_Purpose_week2.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Week_Report2_Print_Excel.aspx?DonateDateS=" + tbxDonateDateS_week2.Text + "&DonateDateE=" + tbxDonateDateE_week2.Text + "&Donate_Total_Begin=" + tbxDonate_Total_Begin_week2.Text + "&Donate_Total_End=" + tbxDonate_Total_End_week2.Text + "&Donate_Purpose=" + Donate_Purpose_Items_week2 + "&Is_Abroad=" + cbxIs_Abroad_week2.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_week2.Checked + "&Sex=" + cbxSex_week2.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_week2.Style.Value = "display:block";
                break;
            //case "btnToxls_week3":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Week_Report3_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_week3.Text + "&DonateDateS=" + tbxDonateDateS_week3.Text + "&DonateDateE=" + tbxDonateDateE_week3.Text + "&AccumulateDateS=" + tbxAccumulateDateS_week3.Text + "&AccumulateDateE=" + tbxAccumulateDateE_week3.Text + "&Is_Abroad=" + cbxIs_Abroad_week3.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_week3.Checked + "&Sex=" + cbxSex_week3.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_week3.Style.Value = "display:block";
            //    break;
            case "btnToxls_week4":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Week_Report4_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_week4.Text + "&DonateDateS=" + tbxDonateDateS_week4.Text + "&DonateDateE=" + tbxDonateDateE_week4.Text + "&AccumulateDateS=" + tbxAccumulateDateS_week4.Text + "&AccumulateDateE=" + tbxAccumulateDateE_week4.Text + "&Is_Abroad=" + cbxIs_Abroad_week4.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_week4.Checked + "&Sex=" + cbxSex_week4.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_week4.Style.Value = "display:block";
                break;
            //Month
            case "btnToxls_month1":
                string Donor_Type_Items_month1 = "";
                for (int i = 0; i < cblDonor_Type_month1.Items.Count; i++)
                {
                    if (cblDonor_Type_month1.Items[i].Selected)
                    {
                        Donor_Type_Items_month1 += "," + cblDonor_Type_month1.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Month_Report1_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_month1.Text + "&Donate_Total_Begin=" + tbxDonate_Total_Begin_month1.Text + "&Donate_Total_End=" + tbxDonate_Total_End_month1.Text + "&Donor_Type=" + Donor_Type_Items_month1 + "&Birthday_Month=" + ddlBirthday_Month_month1.SelectedValue + "&Is_Abroad=" + cbxIs_Abroad_month1.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_month1.Checked + "&Sex=" + cbxSex_month1.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_month1.Style.Value = "display:block";
                break;
            case "btnToxls_month2":
                string Donate_Payment_Items_month2 = "";
                for (int i = 0; i < cblDonate_Payment_month2.Items.Count; i++)
                {
                    if (cblDonate_Payment_month2.Items[i].Selected)
                    {
                        Donate_Payment_Items_month2 += "," + cblDonate_Payment_month2.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Month_Report2_Print_Excel.aspx?Year=" + ddlYear_month2.SelectedItem.Text + "&Month=" + ddlMonth_month2.SelectedItem.Text + "&Donate_Payment=" + Donate_Payment_Items_month2 + "&Is_Abroad=" + cbxIs_Abroad_month2.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_month2.Checked + "&Sex=" + cbxSex_month2.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_month2.Style.Value = "display:block";
                break;
            case "btnToxls_month3":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Month_Report3_Print_Excel.aspx?IsMember=" + rblDonor_Type_month3.SelectedValue + "&Is_Abroad=" + cbxIs_Abroad_month3.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_month3.Checked + "&Sex=" + cbxSex_month3.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_month3.Style.Value = "display:block";
                break;
            case "btnToxls_month4":
                string Donate_Purpose_Items_month4 = "";
                for (int i = 0; i < cblDonate_Purpose_month4.Items.Count; i++)
                {
                    if (cblDonate_Purpose_month4.Items[i].Selected)
                    {
                        Donate_Purpose_Items_month4 += "," + cblDonate_Purpose_month4.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Month_Report4_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_month4.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_month4.Text + "&Donate_Purpose=" + Donate_Purpose_Items_month4 + "&Is_Abroad=" + cbxIs_Abroad_month4.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_month4.Checked + "&Sex=" + cbxSex_month4.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_month4.Style.Value = "display:block";
                break;
            case "btnToxls_month5":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Month_Report5_Print_Excel.aspx?DonateDateS=" + tbxDonateDateS_month5.Text + "&DonateDateE=" + tbxDonateDateE_month5.Text + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_month5.Style.Value = "display:block";
                break;
            //Season
            //case "btnToxls_season1":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Season_Report1_Print_Excel.aspx?DonateDateS=" + tbxDonateDateS_season1.Text + "&DonateDateE=" + tbxDonateDateE_season1.Text + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_season1.Style.Value = "display:block";
            //    break;
            case "btnToxls_season2":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Season_Report2_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_season2.Text + "&DonateDateS=" + tbxDonateDateS_season2.Text + "&DonateDateE=" + tbxDonateDateE_season2.Text + "&Is_Abroad=" + cbxIs_Abroad_season2.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_season2.Checked + "&Sex=" + cbxSex_season2.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_season2.Style.Value = "display:block";
                break;
            case "btnToxls_season3":
                string Donate_Purpose_Items_season3 = "";
                for (int i = 0; i < cblDonate_Purpose_season3.Items.Count; i++)
                {
                    if (cblDonate_Purpose_season3.Items[i].Selected)
                    {
                        Donate_Purpose_Items_season3 += "," + cblDonate_Purpose_season3.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Season_Report3_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_season3.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_season3.Text + "&Donate_Purpose=" + Donate_Purpose_Items_season3 + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_season3.Style.Value = "display:block";
                break;
            case "btnToxls_season4":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Season_Report4_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_season4.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_season4.Text + "&DonateDateS=" + tbxDonateDateS_season4.Text + "&DonateDateE=" + tbxDonateDateE_season4.Text + "&Is_Abroad=" + cbxIs_Abroad_season4.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_season4.Checked + "&Sex=" + cbxSex_season4.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_season4.Style.Value = "display:block";
                break;
            //Year
            case "btnToxls_year1":
                string Donate_Purpose_Items_year1 = "";
                for (int i = 0; i < cblDonate_Purpose_year1.Items.Count; i++)
                {
                    if (cblDonate_Purpose_year1.Items[i].Selected)
                    {
                        Donate_Purpose_Items_year1 += "," + cblDonate_Purpose_year1.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report1_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_year1.Text + "&Donate_Total_Begin=" + tbxDonate_Total_Begin_year1.Text + "&Donate_Total_End=" + tbxDonate_Total_End_year1.Text + "&DonateDateS=" + tbxDonateDateS_year1.Text + "&DonateDateE=" + tbxDonateDateE_year1.Text + "&Donate_Purpose=" + Donate_Purpose_Items_year1 + "&Is_Abroad=" + cbxIs_Abroad_year1.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year1.Checked + "&Sex=" + cbxSex_year1.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year1.Style.Value = "display:block";
                break;
            //case "btnToxls_year2":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report2_Print_Excel.aspx?Year=" + ddlYear_year2.SelectedItem.Text + "&Is_Abroad=" + cbxIs_Abroad_year2.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year2.Checked + "&Sex=" + cbxSex_year2.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_year2.Style.Value = "display:block";
            //    break;
            case "btnToxls_year3":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report3_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_year3.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_year3.Text + "&DonateDateS=" + tbxDonateDateS_year3.Text + "&DonateDateE=" + tbxDonateDateE_year3.Text + "&Is_ErrAddress=" + cbxIs_ErrAddress_year3.Checked + "&Sex=" + cbxSex_year3.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year3.Style.Value = "display:block";
                break;
            case "btnToxls_year4":
                string City_Items_year4 = "";
                for (int i = 0; i < cblCity_year4.Items.Count; i++)
                {
                    if (cblCity_year4.Items[i].Selected)
                    {
                        City_Items_year4 += "," + cblCity_year4.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report4_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_year4.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_year4.Text + "&DonateDateS=" + tbxDonateDateS_year4.Text + "&DonateDateE=" + tbxDonateDateE_year4.Text + "&City=" + City_Items_year4 + "&Is_ErrAddress=" + cbxIs_ErrAddress_year4.Checked + "&Sex=" + cbxSex_year4.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year4.Style.Value = "display:block";
                break;
            case "btnToxls_year5":
                string Donor_Type_Items_year5 = "";
                for (int i = 0; i < cblDonor_Type_year5.Items.Count; i++)
                {
                    if (cblDonor_Type_year5.Items[i].Selected)
                    {
                        Donor_Type_Items_year5 += "," + cblDonor_Type_year5.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report5_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_year5.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_year5.Text + "&Donor_Type=" + Donor_Type_Items_year5 + "&DonateDateS=" + tbxDonateDateS_year5.Text + "&DonateDateE=" + tbxDonateDateE_year5.Text + "&Is_Abroad=" + cbxIs_Abroad_year5.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year5.Checked + "&Sex=" + cbxSex_year5.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year5.Style.Value = "display:block";
                break;
            //case "btnToxls_year6":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report6_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_year6.Text + "&DonateDateS=" + tbxDonateDateS_year6.Text + "&DonateDateE=" + tbxDonateDateE_year6.Text + "&Is_Abroad=" + cbxIs_Abroad_year6.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year6.Checked + "&Sex=" + cbxSex_year6.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_year6.Style.Value = "display:block";
            //    break;
            case "btnToxls_year7":
                string Donate_Purpose_Items_year7 = "";
                for (int i = 0; i < cblDonate_Purpose_year7.Items.Count; i++)
                {
                    if (cblDonate_Purpose_year7.Items[i].Selected)
                    {
                        Donate_Purpose_Items_year7 += "," + cblDonate_Purpose_year7.Items[i].ToString();
                    }
                }
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report7_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_year7.Text + "&Donate_Purpose=" + Donate_Purpose_Items_year7 + "&Year=" + ddlYear_year7.SelectedItem.Text + "&First_DonateYear=" + ddlFirst_DonateYear_year7.SelectedItem.Text + "&Last_Donate_Date=" + tbxLast_Donate_Date_year7.Text + "&Is_Abroad=" + cbxIs_Abroad_year7.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year7.Checked + "&Sex=" + cbxSex_year7.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year7.Style.Value = "display:block";
                break;
            //case "btnToxls_year8":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report8_Print_Excel.aspx?Donate_Amt=" + tbxDonate_Amt_year8.Text + "&Donate_Total_Amt=" + tbxDonate_Total_Amt_year8.Text + "&DonateDateS=" + tbxDonateDateS_year8.Text + "&DonateDateE=" + tbxDonateDateE_year8.Text + "&Is_Abroad=" + cbxIs_Abroad_year8.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year8.Checked + "&Sex=" + cbxSex_year8.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_year8.Style.Value = "display:block";
            //    break;
            case "btnToxls_year9":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report9_Print_Excel.aspx?Donate_Total_Amt=" + tbxDonate_Total_Amt_year9.Text + "&DonateDateS=" + tbxDonateDateS_year9.Text + "&DonateDateE=" + tbxDonateDateE_year9.Text + "&Is_ErrAddress=" + cbxIs_ErrAddress_year9.Checked + "&Sex=" + cbxSex_year9.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year9.Style.Value = "display:block";
                break;
            case "btnToxls_year10":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Year_Report10_Print_Excel.aspx?Year=" + ddlYear_year10.SelectedItem.Text + "&DonorId=" + tbxDonorId_year10.Text + "&Is_Abroad=" + cbxIs_Abroad_year10.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_year10.Checked + "&Sex=" + cbxSex_year10.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_year10.Style.Value = "display:block";
                break;
            //Other
            case "btnToxls_other1":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report1_Print_Excel.aspx?DonateDateS=" + tbxDonateDateS_other1.Text + "&DonateDateE=" + tbxDonateDateE_other1.Text + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other1.Style.Value = "display:block";
                break;
            //case "btnToxls_other2":
            //    ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report2_Print_Excel.aspx?Linked2_Name=" + ddlLinked2Name_other2.SelectedItem.Text + "&ContributeDateS=" + tbxContributeDateS_other2.Text + "&ContributeDateE=" + tbxContributeE_other2.Text + "&Is_Abroad=" + cbxIs_Abroad_other2.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_other2.Checked + "&Sex=" + cbxSex_other2.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
            //    Panel_other2.Style.Value = "display:block";
            //    break;
            case "btnToxls_other3":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report3_Print_Excel.aspx?Is_SameName=" + cbxIs_SameName_other3.Checked + "&Is_SameAddress=" + cbxIs_SameAddress_other3.Checked + "&Address=" + rblAddress_other3.SelectedValue + "&Is_SameTel=" + cbxIs_SameTel_other3.Checked + "&Tel=" + rblTel_other3.SelectedValue + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other3.Style.Value = "display:block";
                break;
            case "btnToxls_other4":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report4_Print_Excel.aspx?IsMember=" + rblDonor_Type_other4.SelectedValue + "&DonateDateS=" + tbxDonateDateS_other4.Text + "&DonateDateE=" + tbxDonateDateE_other4.Text + "&Is_Abroad=" + cbxIs_Abroad_other4.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_other4.Checked + "&Sex=" + cbxSex_other4.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other4.Style.Value = "display:block";
                break;
            case "btnToxls_other5":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report5_Print_Excel.aspx?IsMember=" + rblDonor_Type_other5.SelectedValue + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other5.Style.Value = "display:block";
                break;
            case "btnToxls_other6":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report6_Print_Excel.aspx?Type=" + rbl_Type_other6.SelectedValue + "&DonateDateS=" + tbxDonateDateS_other6.Text + "&DonateDateE=" + tbxDonateDateE_other6.Text + "&Is_Abroad=" + cbxIs_Abroad_other6.Checked + "&Is_ErrAddress=" + cbxIs_ErrAddress_other6.Checked + "&Sex=" + cbxSex_other6.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other6.Style.Value = "display:block";
                break;
            case "btnToxls_other7":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report7_Print_Excel.aspx?DonateDateS=" + tbxDonateDateS_other7.Text + "&DonateDateE=" + tbxDonateDateE_other7.Text + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other7.Style.Value = "display:block";
                break;
            case "btnToxls_other8":
                ResponseHelper.Redirect(Util.RedirectByTime("Custom_Report/Donate_Other_Report8_Print_Excel.aspx?DonorName=" + tbxDonorName_other8.Text + "&DonateMotive1=" + DonateMotive1.Checked + "&DonateMotive2=" + DonateMotive2.Checked + "&DonateMotive3=" + DonateMotive3.Checked + "&DonateMotive4=" + DonateMotive4.Checked + "&DonateMotive5=" + DonateMotive5.Checked + "&DonateMotive6=" + DonateMotive6.Checked + "&WatchMode1=" + WatchMode1.Checked + "&WatchMode2=" + WatchMode2.Checked + "&WatchMode3=" + WatchMode3.Checked + "&WatchMode4=" + WatchMode4.Checked + "&WatchMode5=" + WatchMode5.Checked + "&WatchMode6=" + WatchMode6.Checked + "&WatchMode7=" + WatchMode7.Checked + "&WatchMode8=" + WatchMode8.Checked + "&WatchMode9=" + WatchMode9.Checked + "&"), "_blank", "toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=800,height=600");
                Panel_other8.Style.Value = "display:block";
                break;
        }
    }
    public static class ResponseHelper
    {
        public static void Redirect(string url, string target, string windowFeatures)
        {

            Page page = (Page)HttpContext.Current.Handler;

            if (page == null)
            {
                throw new InvalidOperationException("Cannot redirect to new window outside Page context.");
            }
            url = page.ResolveClientUrl(url);

            string script;
            if (!String.IsNullOrEmpty(windowFeatures))
            {
                script = @"window.open(""{0}"", ""{1}"", ""{2}"");";
            }
            else
            {
                script = @"window.open(""{0}"", ""{1}"");";
            }
            script = String.Format(script, url, target, windowFeatures);
            ScriptManager.RegisterStartupScript(page, typeof(Page), "Redirect", script, true);
        }
    }
    //控制其他-雷同資料查詢報表other3 地址查詢顯示與否
    protected void cbxIs_Address_CheckedChanged(object sender, EventArgs e)
    {
        if (cbxIs_SameAddress_other3.Checked)
        {
            
            IsAddress_other3.Visible = true;
            Panel_other3.Style.Value = "display:block";
        }
        else
        {
            IsAddress_other3.Visible = false;
            Panel_other3.Style.Value = "display:block";
        }
    }
    //控制其他-雷同資料查詢報表other3 電話查詢顯示與否
    protected void cbxIs_Tel_CheckedChanged(object sender, EventArgs e)
    {
        if (cbxIs_SameTel_other3.Checked)
        {

            IsTel_other3.Visible = true;
            Panel_other3.Style.Value = "display:block";
        }
        else
        {
            IsTel_other3.Visible = false;
            Panel_other3.Style.Value = "display:block";
        }
    }
}