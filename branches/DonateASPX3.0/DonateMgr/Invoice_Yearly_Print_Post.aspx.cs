using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class DonateMgr_Invoice_Yearly_Print_Post : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            string strSql = Session["strSql_Print"].ToString();
            Print(strSql);
        }
    }
    public void Print(string strSql)
    {
        //抓DEPT欄位
        string strSql_Dept = "select * from Dept where DeptID='" + SessionInfo.DeptID + "'";
        DataTable dt_Dept = NpoDB.GetDataTableS(strSql_Dept, null);
        DataRow dr_Dept = dt_Dept.Rows[0];

        int Row = 1;
        int LastRow = 0;
        DataTable dt_Donor = NpoDB.GetDataTableS(strSql, null);
        DataRow dr_Donor;
        if (dt_Donor.Rows.Count != 0)
        {
            GridList.Text += "<table width='720'  border='0' cellspacing='0'>";
            GridList.Text += "  <tr>";
            GridList.Text += "    <td width='40%' height='120' valign='bottom' colspan='2'>";
            GridList.Text += "      <table width='100%' border='0' cellspacing='0'>";
            GridList.Text += "        <tr>";
            GridList.Text += "          <td height='25'><span class='style3'>中華民國&nbsp;" + (DateTime.Now.Year - 1911).ToString() + "&nbsp;年&nbsp;" + DateTime.Now.Month + "&nbsp;月&nbsp;" + DateTime.Now.Day + "&nbsp;日</span></td>";
            GridList.Text += "        </tr>";
            GridList.Text += "        <tr>";
            GridList.Text += "          <td height='25'><span class='style3'>寄件人：" + SessionInfo.UserName + "</span></td>";
            GridList.Text += "        </tr>";
            GridList.Text += "        <tr>";
            GridList.Text += "          <td height='25'><span class='style3'>名&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;稱：" + SessionInfo.OrgName + "</span></td>";
            GridList.Text += "        </tr>";
            GridList.Text += "      </table>";
            GridList.Text += "    </td>";
            GridList.Text += "    <td width='30%' height='120' align='center'>";
            GridList.Text += "      <p class='style1'><u>中&nbsp; 華&nbsp; 民&nbsp; 國&nbsp; 郵&nbsp; 政</u></p>";
            GridList.Text += "      <p class='style1'><u>限時掛號</u></p>";
            GridList.Text += "      <p class='style1'> 交寄大宗<u>掛&nbsp;&nbsp;&nbsp; 號</u>函件執據</p>";
            GridList.Text += "      <p class='style1'> <u>快捷郵件</u> </p>";
            GridList.Text += "    </td>";
            GridList.Text += "    <td width='30%' height='120' align='center'>";
            GridList.Text += "      <p><img src='../images/post.gif' width='82' height='80'></p>";
            GridList.Text += "      <p class='style3'>郵局郵戳</p>";
            GridList.Text += "    </td>";
            GridList.Text += "  </tr>";
            GridList.Text += "  <tr>";
            GridList.Text += "    <td width='21%' height='25'><span class='style3'>寄件人代表：" + dr_Dept["Contactor"].ToString() + "</span></td>";
            GridList.Text += "    <td width='48%' height='25' colspan='2'><span class='style3'>詳細地址：235" + dr_Dept["Address"].ToString() + "</span></td>";
            GridList.Text += "    <td width='30%' height='25'><span class='style3'>電話號碼：" + dr_Dept["TEL"].ToString() + "</span></td>";
            GridList.Text += "  </tr>";
            GridList.Text += "</table>";
            GridList.Text += "<table width='720'  border='1' cellspacing='0' bordercolor='#000000' style='border-collapse: collapse'>";
            GridList.Text += "  <tr>";
            GridList.Text += "    <td width='5%' height='25' rowspan='2'><div align='center' class='style3'>順序號碼</div></td>";
            GridList.Text += "    <td width='10%' height='25' rowspan='2'><div align='center' class='style3'>掛號號碼</div></td>";
            GridList.Text += "    <td height='25' colspan='2'><div align='center' class='style3'>收件人</div></td>";
            GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否回執(V)</div></td>";
            GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否航空(V)</div></td>";
            GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否印刷物(V)</div></td>";
            GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>重量</div></td>";
            GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>郵資</div></td>";
            GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>備考</div></td>";
            GridList.Text += "  </tr>";
            GridList.Text += "  <tr>";
            GridList.Text += "    <td width='10%' height='25'><div align='center' class='style3'>姓名</div></td>";
            GridList.Text += "    <td width='38%' height='25'><p align='center' class='style3'>寄達第名(或地址)</p>    </td>";
            GridList.Text += "  </tr>";

            bool end = false;
            while (end == false)
            {
                dr_Donor = dt_Donor.Rows[Row - 1];
                GridList.Text += "<tr>";
                GridList.Text += "  <td width='5%' height='30'><div align='center' class='style3'>" + Row + "</div></td>";
                GridList.Text += "  <td width='10%' height='30'><span class='style3'></span></td>";
                GridList.Text += "  <td width='10%' height='30'><span class='style3'>" + dr_Donor["捐款人"].ToString() + "</span></td>";
                GridList.Text += "  <td width='38%' height='30'><span class='style3'>" + dr_Donor["郵遞區號"].ToString() + dr_Donor["地址"].ToString() + "</span></td>";
                GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                GridList.Text += "</tr>";

                if ((Row + 20) % 20 == 0)
                {
                    GridList.Text += "</table>";
                    GridList.Text += "<table width='720'  border='0' cellspacing='10'>";
                    GridList.Text += "  <tr>";
                    GridList.Text += "    <td width='58%' valign='top'>";
                    GridList.Text += "      <span class='style3'>";
                    GridList.Text += "        限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br />";
                    GridList.Text += "	      函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br />";
                    GridList.Text += "	      將本埠與外埠函件分別列單交寄。<br />";
                    GridList.Text += "	      此單由郵局免費供給，應由寄件人清晰填寫一式二份。<br />";
                    GridList.Text += "	      如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局經辦員逐件核對。<br />";
                    GridList.Text += "	      日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br />";
                    GridList.Text += "	      錢鈔或有價證券請利用報值或保價交寄。<br />";
                    GridList.Text += "	    </span>";
                    GridList.Text += "	  </td>";
                    GridList.Text += "    <td width='38%'>";
                    GridList.Text += "      <p class='style3'>上開　限時掛號 </p>";
                    GridList.Text += "      <p class='style3'>掛號函件 / 共　　20　　件 照收無誤　快捷郵件 </p>";
                    GridList.Text += "	    <p class='style3'>　</p>";
                    GridList.Text += "      <p class='style3'><font size='3'>郵資共計　　　　&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 元</font> </p>";
                    GridList.Text += "      <p align='center' class='style3'>經辦員簽署</p>";
                    GridList.Text += "    </td>";
                    GridList.Text += "  </tr>";
                    GridList.Text += "</table>";
                    if (Row < dt_Donor.Rows.Count)
                    {
                        GridList.Text += "<div class='pagebreak'>&nbsp;</div>";
                        GridList.Text += "<table width='720'  border='0' cellspacing='0'>";
                        GridList.Text += "  <tr>";
                        GridList.Text += "    <td width='40%' height='120' valign='bottom' colspan='2'>";
                        GridList.Text += "      <table width='100%' border='0' cellspacing='0'>";
                        GridList.Text += "        <tr>";
                        GridList.Text += "          <td height='25'><span class='style3'>中華民國&nbsp;" + (DateTime.Now.Year - 1911).ToString() + "&nbsp;年&nbsp;" + DateTime.Now.Month + "&nbsp;月&nbsp;" + DateTime.Now.Day + "&nbsp;日</span></td>";
                        GridList.Text += "        </tr>";
                        GridList.Text += "        <tr>";
                        GridList.Text += "          <td height='25'><span class='style3'>寄件人：" + SessionInfo.UserName + "</span></td>";
                        GridList.Text += "        </tr>";
                        GridList.Text += "        <tr>";
                        GridList.Text += "          <td height='25'><span class='style3'>名&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;稱：" + SessionInfo.OrgName + "</span></td>";
                        GridList.Text += "        </tr>";
                        GridList.Text += "      </table>";
                        GridList.Text += "    </td>";
                        GridList.Text += "    <td width='30%' height='120' align='center'>";
                        GridList.Text += "      <p class='style1'><u>中&nbsp; 華&nbsp; 民&nbsp; 國&nbsp; 郵&nbsp; 政</u></p>";
                        GridList.Text += "      <p class='style1'><u>限時掛號</u></p>";
                        GridList.Text += "      <p class='style1'> 交寄大宗<u>掛&nbsp;&nbsp;&nbsp; 號</u>函件執據</p>";
                        GridList.Text += "      <p class='style1'> <u>快捷郵件</u> </p>";
                        GridList.Text += "    </td>";
                        GridList.Text += "    <td width='30%' height='120' align='center'>";
                        GridList.Text += "      <p><img src='../images/post.gif' width='82' height='80'></p>";
                        GridList.Text += "      <p class='style3'>郵局郵戳</p>";
                        GridList.Text += "    </td>";
                        GridList.Text += "  </tr>";
                        GridList.Text += "  <tr>";
                        GridList.Text += "    <td width='21%' height='25'><span class='style3'>寄件人代表：" + dr_Dept["Contactor"].ToString() + "</span></td>";
                        GridList.Text += "    <td width='48%' height='25' colspan='2'><span class='style3'>詳細地址：235" + dr_Dept["Address"].ToString() + "</span></td>";
                        GridList.Text += "    <td width='30%' height='25'><span class='style3'>電話號碼：" + dr_Dept["TEL"].ToString() + "</span></td>";
                        GridList.Text += "  </tr>";
                        GridList.Text += "</table>";
                        GridList.Text += "<table width='720'  border='1' cellspacing='0' bordercolor='#000000' style='border-collapse: collapse'>";
                        GridList.Text += "  <tr>";
                        GridList.Text += "    <td width='5%' height='25' rowspan='2'><div align='center' class='style3'>順序號碼</div></td>";
                        GridList.Text += "    <td width='10%' height='25' rowspan='2'><div align='center' class='style3'>掛號號碼</div></td>";
                        GridList.Text += "    <td height='25' colspan='2'><div align='center' class='style3'>收件人</div></td>";
                        GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否回執(V)</div></td>";
                        GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否航空(V)</div></td>";
                        GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>是否印刷物(V)</div></td>";
                        GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>重量</div></td>";
                        GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>郵資</div></td>";
                        GridList.Text += "    <td width='6%' height='25' rowspan='2'><div align='center' class='style3'>備考</div></td>";
                        GridList.Text += "  </tr>";
                        GridList.Text += "  <tr>";
                        GridList.Text += "    <td width='10%' height='25'><div align='center' class='style3'>姓名</div></td>";
                        GridList.Text += "    <td width='38%' height='25'><p align='center' class='style3'>寄達第名(或地址)</p>    </td>";
                        GridList.Text += "  </tr>";
                    }
                    LastRow = 0;
                }
                else
                {
                    LastRow += 1;
                }
                Row += 1;
                if (Row - 1 == dt_Donor.Rows.Count)
                {
                    end = true;
                }
            }
            //Row = Row-1;
            if ((Row + 20) % 20 != 0)
            {
                while ((Row + 20) % 20 != 0)
                {
                    GridList.Text += "<tr>";
                    GridList.Text += "  <td width='5%' height='30'><div align='center' class='style3'> </div></td>";
                    GridList.Text += "  <td width='10%' height='30'><span class='style3'></span></td>";
                    GridList.Text += "  <td width='10%' height='30'><span class='style3'></span></td>";
                    GridList.Text += "  <td width='38%' height='30'><span class='style3'></span></td>";
                    GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                    GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                    GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                    GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                    GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                    GridList.Text += "  <td width='6%' height='30'><span class='style3'></span></td>";
                    GridList.Text += "</tr>";
                    Row += 1;
                }
                GridList.Text += "</table>";
                GridList.Text += "<table width='720'  border='0' cellspacing='10'>";
                GridList.Text += "  <tr>";
                GridList.Text += "    <td width='58%' valign='top'>";
                GridList.Text += "      <span class='style3'>";
                GridList.Text += "        限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br />";
                GridList.Text += "	      函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br />";
                GridList.Text += "	      將本埠與外埠函件分別列單交寄。<br />";
                GridList.Text += "	      此單由郵局免費供給，應由寄件人清晰填寫一式二份。<br />";
                GridList.Text += "	      如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局經辦員逐件核對。<br />";
                GridList.Text += "	      日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br />";
                GridList.Text += "	      錢鈔或有價證券請利用報值或保價交寄。<br />";
                GridList.Text += "	    </span>";
                GridList.Text += "	  </td>";
                GridList.Text += "    <td width='38%'>";
                GridList.Text += "      <p class='style3'>上開　限時掛號 </p>";
                GridList.Text += "      <p class='style3'>掛號函件 / 共　　" + LastRow + "　　件 照收無誤　快捷郵件 </p>";
                GridList.Text += "	    <p class='style3'>　</p>";
                GridList.Text += "      <p class='style3'><font size='3'>郵資共計　　　　&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 元</font> </p>";
                GridList.Text += "      <p align='center' class='style3'>經辦員簽署</p>";
                GridList.Text += "    </td>";
                GridList.Text += "  </tr>";
                GridList.Text += "</table>";
            }
        }
        else
        {
            GridList.Text = "** 無符合的資料 **";
        }
    }
}